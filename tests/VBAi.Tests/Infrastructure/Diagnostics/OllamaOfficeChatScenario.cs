using System;
using System.Linq;

namespace VBAi.Tests.Integration
{
    /// <summary>The same genuine assistant workflow is used by the in-process and isolated observers.</summary>
    internal static class OllamaOfficeChatScenario
    {
        // UiText.Apply localizes string ComboBox items after Designer construction.
        internal static string ReadOnlyApprovalChoice => UiText.Get("Read-only");
        /// <summary>Requests the original numbered streaming workload within the assistant's actual VBA context.</summary>
        internal static string StreamingRequest()
            => "This is a synthetic VBAi display qualification. Produce a numbered catalogue of 1000 fictional VBA object labels as text, " +
                "using exactly '1. Object 1', '2. Object 2', '3. Object 3', and so on. " +
                "Do not inspect, modify or run a project and do not use tools. Start with the first numbered line, without an introduction or analysis.";

        /// <summary>Requests the unchanged recovery acknowledgment for this VBAi conversation.</summary>
        internal static string RecoveryRequest()
            => "The previous VBAi display reply was intentionally stopped. Confirm that this VBA assistant conversation can respond again: " +
                "reply with exactly UI_READY_42 and nothing else. Do not use tools.";

        internal static string MarkerRequest(string project, string module, string marker)
        {
            if (string.IsNullOrWhiteSpace(project) || module != "Q028Marker" || string.IsNullOrEmpty(marker))
                throw new ArgumentException("An owned disposable marker module is required.");
            string request = "This is a synthetic qualification. Inspect only the selected disposable project. Call read_module exactly once for Project=" + project +
                ", Module=" + module + ". What is the exact value of the string constant ObservedMarker? The value is only in the module. Do not modify or execute code. Return the value after reading it.";
            if (request.Contains(marker)) throw new InvalidOperationException("The native marker must not be supplied in the prompt.");
            return request;
        }

        internal static void Run(OllamaOfficeUi ui, string label, string project, string module, string marker,
            string model, Action<string> passed)
        {
            ui.Discover(); ui.DetectNativeLanguage();
            ui.Select("scopePicker", label); ui.RequireScope(label);
            ui.Click("modelSummary");
            ui.Select("providerPicker", LlmProvider.All.Single(item => item.IsOllama).ToString());
            ui.Wait(() => ui.Leaf("modelPicker").Cached.IsEnabled, 30, "model catalogue ready");
            ui.Select("modelPicker", model);
            string reasoning = OllamaQualificationProfile.Resolve().ReasoningEffort;
            if (reasoning != null) ui.Select("effortPicker", ReasoningLabel(reasoning));
            // Keep the real controls available for terminal enabled-state readback.
            ui.Select("modePicker", UiText.Get("Chat"));
            ui.Select("approvalPicker", ReadOnlyApprovalChoice);
            ui.BindTranscript();
            ui.SendOnce(StreamingRequest());
            ui.Wait(() => ui.IsBusy && OllamaOfficeStreamOracle.HasNumberedResponse(ui.VisibleTranscript()),
                125, "visible streamed text while busy");
            passed("StreamedWhileBusy");
            ui.StopOnce(); ui.WaitIdle(30);
            ui.Wait(() => ui.VisibleTranscript().Any(text => text.Contains(UiText.Get("Response interrupted. Changes already applied can still be undone in the chat."))),
                10, "visible cancellation acknowledgement");
            passed("CancellationAcknowledged");
            ui.SendOnce(RecoveryRequest());
            ui.WaitIdle(125);
            ui.Wait(() => ui.VisibleTranscript().Any(OllamaOfficeStreamOracle.IsReadyResponse), 10, "visible completed next reply");
            passed("NextSendCompleted");
            ui.SendOnce(MarkerRequest(project, module, marker)); ui.WaitIdle(125);
            ui.Wait(() => ui.VisibleTranscript().Any(text => text.Contains(marker)), 15, "visible unprompted native marker");
            passed("NativeMarkerVisible");
        }

        /// <summary>Uses the actual localized choice for the declared profile; does not silently accept another effort.</summary>
        internal static string ReasoningLabel(string effort)
        {
            if (effort == OllamaReasoning.None) return UiText.Get("Reasoning off");
            if (effort == OllamaReasoning.Default) return UiText.Get("Model default");
            throw new InvalidOperationException("The native reasoning profile must be default or none.");
        }
    }
}
