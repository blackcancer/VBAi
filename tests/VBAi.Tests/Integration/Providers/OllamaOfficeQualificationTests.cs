using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    /// <summary>Real local model, installed embedded assistant and disposable Office project; never executes VBA.</summary>
    [TestClass, TestCategory("OllamaOffice"), DoNotParallelize]
    public sealed class OllamaOfficeQualificationTests
    {
        public TestContext TestContext { get; set; }
        [STATestMethod] public void ExcelEmbeddedAssistantStreamsStopsRecoversAndReadsNativeMarker() { Run("Excel"); }
        [STATestMethod] public void WordEmbeddedAssistantStreamsStopsRecoversAndReadsNativeMarker() { Run("Word"); }
        [STATestMethod] public void PowerPointEmbeddedAssistantStreamsStopsRecoversAndReadsNativeMarker() { Run("PowerPoint"); }
        [STATestMethod] public void AccessEmbeddedAssistantStreamsStopsRecoversAndReadsNativeMarker() { Run("Access"); }
        [STATestMethod] public void PublisherEmbeddedAssistantStreamsStopsRecoversAndReadsNativeMarker() { Run("Publisher"); }
        [STATestMethod] public void OutlookEmbeddedAssistantStreamsStopsRecoversAndReadsNativeMarker() { Run("Outlook"); }

        private void Run(string kind)
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_OLLAMA_OFFICE_TESTS") != "1")
                Assert.Inconclusive("Q028 Office assistant requires its explicit opt-in and frozen private-desktop campaign.");
            string root = ExcelOwnedBootstrapPlan.RequireLocalAbsolutePath(Environment.GetEnvironmentVariable("VBAi_Q028_RESULTS"));
            Directory.CreateDirectory(root);
            string path = Path.Combine(root, kind + "-assistant.json");
            Assert.IsFalse(File.Exists(path), "A host bank is one-shot; preserve prior evidence instead of replaying it.");
            var profile = OllamaQualificationProfile.Resolve();
            var languageScope = new VBAi.Tests.Infrastructure.LocalizationScope(UiText.Culture.Name);
            var observations = new List<object>();
            Action flush = null;
            Action<object> record = value => { lock (observations) { observations.Add(value); flush?.Invoke(); } };
            var report = new Dictionary<string, object> { ["State"] = "STARTED_ONCE", ["Host"] = kind,
                ["SourceMvid"] = typeof(VbeSession).Module.ModuleVersionId.ToString("D"), ["Profile"] = profile.Describe(),
                ["MacroExecutionRequested"] = false, ["StartedUtc"] = DateTime.UtcNow.ToString("o"), ["Observations"] = observations };
            flush = () => File.WriteAllText(path, new JavaScriptSerializer { MaxJsonLength = 8 * 1024 * 1024 }.Serialize(report));
            flush();
            OllamaOfficeHost host = null;
            OllamaOfficeUi ui = null;
            ExceptionDispatchInfo failure = null;
            try
            {
                host = OllamaOfficeHost.Start(kind, assigned => {
                    host = assigned; report["ProcessId"] = assigned.ProcessId;
                    report["NativeEvidenceRoot"] = assigned.Root; flush();
                });
                report["ProcessId"] = host.ProcessId; report["Project"] = host.Project; report["NativeEvidenceRoot"] = host.Root;
                var status = host.Data("status");
                Assert.AreEqual(typeof(VbeSession).Module.ModuleVersionId.ToString("D"), status["AssemblyModuleVersionId"]);
                Assert.AreEqual(host.ProcessId, Convert.ToInt32(status["HostProcessId"]));
                Assert.AreEqual(64, Convert.ToInt32(status["ProcessBitness"])); record(status);
                const string module = "Q028Marker";
                string marker = "NATIVE_" + Guid.NewGuid().ToString("N");
                host.Data("create_module", "Module", module, "ExpectedMode", 2);
                var empty = host.Data("read_module", "Module", module);
                int lines = Convert.ToInt32(host.Data("component_properties", "Module", module)["CodeLines"]);
                host.Data("replace_lines", "Module", module, "StartLine", 1, "Count", lines,
                    "ExpectedSha256", empty["Sha256"], "Text", "Option Explicit\r\nPublic Const ObservedMarker As String = \"" + marker + "\"\r\n");
                host.TrackOutlookModule(module);
                var before = host.Data("read_module", "Module", module);
                host.Data("select_code", "Module", module, "StartLine", 1, "ExpectedSha256", before["Sha256"]);
                string references = Convert.ToString(host.Data("list_references")["Version"]);
                var sources = host.Items("list_modules").ToDictionary(row => Convert.ToString(row["Name"]),
                    row => Convert.ToString(host.Data("read_module", "Module", row["Name"])["Sha256"]));
                report["SourceBeforeSha256"] = before["Sha256"]; report["MarkerSha256"] = EditorDocument.Hash(marker);
                var worker = new Thread(() => {
                    try
                    {
                        ui = new OllamaOfficeUi(host.ProcessId, record); ui.Discover();
                        ui.DetectNativeLanguage();
                        ui.Select("scopePicker", host.Label);
                        ui.RequireScope(host.Label);
                        ui.Click("modelSummary");
                        ui.Select("providerPicker", LlmProvider.All.Single(item => item.IsOllama).ToString());
                        ui.Wait(() => ui.Leaf("modelPicker").Current.IsEnabled, 30, "model catalogue ready");
                        ui.Select("modelPicker", profile.Model);
                        // Keep the genuine provider controls visible for terminal enabled-state
                        // readback; the collapsed layout intentionally removes them from UIA.
                        ui.Select("modePicker", UiText.Get("Chat"));
                        ui.Select("approvalPicker", UiText.Get("Read-only"));
                        ui.BindTranscript();
                        ui.SendOnce("This is a synthetic interface test unrelated to VBA. Do not use tools. Output 1000 lines with exactly this format: '1. Object 1', then '2. Object 2', then '3. Object 3', continuing the same pattern. Begin immediately with the first line; no introductory remarks.");
                        ui.Wait(() => ui.IsBusy && ui.VisibleTranscript().Any(OllamaOfficeStreamOracle.IsNumberedResponse),
                            125, "visible streamed text while busy");
                        report["StreamedWhileBusy"] = true; flush();
                        ui.StopOnce(); ui.WaitIdle(30);
                        ui.Wait(() => ui.VisibleTranscript().Any(text => text.Contains(UiText.Get("Response interrupted. Changes already applied can still be undone in the chat."))),
                            10, "visible cancellation acknowledgement");
                        report["CancellationAcknowledged"] = true; flush();
                        ui.SendOnce("The previous response was intentionally stopped. This is a synthetic UI test unrelated to VBA. Do not use tools. Reply with exactly UI_READY_42 and nothing else.");
                        ui.WaitIdle(125);
                        ui.Wait(() => ui.VisibleTranscript().Any(OllamaOfficeStreamOracle.IsReadyResponse), 10, "visible completed next reply");
                        report["NextSendCompleted"] = true; flush();
                        string request = "This is a synthetic qualification. Inspect only the selected disposable project. Call read_module exactly once for Project=" + host.Project +
                            ", Module=" + module + ". What is the exact value of the string constant ObservedMarker? The value is only in the module. Do not modify or execute code. Return the value after reading it.";
                        Assert.IsFalse(request.Contains(marker));
                        ui.SendOnce(request); ui.WaitIdle(125);
                        ui.Wait(() => ui.VisibleTranscript().Any(text => text.Contains(marker)), 15, "visible unprompted native marker");
                        report["NativeMarkerVisible"] = true; flush();
                    }
                    catch (Exception error)
                    {
                        failure = ExceptionDispatchInfo.Capture(error);
                        try { if (ui != null) record(new { Phase = "FailureTerminalObservation", AlreadyIdle = ui.ObserveAlreadyIdle() }); }
                        catch (Exception observation) { record(new { Phase = "TerminalObservationFailed", Error = observation.ToString() }); }
                    }
                });
                worker.SetApartmentState(ApartmentState.MTA); worker.Start();
                // Pump the owning STA while the independent native UI observer runs. Do not dispatch
                // another native command or teardown until its terminal receipt is available.
                while (!worker.Join(20)) System.Windows.Forms.Application.DoEvents();
                failure?.Throw();
                Assert.AreEqual(before["Sha256"], host.Data("read_module", "Module", module)["Sha256"]);
                Assert.AreEqual(before["Code"], host.Data("read_module", "Module", module)["Code"]);
                Assert.AreEqual(references, host.Data("list_references")["Version"]);
                Assert.AreEqual(sources.Count, host.Items("list_modules").Length);
                foreach (var source in sources) Assert.AreEqual(source.Value, host.Data("read_module", "Module", source.Key)["Sha256"]);
                report["SourceAndReferencesUnchanged"] = true;
                report["State"] = "PASS";
            }
            catch (Exception error)
            {
                if (failure == null) failure = ExceptionDispatchInfo.Capture(error);
                report["State"] = error is AssertInconclusiveException ? "BLOCKED" : "FAIL";
                report["Error"] = error.ToString();
            }
            finally
            {
                if ((ui != null && ui.SentUnsettled) || (host != null && (host.NativeDispatchUnsettled || host.IsRetained)))
                {
                    report["HostRetained"] = true;
                    report["State"] = "FAILED_OR_UNCERTAIN";
                    // No second Send/Stop or native mutation to turn an uncertain result into a pass.
                    host?.Retain();
                }
                else if (host != null)
                {
                    try { host.Dispose(); report["FixtureShutdownVerified"] = true; }
                    catch (Exception cleanup)
                    {
                        report["State"] = "FAIL"; report["CleanupError"] = cleanup.ToString();
                        if (failure == null) failure = ExceptionDispatchInfo.Capture(cleanup);
                    }
                }
                languageScope.Dispose();
                report["CompletedUtc"] = DateTime.UtcNow.ToString("o"); flush(); TestContext.AddResultFile(path);
            }
            failure?.Throw();
        }
    }
}
