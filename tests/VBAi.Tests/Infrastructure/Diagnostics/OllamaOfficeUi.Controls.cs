namespace VBAi.Tests.Integration
{
    internal sealed partial class OllamaOfficeUi
    {
        /// <summary>Maps scenario roles to the exact shared Designer identities, never captions.</summary>
        internal static string AssistantControlId(string role)
        {
            switch (role)
            {
                case "options": return "settingsButton";
                case "send": return "sendButton";
                case "stop": return "stopButton";
                case "promptHost": return "inputTextBox";
                case "transcriptPanel": return "TranscriptHost";
                default: return role;
            }
        }

        /// <summary>Only the initially unbound Stop may have no HWND while WinForms defers its creation.</summary>
        internal static bool InitialStopCanBeAbsent(string role, int peerCount)
            => role == "stop" && peerCount == 0;

        /// <summary>Fresh states from the original send/Stop HWNDs and optional terminal selectors.</summary>
        internal sealed class AssistantActionState
        {
            internal bool SendVisible, SendEnabled, StopVisible, ProviderEnabled, ModeEnabled;
        }

        /// <summary>Idle send must be visible and enabled; drafting during a live turn never qualifies it.</summary>
        internal static bool CanSend(AssistantActionState state)
            => state != null && !state.StopVisible && state.SendVisible && state.SendEnabled;

        /// <summary>A terminal turn requires the independent Stop to disappear and idle selectors to recover.</summary>
        internal static bool CanFinish(AssistantActionState state)
            => CanSend(state) && state.ProviderEnabled && state.ModeEnabled;
    }
}
