using System;
using System.Collections.Generic;

namespace VBAi.Tests.Integration
{
    internal sealed partial class OutlookVbaTestFixture
    {
        /// <summary>Observes the retained process after a failed close without another native action or a success verdict.</summary>
        /// <param name="stage">The single close or quit operation that failed.</param>
        /// <param name="cause">The original error, retained even when process exit is subsequently observed.</param>
        /// <param name="failures">The original disposal failures; diagnostic errors are added without replacing them.</param>
        internal void ObserveFailedShutdown(string stage, Exception cause, List<string> failures)
        {
            var observation = new Dictionary<string, object>
            {
                ["Stage"] = stage,
                ["ProcessId"] = ProcessId,
                ["OriginalErrorType"] = cause.GetType().FullName,
                ["OriginalErrorHResult"] = cause.HResult,
                ["FailureRetained"] = true,
                ["NativeActionsRetried"] = false,
                ["ComReferencesReleased"] = false,
                ["OtmRecoveryAttempted"] = false,
                ["NormalQuitQualified"] = false,
                ["ExitObserved"] = false,
                ["ExitCodeObserved"] = false,
                ["StartedUtc"] = DateTime.UtcNow.ToString("o")
            };
            try
            {
                if (process == null || process.Id != ProcessId)
                    throw new InvalidOperationException("The retained Outlook process identity is unavailable.");
                observation["OriginalProcessHandle"] = process.Handle.ToInt64();
                observation["HandleSource"] = "FirstObservedOwnedProcessHandle";
                bool exited = WaitForOwnedShutdownExit(15000);
                if (exited && !process.HasExited)
                    throw new InvalidOperationException("The retained process handle does not confirm the reported exit.");
                observation["ExitObserved"] = exited;
                if (exited)
                {
                    observation["ExitCode"] = process.ExitCode;
                    observation["ExitCodeObserved"] = true;
                }
            }
            catch (Exception error)
            {
                observation["ObservationErrorType"] = error.GetType().FullName;
                observation["ObservationError"] = error.Message;
                failures.Add("Failed-shutdown observation: " + error.Message);
            }
            observation["CompletedUtc"] = DateTime.UtcNow.ToString("o");
            try { Save("shutdown-failure-observation.json", observation); }
            catch (Exception error) { failures.Add("Failed-shutdown observation receipt: " + error.Message); }
        }
    }
}
