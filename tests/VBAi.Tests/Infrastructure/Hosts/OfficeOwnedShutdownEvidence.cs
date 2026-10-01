using System;
using System.Collections.Generic;
using System.Runtime.ExceptionServices;

namespace VBAi.Tests.Integration
{
    /// <summary>Records one owned host generation's shutdown without disposing its identity handle when exit remains unverified.</summary>
    internal sealed class OfficeOwnedShutdownEvidence
    {
        internal readonly IDictionary<string, object> Record;
        private bool prepared;
        private bool preparationDurable;
        private bool quitEntered;
        private bool quitTerminalDurable;
        private bool exitObservationEntered;

        internal OfficeOwnedShutdownEvidence(int pid, string startedUtc, string image, string handle, string mvid)
        {
            if (pid <= 0 || string.IsNullOrWhiteSpace(startedUtc) || string.IsNullOrWhiteSpace(image) || string.IsNullOrWhiteSpace(handle))
                throw new ArgumentException("Exact original PID/start/image/retained-handle identity is required.");
            Record = new Dictionary<string, object> {
                ["ProcessId"] = pid, ["ProcessStartedUtc"] = startedUtc, ["ProcessImage"] = image,
                ["OriginalProcessHandle"] = handle, ["ExpectedAssemblyMvid"] = mvid,
                ["State"] = "OWNED", ["TeardownPrepared"] = false, ["QuitEntries"] = 0,
                ["QuitOutcome"] = "NOT_ENTERED", ["ProcessExitObserved"] = false,
                ["ExitCodeObserved"] = false, ["ExitCode"] = null, ["OwnershipRetained"] = true,
                ["ProcessHandleRetained"] = true, ["ForcedTermination"] = false, ["NativeCleanupReplayAllowed"] = false };
            Record["ExitObservationAttempts"] = 0;
        }

        internal void Prepare(Action persist)
        {
            if (prepared) throw new InvalidOperationException("This host generation already has a shutdown attempt; no Close/Quit replay is permitted.");
            prepared = true; Record["TeardownPrepared"] = true; Record["State"] = "TEARDOWN_PREPARED";
            persist(); // Preparation failure emits no native shutdown request.
            preparationDurable = true;
        }

        internal void QuitOnce(Action quit, Action persist)
        {
            if (!prepared || !preparationDurable || quitEntered) throw new InvalidOperationException("Quit requires one durably prepared, unreplayed shutdown attempt.");
            quitEntered = true; Record["QuitOutcome"] = "PREPARED"; Record["State"] = "QUIT_PREPARED";
            persist();
            Exception failure = null;
            Record["QuitEntries"] = 1; Record["QuitOutcome"] = "CLIENT_CALL_ENTERED";
            try
            {
                quit();
                Record["QuitOutcome"] = "RETURNED"; Record["QuitReturnedUtc"] = DateTime.UtcNow.ToString("o");
                Record["State"] = "QUIT_RETURNED_EXIT_NOT_YET_OBSERVED";
            }
            catch (Exception error)
            {
                failure = error; Record["QuitOutcome"] = "CALL_FAILED_EFFECT_UNKNOWN"; Record["State"] = "RETAINED_QUIT_FAILURE";
                Record["OriginalQuitError"] = error.ToString();
                Record["OriginalQuitHResult"] = "0x" + unchecked((uint)error.HResult).ToString("X8");
            }
            try { persist(); quitTerminalDurable = failure == null; }
            catch (Exception evidence) { failure = failure == null ? evidence : new AggregateException("The original Quit failure and its evidence write both failed.", failure, evidence); }
            if (failure != null) ExceptionDispatchInfo.Capture(failure).Throw();
        }

        /// <summary>Returns true only after terminal Quit, observed exit/code and handle disposal; otherwise retains original ownership.</summary>
        internal bool ObserveExit(Func<bool> wait, Func<int> exitCode, Action disposeHandle, Action persist)
        {
            if (!quitTerminalDurable || !Equals(Record["QuitOutcome"], "RETURNED") || exitObservationEntered)
                throw new InvalidOperationException("Exit acceptance requires the original Quit call to have returned; unknown outcomes retain ownership.");
            exitObservationEntered = true; Record["ExitObservationAttempts"] = 1;
            try
            {
                bool exited = wait();
                Record["ProcessExitObserved"] = exited;
                Record["ExitObservedUtc"] = DateTime.UtcNow.ToString("o");
                if (!exited)
                {
                    Record["State"] = "RETAINED_EXIT_NOT_OBSERVED";
                    persist();
                    return false; // Never release the original process handle after a timeout.
                }
                Record["ExitCode"] = exitCode(); Record["ExitCodeObserved"] = true;
                Record["State"] = "EXIT_OBSERVED_HANDLE_STILL_RETAINED"; persist();
                disposeHandle();
                Record["ProcessHandleRetained"] = false; Record["OwnershipRetained"] = false;
                Record["State"] = "EXIT_OBSERVED_HANDLE_RELEASED"; persist();
                return true;
            }
            catch (Exception error)
            {
                // A disposal failure can make handle state uncertain; never claim
                // disposal or normal exit that was not actually observed.
                Record["ExitObservationError"] = error.ToString(); Record["State"] = "EXIT_OBSERVATION_OR_EVIDENCE_FAILED";
                try { persist(); }
                catch (Exception evidence) { throw new AggregateException("Original exit observation/cleanup and its evidence write both failed.", error, evidence); }
                throw;
            }
        }
    }
}
