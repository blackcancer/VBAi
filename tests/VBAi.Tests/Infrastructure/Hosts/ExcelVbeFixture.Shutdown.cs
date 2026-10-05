using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;

namespace VBAi.Tests.Integration
{
    /// <summary>Retains failed original-process observation without rearming any native cleanup.</summary>
    internal sealed partial class ExcelVbeFixture
    {
        private int shutdownState;
        private bool processReleaseEntered, processReleaseReturned;
        private bool exitWaitAttempted, exitWaitReturned, exitCodeObserved;
        private string cleanupStage = "CLEANUP_BEFORE_WAIT";
        private readonly List<Exception> shutdownFailures = new List<Exception>();
        /// <summary>Optional synthetic receipt writer; no live host or file service is needed in mirrors.</summary>
        internal Action<IDictionary<string, object>> WriteShutdownReceipt;
        /// <summary>Collects only after the scenario frame returns; mirrors record the order without invoking the GC.</summary>
        internal Action CollectScenarioReferences = () => {
            GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect(); GC.WaitForPendingFinalizers();
        };
        private readonly int shutdownOwnerThread = Thread.CurrentThread.ManagedThreadId;
        /// <summary>One original-handle exit observation; synthetic mirrors replace it without starting Office.</summary>
        internal Func<Process, int, bool> WaitForOwnedExcelExit = (process, timeout) => process.WaitForExit(timeout);
        /// <summary>Reads exit status only after the original wait succeeds.</summary>
        internal Func<Process, int> ReadOwnedExcelExitCode = process => process.ExitCode;
        /// <summary>Releases only a successfully observed original process descriptor.</summary>
        internal Action<Process> ReleaseOwnedExcelProcess = process => process.Dispose();

        private bool ClaimShutdown()
        {
            if (Thread.CurrentThread.ManagedThreadId != shutdownOwnerThread)
                throw new InvalidOperationException("Excel cleanup must stay on its original owner thread.");
            int previous = Interlocked.CompareExchange(ref shutdownState, 1, 0);
            if (previous == 2) return false;
            if (previous != 0) throw new InvalidOperationException("Excel cleanup was already entered or retained; no Close/Quit, drain or exit-wait replay is allowed.");
            return true;
        }

        private void CompleteShutdown() { Interlocked.Exchange(ref shutdownState, 2); }

        private void RememberShutdownFailure(Exception error)
        {
            if (error != null && !shutdownFailures.Any(known => ReferenceEquals(known, error))) shutdownFailures.Add(error);
        }

        private Exception CombineShutdownFailure(Exception error)
        {
            var aggregate = error as AggregateException;
            if (aggregate == null) RememberShutdownFailure(error);
            else {
                var flattened = aggregate.Flatten();
                foreach (Exception inner in flattened.InnerExceptions) RememberShutdownFailure(inner);
                if (shutdownFailures.All(known => flattened.InnerExceptions.Any(inner => ReferenceEquals(inner, known)))) return error;
            }
            return shutdownFailures.Count == 1 ? shutdownFailures[0] :
                new AggregateException("Excel cleanup failures are all retained; no cleanup can be replayed.", shutdownFailures);
        }

        private void RetainFailedShutdown(Exception error)
        {
            Interlocked.Exchange(ref shutdownState, 3);
            // Keep the same descriptor and generation alive until this testhost exits.
            // A future diagnostic observer must use this lease; it must never reopen by PID.
            lock (retainedBootstraps)
                if (!retainedBootstraps.Contains(this)) retainedBootstraps.Add(this);
            if (ShutdownDiagnostics != null)
            {
                ShutdownDiagnostics["CleanupState"] = "FAILED_NO_REPLAY";
                ShutdownDiagnostics["ProcessHandleRetained"] = ownedProcess == null ? (object)false :
                    processReleaseEntered && !processReleaseReturned ? null : (object)true;
                ShutdownDiagnostics["ProcessHandleReleaseOutcome"] = processReleaseReturned ? "RETURNED" :
                    processReleaseEntered ? "UNKNOWN_AFTER_ENTRY" : "NOT_ENTERED";
                ShutdownDiagnostics["OriginalCleanupVerdictFinal"] = true;
                ShutdownDiagnostics["ExitWaitAttempted"] = exitWaitAttempted;
                ShutdownDiagnostics["ExitWaitReturned"] = exitWaitReturned;
                ShutdownDiagnostics["ExitCodeObserved"] = exitCodeObserved;
                ShutdownDiagnostics["FailureStage"] = cleanupStage;
                ShutdownDiagnostics["CleanupFailure"] = error.ToString();
                ShutdownDiagnostics["LaterObservationCanQualify"] = false;
                try { WriteShutdownDiagnostics(ShutdownDiagnostics); }
                catch { /* Preserve the original cleanup/evidence error; no cleanup follows. */ }
            }
        }

        private static void RecordShutdownThread(IDictionary<string, object> diagnostics, string phase)
        {
            diagnostics[phase + "ManagedThreadId"] = Thread.CurrentThread.ManagedThreadId;
            diagnostics[phase + "NativeThreadId"] = IsolatedTestDesktop.GetCurrentThreadId();
            diagnostics[phase + "Apartment"] = Thread.CurrentThread.GetApartmentState().ToString();
        }
    }
}
