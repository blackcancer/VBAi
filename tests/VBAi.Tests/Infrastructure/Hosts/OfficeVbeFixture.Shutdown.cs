using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace VBAi.Tests.Integration
{
    /// <summary>Retains the exact original process handle and a durable one-shot Quit/exit lifecycle on unsuccessful shutdown.</summary>
    internal sealed partial class OfficeVbeFixture
    {
        private OfficeOwnedShutdownEvidence shutdownEvidence;
        // Synthetic mirror tests replace only exit observation; they never launch
        // or terminate a host or change production COM behavior.
        private static readonly Func<Process, int, bool> DefaultOwnedExitWait = (process, timeout) => process.WaitForExit(timeout);
        internal Func<Process, int, bool> WaitForOwnedExit = DefaultOwnedExitWait;
        /// <summary>Bounds one read-only original-handle observation after the sole Quit call.</summary>
        internal int OwnedExitWaitMilliseconds = 5000;
        internal Action PumpWordShutdownMessages = Application.DoEvents;
        internal Func<Process, bool> ReadWordProcessExit = process => process.WaitForExit(0);
        private readonly int shutdownOwnerThread = Thread.CurrentThread.ManagedThreadId;
        internal Func<Process, int> ReadOwnedExitCode = process => process.ExitCode;

        private void RequireUsableOwnedHost()
        {
            commandContainment.RequireTerminal();
            if (hostTeardownRefused || (shutdownEvidence != null && (!owned || (Kind == "Word" && Equals(shutdownEvidence.Record["TeardownPrepared"], true)))))
                throw new InvalidOperationException("The original Office ownership scope is closed or unverified; no native request, Save, Close/Quit retry or reopen is permitted.");
            RequirePrivateHostDesktop(true);
        }

        private bool WaitForShutdownExit(Process process, int timeout, bool externalReferencesReleased)
        {
            if (Kind != "Word" || !ReferenceEquals(WaitForOwnedExit, DefaultOwnedExitWait) || !externalReferencesReleased)
                return WaitForOwnedExit(process, timeout);
            Assert.AreEqual(shutdownOwnerThread, Thread.CurrentThread.ManagedThreadId, "Word shutdown observation must stay on its owning thread.");
            Assert.AreSame(ownedProcess, process, "Only the originally retained Word process may be observed.");
            Assert.AreEqual("RETURNED", shutdownEvidence.Record["QuitOutcome"]);
            Assert.IsNull(application); Assert.IsNull(document);
            string originalHandle = "0x" + unchecked((ulong)process.Handle.ToInt64()).ToString("X16");
            Assert.AreEqual(shutdownEvidence.Record["OriginalProcessHandle"], originalHandle);
            var clock = Stopwatch.StartNew();
            int pumps = 0;
            shutdownEvidence.Record["WordMessagePumpEnabled"] = true;
            shutdownEvidence.Record["ExitObservationThread"] = shutdownOwnerThread;
            try
            {
                return WaitForWordExit(timeout, () => clock.ElapsedMilliseconds,
                    () => ReadWordProcessExit(process), () => { pumps++; PumpWordShutdownMessages(); }, Thread.Sleep);
            }
            finally
            {
                shutdownEvidence.Record["ExitWaitElapsedMilliseconds"] = clock.ElapsedMilliseconds;
                shutdownEvidence.Record["WordMessagePumpAttempts"] = pumps;
                shutdownEvidence.Record["ExitWaitProcessHandle"] = originalHandle;
            }
        }

        /// <summary>Pumps only during one known shutdown observation; it never dispatches COM or repeats Close/Quit.</summary>
        internal static bool WaitForWordExit(int timeout, Func<long> elapsed, Func<bool> exited, Action pump, Action<int> pause)
        {
            while (true)
            {
                if (elapsed() > timeout) return false;
                bool observedExit = exited();
                if (elapsed() > timeout) return false;
                if (observedExit) return true;
                long remaining = timeout - elapsed();
                if (remaining <= 0) return false;
                pump();
                if (elapsed() > timeout) return false;
                observedExit = exited();
                if (elapsed() > timeout) return false;
                if (observedExit) return true;
                remaining = timeout - elapsed();
                if (remaining <= 0) return false;
                pause((int)Math.Min(25, remaining));
            }
        }
        internal Action CollectSettledWordReferences = () =>
        {
            GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect();
        };

        /// <summary>Optional testhost-only experiment after a successful non-inlined Word qualification scope.</summary>
        internal void CollectSettledWordScopeDiagnostic()
        {
            if (Kind != "Word" || Environment.GetEnvironmentVariable("VBAi_TEST_WORD_SETTLED_SCOPE_GC") != "1") return;
            RequireUsableOwnedHost();
            Assert.IsFalse(NativeExecutionUnsettled, "Uncertain or pending Word execution must not be finalized by this diagnostic.");
            Assert.AreEqual(shutdownOwnerThread, Thread.CurrentThread.ManagedThreadId);
            Assert.IsTrue(owned); Assert.IsNotNull(ownedProcess); Assert.IsNotNull(shutdownEvidence);
            Assert.AreEqual(ProcessId, ownedProcess.Id);
            Assert.IsFalse(ownedProcess.HasExited, "The original owned Word process must still be alive before the diagnostic.");
            Assert.AreEqual(shutdownEvidence.Record["ProcessStartedUtc"], ownedProcess.StartTime.ToUniversalTime().ToString("o"));
            string handle = "0x" + unchecked((ulong)ownedProcess.Handle.ToInt64()).ToString("X16");
            Assert.AreEqual(shutdownEvidence.Record["OriginalProcessHandle"], handle);
            Assert.AreEqual(shutdownEvidence.Record["ProcessImage"], ExcelOwnedProcessImage.Read(ownedProcess.Handle));
            var diagnostic = new System.Collections.Generic.Dictionary<string, object>
            {
                ["DiagnosticOnly"] = true,
                ["ProcessId"] = ProcessId,
                ["OriginalProcessHandle"] = handle,
                ["ScopeReturnedUtc"] = DateTime.UtcNow.ToString("o"),
                ["NativeExecutionUnsettled"] = false,
                ["BridgePending"] = commandContainment.Pending,
                ["BridgeUncertain"] = commandContainment.Uncertain,
                ["CollectionCountsBefore"] = new[] { GC.CollectionCount(0), GC.CollectionCount(1), GC.CollectionCount(2) },
                ["CollectStartedUtc"] = DateTime.UtcNow.ToString("o"),
                ["CollectCompleted"] = false
            };
            string path = Path.Combine(Root, "word-settled-scope-gc.json");
            var serializer = new JavaScriptSerializer();
            File.WriteAllText(path, serializer.Serialize(diagnostic));
            CollectSettledWordReferences();
            diagnostic["CollectCompleted"] = true;
            diagnostic["CollectFinishedUtc"] = DateTime.UtcNow.ToString("o");
            diagnostic["CollectionCountsAfter"] = new[] { GC.CollectionCount(0), GC.CollectionCount(1), GC.CollectionCount(2) };
            File.WriteAllText(path, serializer.Serialize(diagnostic));
        }
        private int ResolveOwnedExitWaitBound()
        {
            if (Kind != "Word") return OwnedExitWaitMilliseconds;
            string configured = Environment.GetEnvironmentVariable("VBAi_TEST_WORD_EXIT_WAIT_BOUND_MS");
            if (string.IsNullOrEmpty(configured)) return OwnedExitWaitMilliseconds;
            if (configured != "15000")
                throw new InvalidOperationException("The Word exit-observation diagnostic accepts only VBAi_TEST_WORD_EXIT_WAIT_BOUND_MS=15000; unset it for the default 5000 milliseconds.");
            return 15000;
        }
        private void PrepareOwnedShutdown()
        {
            Assert.IsNotNull(ownedProcess, "The original owned process handle must be retained before Close/Quit.");
            Assert.AreEqual(ProcessId, ownedProcess.Id);
            Assert.IsNotNull(shutdownEvidence, "Original process identity must have been captured before shutdown.");
            Assert.AreEqual(shutdownEvidence.Record["ProcessStartedUtc"], ownedProcess.StartTime.ToUniversalTime().ToString("o"));
            Assert.AreEqual(shutdownEvidence.Record["OriginalProcessHandle"], "0x" + unchecked((ulong)ownedProcess.Handle.ToInt64()).ToString("X16"));
            Assert.AreEqual(shutdownEvidence.Record["ProcessImage"], ExcelOwnedProcessImage.Read(ownedProcess.Handle));
            shutdownEvidence.Record["WaitBoundMilliseconds"] = ResolveOwnedExitWaitBound();
            shutdownEvidence.Prepare(FlushShutdownEvidence);
        }

        private void QuitOwnedOnce(Action quit) { shutdownEvidence.QuitOnce(quit, FlushShutdownEvidence); }

        private void FlushShutdownEvidence()
        {
            File.WriteAllText(Path.Combine(Root, "shutdown-lifecycle.json"), new JavaScriptSerializer().Serialize(new
            {
                Host = Kind,
                Project,
                DocumentPath,
                Lifecycle = shutdownEvidence.Record
            }));
        }
    }
}
