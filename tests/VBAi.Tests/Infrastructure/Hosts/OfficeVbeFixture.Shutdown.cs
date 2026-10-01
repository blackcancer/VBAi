using System;
using System.Diagnostics;
using System.IO;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    /// <summary>Retains the exact original process handle and a durable one-shot Quit/exit lifecycle on unsuccessful shutdown.</summary>
    internal sealed partial class OfficeVbeFixture
    {
        private OfficeOwnedShutdownEvidence shutdownEvidence;
        // Synthetic mirror tests replace only exit observation; they never launch
        // or terminate a host or change production COM behavior.
        internal Func<Process, int, bool> WaitForOwnedExit = (process, timeout) => process.WaitForExit(timeout);
        /// <summary>Bounds one read-only original-handle observation after the sole Quit call.</summary>
        internal int OwnedExitWaitMilliseconds = 5000;
        internal Func<Process, int> ReadOwnedExitCode = process => process.ExitCode;

        private void RequireUsableOwnedHost()
        {
            commandContainment.RequireTerminal();
            if (hostTeardownRefused || (shutdownEvidence != null && !owned))
                throw new InvalidOperationException("The original Office ownership scope is closed or unverified; no native request, Save, Close/Quit retry or reopen is permitted.");
        }

        private void PrepareOwnedShutdown()
        {
            Assert.IsNotNull(ownedProcess, "The original owned process handle must be retained before Close/Quit.");
            Assert.AreEqual(ProcessId, ownedProcess.Id);
            Assert.IsNotNull(shutdownEvidence, "Original process identity must have been captured before shutdown.");
            Assert.AreEqual(shutdownEvidence.Record["ProcessStartedUtc"], ownedProcess.StartTime.ToUniversalTime().ToString("o"));
            Assert.AreEqual(shutdownEvidence.Record["OriginalProcessHandle"], "0x" + unchecked((ulong)ownedProcess.Handle.ToInt64()).ToString("X16"));
            Assert.AreEqual(shutdownEvidence.Record["ProcessImage"], ExcelOwnedProcessImage.Read(ownedProcess.Handle));
            shutdownEvidence.Record["WaitBoundMilliseconds"] = OwnedExitWaitMilliseconds;
            shutdownEvidence.Prepare(FlushShutdownEvidence);
        }

        private void QuitOwnedOnce(Action quit) { shutdownEvidence.QuitOnce(quit, FlushShutdownEvidence); }

        private void FlushShutdownEvidence()
        {
            File.WriteAllText(Path.Combine(Root, "shutdown-lifecycle.json"), new JavaScriptSerializer().Serialize(new {
                Host = Kind, Project, DocumentPath, Lifecycle = shutdownEvidence.Record }));
        }
    }
}
