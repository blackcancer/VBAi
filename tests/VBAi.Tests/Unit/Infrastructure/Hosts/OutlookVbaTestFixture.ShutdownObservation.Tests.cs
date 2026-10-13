using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    public sealed partial class OutlookVbaTestFixtureShutdownTests
    {
        [TestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void FailedNativeShutdownObservesOnceButKeepsTheOriginalFailureAndOwnedReferences(bool inspectorFails)
        {
            WithFakeFixture((fixture, application, inspector, process, root) =>
            {
                var error = new COMException("Synthetic shutdown RPC failure", unchecked((int)0x800706BE));
                if (inspectorFails) inspector.CloseError = error; else application.QuitError = error;
                int waits = 0;
                fixture.WaitForOwnedShutdownExit = timeout => { waits++; Assert.AreEqual(15000, timeout); return false; };
                Assert.ThrowsException<AssertFailedException>(() => fixture.Dispose());
                AssertRetained(fixture, application, inspector, process, root);
                Assert.AreEqual(1, waits);
                Assert.AreEqual(1, inspector.CloseCount);
                Assert.AreEqual(inspectorFails ? 0 : 1, application.QuitCount);
                var receipt = Read(Path.Combine(root, "shutdown-failure-observation.json"));
                Assert.AreEqual(inspectorFails ? "Inspector.Close" : "Application.Quit", receipt["Stage"]);
                Assert.AreEqual(error.HResult, receipt["OriginalErrorHResult"]);
                Assert.AreEqual(false, receipt["ExitObserved"]);
                Assert.AreEqual(false, receipt["NormalQuitQualified"]);
                Assert.AreEqual(false, receipt["NativeActionsRetried"]);
                Assert.AreEqual(false, receipt["ComReferencesReleased"]);
                Assert.AreEqual(false, receipt["OtmRecoveryAttempted"]);
                fixture.Dispose();
                Assert.AreEqual(1, waits, "A later Dispose must not replay or restart observation.");
            });
        }

        [TestMethod]
        [DataRow("wait-error")]
        [DataRow("wrong-pid")]
        [DataRow("missing-process")]
        [DataRow("unconfirmed-exit")]
        public void UnavailableExitEvidenceCannotReplaceTheOriginalRpcFailure(string fault)
        {
            WithFakeFixture((fixture, application, inspector, process, root) =>
            {
                var failures = new List<string> { "Original RPC failure" };
                if (fault == "wrong-pid") SetProperty(fixture, "ProcessId", -1);
                if (fault == "missing-process") SetField(fixture, "process", null);
                if (fault == "wait-error") fixture.WaitForOwnedShutdownExit = _ => throw new IOException("Synthetic observer failure");
                if (fault == "unconfirmed-exit") fixture.WaitForOwnedShutdownExit = _ => true;
                fixture.ObserveFailedShutdown("Application.Quit", new COMException("RPC failure"), failures);
                Assert.AreEqual("Original RPC failure", failures[0]);
                Assert.AreEqual(2, failures.Count);
                var receipt = Read(Path.Combine(root, "shutdown-failure-observation.json"));
                Assert.AreEqual(false, receipt["ExitObserved"]);
                Assert.AreEqual(false, receipt["ExitCodeObserved"]);
                Assert.AreEqual(false, receipt["NormalQuitQualified"]);
                Assert.IsTrue(receipt.ContainsKey("ObservationErrorType"));
                Assert.AreEqual(0, inspector.CloseCount);
                Assert.AreEqual(0, application.QuitCount);
            });
        }

        [TestMethod]
        [DataRow(0)]
        [DataRow(7)]
        public void ObservedOwnedChildExitRetainsItsActualCodeWithoutQualifyingTheFailedQuit(int exitCode)
        {
            WithFakeFixture((fixture, application, inspector, process, root) =>
            {
                var start = new ProcessStartInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System),
                    "WindowsPowerShell", "v1.0", "powershell.exe"), "-NoProfile -NonInteractive -Command \"exit " + exitCode + "\"")
                { UseShellExecute = false, CreateNoWindow = true };
                using (var child = Process.Start(start))
                {
                    long originalHandle = child.Handle.ToInt64();
                    SetField(fixture, "process", child);
                    SetProperty(fixture, "ProcessId", child.Id);
                    fixture.WaitForOwnedShutdownExit = timeout => child.WaitForExit(timeout);
                    var failures = new List<string> { "Original failed Quit" };
                    fixture.ObserveFailedShutdown("Application.Quit", new COMException("RPC failure"), failures);
                    var receipt = Read(Path.Combine(root, "shutdown-failure-observation.json"));
                    Assert.AreEqual(true, receipt["ExitObserved"]);
                    Assert.AreEqual(true, receipt["ExitCodeObserved"]);
                    Assert.AreEqual(exitCode, receipt["ExitCode"]);
                    Assert.AreEqual(originalHandle, Convert.ToInt64(receipt["OriginalProcessHandle"]));
                    Assert.AreEqual(false, receipt["NormalQuitQualified"]);
                    CollectionAssert.AreEqual(new[] { "Original failed Quit" }, failures);
                    Assert.AreEqual(0, inspector.CloseCount);
                    Assert.AreEqual(0, application.QuitCount);
                }
            });
        }

        [TestMethod]
        public void ReceiptWriteFailureAddsDiagnosticContextWithoutSuppressingTheOriginalFailure()
        {
            WithFakeFixture((fixture, application, inspector, process, root) =>
            {
                var failures = new List<string> { "Original failed Quit" };
                string file = Path.Combine(root, "not-a-directory");
                File.WriteAllText(file, "fixture");
                SetProperty(fixture, "Root", file);
                fixture.ObserveFailedShutdown("Application.Quit", new COMException("RPC failure"), failures);
                Assert.AreEqual("Original failed Quit", failures[0]);
                Assert.AreEqual(2, failures.Count);
                StringAssert.Contains(failures[1], "receipt");
                Assert.AreEqual(0, inspector.CloseCount);
                Assert.AreEqual(0, application.QuitCount);
            });
        }
    }
}
