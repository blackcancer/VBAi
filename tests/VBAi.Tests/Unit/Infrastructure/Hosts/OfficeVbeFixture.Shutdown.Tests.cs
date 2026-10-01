using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Exercises real fixture teardown with managed fake COM objects and only the testhost's own query handle; no Office is launched or quit.</summary>
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class OfficeVbeFixtureShutdownTests
    {
        [TestMethod]
        public void FixtureNonexitRetainsExactOriginalHandleAndBlocksAllNativeReuseWithoutMaskingFirstFailure()
        {
            WithFakeFixture((fixture, application, document, process, root) => {
                fixture.WaitForOwnedExit = (observed, timeout) => { Assert.AreSame(process, observed); Assert.AreEqual(5000, timeout); return false; };
                fixture.ReadOwnedExitCode = observed => { Assert.Fail("A nonexited host has no observed exit code."); return 0; };
                var original = new InvalidOperationException("original qualification failure");
                var actual = Assert.ThrowsException<AggregateException>(() => OfficeMetadataMutationEvidence.Run(() => { throw original; }, fixture.Dispose));
                Assert.AreSame(original, actual.InnerExceptions[0]);
                Assert.AreSame(process, Field(fixture, "ownedProcess")); Assert.AreEqual(true, Field(fixture, "owned"));
                Assert.AreEqual(true, Field(fixture, "hostTeardownRefused")); Assert.AreNotEqual(IntPtr.Zero, process.Handle);
                Assert.AreEqual(1, application.QuitCount); Assert.AreEqual(1, document.CloseCount);
                int dispatches = 0; fixture.Dispatch = (pid, request) => { dispatches++; throw new Exception("No dispatch is allowed."); };
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Response("save_host_document"));
                Assert.ThrowsException<InvalidOperationException>(() => fixture.SaveNative());
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Reopen());
                Assert.ThrowsException<InvalidOperationException>(() => fixture.ReopenFromDisk());
                Assert.ThrowsException<AssertFailedException>(() => fixture.Dispose());
                Assert.AreEqual(0, dispatches); Assert.AreEqual(1, application.QuitCount); Assert.AreEqual(1, document.CloseCount);
                var report = Read(Path.Combine(root, "shutdown-lifecycle.json"));
                var life = (IDictionary<string, object>)report["Lifecycle"];
                Assert.AreEqual(process.Id, life["ProcessId"]); Assert.AreEqual("RETURNED", life["QuitOutcome"]);
                Assert.AreEqual(true, life["ProcessHandleRetained"]); Assert.AreEqual(false, life["ProcessExitObserved"]);
                Assert.AreEqual(false, life["ExitCodeObserved"]); Assert.IsNull(life["ExitCode"]);
                Assert.IsTrue(((object[])Read(Path.Combine(root, "qualification.json"))["Failures"]).Length > 0);
            });
        }

        [TestMethod]
        public void FixtureKnownExitReleasesOnlyItsQueryHandleAndDoesNotQuitAgain()
        {
            WithFakeFixture((fixture, application, document, process, root) => {
                // Both observations are synthetic. Disposing this Process wrapper
                // releases a query handle; it never terminates the testhost.
                fixture.WaitForOwnedExit = (observed, timeout) => true;
                fixture.ReadOwnedExitCode = observed => 0;
                fixture.Dispose();
                Assert.IsNull(Field(fixture, "ownedProcess")); Assert.AreEqual(false, Field(fixture, "owned"));
                Assert.AreEqual(1, application.QuitCount); Assert.AreEqual(1, document.CloseCount);
                fixture.Dispose(); Assert.AreEqual(1, application.QuitCount);
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Response("save_host_document"));
                var life = (IDictionary<string, object>)Read(Path.Combine(root, "shutdown-lifecycle.json"))["Lifecycle"];
                Assert.AreEqual(true, life["ProcessExitObserved"]); Assert.AreEqual(0, life["ExitCode"]);
                Assert.AreEqual(false, life["ProcessHandleRetained"]); Assert.AreEqual(false, life["ForcedTermination"]);
            });
        }

        [TestMethod]
        public void FixtureFailedQuitRetainsOriginalApplicationHandleAndRejectsRetry()
        {
            WithFakeFixture((fixture, application, document, process, root) => {
                application.QuitError = new InvalidOperationException("synthetic original Quit failure");
                fixture.WaitForOwnedExit = (observed, timeout) => { Assert.Fail("Unknown Quit must retain its process without exit acceptance."); return true; };
                Assert.ThrowsException<AssertFailedException>(() => fixture.Dispose());
                Assert.AreSame(process, Field(fixture, "ownedProcess")); Assert.AreSame(application, Field(fixture, "application"));
                Assert.AreEqual(true, Field(fixture, "owned")); Assert.AreEqual(1, application.QuitCount);
                Assert.ThrowsException<AssertFailedException>(() => fixture.Dispose()); Assert.AreEqual(1, application.QuitCount);
                var life = (IDictionary<string, object>)Read(Path.Combine(root, "shutdown-lifecycle.json"))["Lifecycle"];
                Assert.AreEqual("CALL_FAILED_EFFECT_UNKNOWN", life["QuitOutcome"]);
                StringAssert.Contains(Convert.ToString(life["OriginalQuitError"]), application.QuitError.Message);
                Assert.AreEqual(true, life["ProcessHandleRetained"]); Assert.AreEqual(false, life["ProcessExitObserved"]);
            });
        }

        [TestMethod]
        public void FixturePendingNativeTestRetainsAllOwnershipBeforeCloseOrQuitAndRejectsLaterCleanupRetry()
        {
            WithFakeFixture((fixture, application, document, process, root) => {
                fixture.NativeExecutionUnsettled = true;
                fixture.WaitForOwnedExit = (observed, timeout) => { Assert.Fail("A pending test must not enter shutdown observation."); return false; };
                Assert.ThrowsException<AssertFailedException>(() => fixture.Dispose());
                Assert.AreSame(process, Field(fixture, "ownedProcess"));
                Assert.AreSame(application, Field(fixture, "application"));
                Assert.AreSame(document, Field(fixture, "document"));
                Assert.AreEqual(true, Field(fixture, "hostTeardownRefused"));
                Assert.AreEqual(0, application.QuitCount); Assert.AreEqual(0, document.CloseCount);
                Assert.AreEqual(true, Read(Path.Combine(root, "qualification.json"))["NativeExecutionUnsettled"]);
                fixture.NativeExecutionUnsettled = false;
                Assert.ThrowsException<AssertFailedException>(() => fixture.Dispose());
                Assert.AreEqual(0, application.QuitCount); Assert.AreEqual(0, document.CloseCount);
            });
        }

        [TestMethod]
        public void ReviewedSupportApprovalCannotLatchWithoutItsOriginalLiveDialogGeneration()
        {
            WithFakeFixture((fixture, application, document, process, root) => {
                var approve = typeof(OfficeVbeFixture).GetMethod("SetReviewedSupportSaveAllowed", BindingFlags.Instance | BindingFlags.NonPublic);
                Assert.IsInstanceOfType(Assert.ThrowsException<TargetInvocationException>(() => approve.Invoke(fixture, new object[] { true })).InnerException,
                    typeof(AssertFailedException));
                Assert.AreEqual(false, Field(fixture, "allowSupportSavePrompt"));
                var workerType = typeof(OfficeVbeFixture).GetNestedType("OwnedDialogWorker", BindingFlags.NonPublic);
                var worker = Activator.CreateInstance(workerType, BindingFlags.Instance | BindingFlags.NonPublic, null,
                    new object[] { process, "Access" }, null);
                SetField(worker, "StopRequested", true); SetField(fixture, "dialogWorker", worker);
                Assert.IsInstanceOfType(Assert.ThrowsException<TargetInvocationException>(() => approve.Invoke(fixture, new object[] { true })).InnerException,
                    typeof(AssertFailedException));
                Assert.AreEqual(false, Field(fixture, "allowSupportSavePrompt"));
                Assert.AreEqual(false, Field(worker, "ReviewedSupportSaveAllowed"));
                approve.Invoke(fixture, new object[] { false });
                SetField(fixture, "dialogWorker", null);
            });
        }

        [STATestMethod]
        public void WordShutdownPumpRunsOnOwnerAfterKnownQuitAndReleaseAndObservesOriginalHandle()
        {
            WithFakeFixture((fixture, application, document, process, root) => {
                bool exited = false; int pumps = 0;
                int ownerThread = System.Threading.Thread.CurrentThread.ManagedThreadId;
                IntPtr original = process.Handle;
                fixture.ReadWordProcessExit = observed => {
                    Assert.AreSame(process, observed); Assert.AreEqual(original, observed.Handle);
                    Assert.IsNull(Field(fixture, "application")); Assert.IsNull(Field(fixture, "document"));
                    Assert.AreEqual(1, application.QuitCount); Assert.AreEqual(1, document.CloseCount);
                    return exited;
                };
                fixture.PumpWordShutdownMessages = () => {
                    Assert.AreEqual(ownerThread, System.Threading.Thread.CurrentThread.ManagedThreadId);
                    Assert.ThrowsException<InvalidOperationException>(() => fixture.Data("status"), "Pumped messages must not reenter native fixture requests after shutdown was prepared.");
                    pumps++; exited = true;
                };
                fixture.ReadOwnedExitCode = observed => { Assert.AreSame(process, observed); return 0; };
                fixture.Dispose();
                Assert.AreEqual(1, pumps);
                var lifecycle = (IDictionary<string, object>)Read(Path.Combine(root, "shutdown-lifecycle.json"))["Lifecycle"];
                Assert.AreEqual(true, lifecycle["ProcessExitObserved"]);
                Assert.AreEqual(true, lifecycle["WordMessagePumpEnabled"]);
                Assert.AreEqual(1, lifecycle["WordMessagePumpAttempts"]);
                Assert.AreEqual(5000, lifecycle["WaitBoundMilliseconds"]);
                Assert.AreEqual(ownerThread, lifecycle["ExitObservationThread"]);
                Assert.AreEqual(lifecycle["OriginalProcessHandle"], lifecycle["ExitWaitProcessHandle"]);
                Assert.IsTrue(Convert.ToInt64(lifecycle["ExitWaitElapsedMilliseconds"]) < 5000);
            });
        }

        [TestMethod]
        public void WordShutdownPumpKeepsTheDeadlineAndReportsNonexitWithoutNativeRetry()
        {
            long elapsed = 0; int pumps = 0, reads = 0;
            bool result = OfficeVbeFixture.WaitForWordExit(5000, () => elapsed, () => { reads++; return false; },
                () => pumps++, milliseconds => { Assert.IsTrue(milliseconds > 0 && milliseconds <= 25); elapsed += milliseconds; });
            Assert.IsFalse(result); Assert.AreEqual(5000L, elapsed); Assert.AreEqual(200, pumps);
            Assert.AreEqual(401, reads);
            elapsed = 4999; pumps = 0;
            Assert.IsFalse(OfficeVbeFixture.WaitForWordExit(5000, () => elapsed, () => false,
                () => { pumps++; elapsed += 1; }, _ => Assert.Fail("No pause may extend an expired deadline.")));
            Assert.AreEqual(1, pumps);
            elapsed = 0; bool lateExit = false;
            Assert.IsFalse(OfficeVbeFixture.WaitForWordExit(5000, () => elapsed, () => lateExit,
                () => { elapsed = 5001; lateExit = true; }, _ => Assert.Fail("A message pump cannot extend the exit-acceptance deadline.")));
            Assert.IsTrue(OfficeVbeFixture.WaitForWordExit(5000, () => 5000, () => true,
                () => Assert.Fail("Already-exited processes require no pump."), _ => Assert.Fail("Already-exited processes require no pause.")));
        }
        [TestMethod]
        public void WordShutdownPumpRejectsPauseAndExitGetterOvershootButAcceptsTheExactDeadline()
        {
            long elapsed = 0; bool exited = false; int reads = 0;
            Assert.IsFalse(OfficeVbeFixture.WaitForWordExit(5000, () => elapsed,
                () => { reads++; return exited; }, () => { }, _ => { elapsed = 5010; exited = true; }));
            Assert.AreEqual(2, reads, "A pause beyond the deadline must not permit another exit read.");
            foreach (bool afterPump in new[] { false, true })
            {
                elapsed = 0; reads = 0;
                Assert.IsFalse(OfficeVbeFixture.WaitForWordExit(5000, () => elapsed, () => {
                    reads++;
                    if (!afterPump || reads == 2) { elapsed = 5001; return true; }
                    return false;
                }, () => { }, _ => Assert.Fail("An exit read exceeding the deadline must stop observation.")));
                Assert.AreEqual(afterPump ? 2 : 1, reads);
            }
            elapsed = 5000;
            Assert.IsTrue(OfficeVbeFixture.WaitForWordExit(5000, () => elapsed, () => true,
                () => Assert.Fail("A confirmed exit at the exact deadline needs no pump."), _ => Assert.Fail()));
            elapsed = 5001;
            Assert.IsFalse(OfficeVbeFixture.WaitForWordExit(5000, () => elapsed,
                () => { Assert.Fail("No exit read is allowed after the deadline."); return true; }, () => Assert.Fail(), _ => Assert.Fail()));
        }
        [TestMethod]
        public void PreparedAccessShutdownAllowsSupportReadbackWhilePreparedWordBlocksReentry()
        {
            foreach (string kind in new[] { "Access", "Word" })
            WithFakeFixture((fixture, application, document, process, root) => {
                SetProperty(fixture, "Kind", kind);
                int dispatches = 0;
                fixture.Dispatch = (pid, request) => {
                    Assert.AreEqual(process.Id, pid); Assert.AreEqual("read_module", ((IDictionary<string, object>)request)["Command"]);
                    dispatches++;
                    return new Dictionary<string, object> { ["Ok"] = true, ["Error"] = null, ["Data"] = new Dictionary<string, object> { ["Code"] = "synthetic reviewed support" } };
                };
                typeof(OfficeVbeFixture).GetMethod("PrepareOwnedShutdown", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(fixture, null);
                if (kind == "Word")
                    Assert.ThrowsException<InvalidOperationException>(() => fixture.Data("read_module", "Module", "Support"));
                else
                    Assert.AreEqual("synthetic reviewed support", fixture.Data("read_module", "Module", "Support")["Code"]);
                Assert.AreEqual(kind == "Word" ? 0 : 1, dispatches);
                Assert.AreEqual(0, application.QuitCount); Assert.AreEqual(0, document.CloseCount);
            });
        }
        private static void WithFakeFixture(Action<OfficeVbeFixture, FakeApplication, FakeDocument, Process, string> action)
        {
            string root = Path.Combine(Path.GetTempPath(), "VBAi-OwnedShutdown-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            var fixture = (OfficeVbeFixture)Activator.CreateInstance(typeof(OfficeVbeFixture), true);
            var application = new FakeApplication(); var document = new FakeDocument();
            SetProperty(fixture, "Root", root); SetProperty(fixture, "Kind", "Word");
            using (var current = Process.GetCurrentProcess()) SetProperty(fixture, "ProcessId", current.Id);
            SetProperty(fixture, "DocumentPath", Path.Combine(root, "ManagedFakeOnly.docm"));
            SetField(fixture, "owned", true); SetField(fixture, "application", application); SetField(fixture, "document", document);
            typeof(OfficeVbeFixture).GetMethod("CaptureOwnedProcess", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(fixture, null);
            var process = (Process)Field(fixture, "ownedProcess");
            try { action(fixture, application, document, process, root); }
            finally
            {
                process.Dispose(); // Query handle only; no Kill/CloseMainWindow/COM operation.
                var retained = (IList)typeof(OfficeVbeFixture).GetField("retainedOfficeFixtures", BindingFlags.Static | BindingFlags.NonPublic).GetValue(null);
                lock (retained) retained.Remove(fixture); // Remove only this synthetic managed fixture.
                Directory.Delete(root, true);
            }
        }

        private static object Field(object target, string name) => target.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(target);
        private static void SetField(object target, string name, object value) => target.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).SetValue(target, value);
        private static void SetProperty(object target, string name, object value) => target.GetType().GetProperty(name, BindingFlags.Instance | BindingFlags.NonPublic).SetValue(target, value);
        private static IDictionary<string, object> Read(string path) => (IDictionary<string, object>)new JavaScriptSerializer().DeserializeObject(File.ReadAllText(path));
        public sealed class FakeApplication { public int QuitCount; public Exception QuitError; public void Quit(int option) { QuitCount++; if (QuitError != null) throw QuitError; } }
        public sealed class FakeDocument { public int CloseCount; public void Close(int option) { CloseCount++; } }
    }
}
