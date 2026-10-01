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
