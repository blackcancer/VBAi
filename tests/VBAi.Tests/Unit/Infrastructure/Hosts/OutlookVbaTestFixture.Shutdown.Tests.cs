using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Web.Script.Serialization;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Uses managed fake COM objects and a testhost query handle; never activates or closes Outlook.</summary>
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class OutlookVbaTestFixtureShutdownTests
    {
        [TestMethod]
        public void PendingNativeRunRetainsAllReferencesAndOtmWithoutAnyCleanupOrRetry()
        {
            WithFakeFixture((fixture, application, inspector, process, root) =>
            {
                SetField(fixture, "runPending", true);
                fixture.Dispatch = (pid, request) => { Assert.Fail("A pending run cannot enter native cleanup."); return null; };
                Assert.ThrowsException<AssertFailedException>(() => fixture.Dispose());
                AssertRetained(fixture, application, inspector, process, root);
                Assert.AreEqual(0, inspector.CloseCount); Assert.AreEqual(0, application.QuitCount);
                SetField(fixture, "runPending", false);
                fixture.Dispose(); // A later observation must not re-arm disposal of this retained generation.
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Response("list_modules"));
                Assert.AreEqual(0, inspector.CloseCount); Assert.AreEqual(0, application.QuitCount);
            });
        }

        [TestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void UncertainOwnedModuleRemovalRetainsOwnershipAndNeverClosesOrQuits(bool throwAfterDispatch)
        {
            WithFakeFixture((fixture, application, inspector, process, root) =>
            {
                SetField(fixture, "baselineVerified", true);
                var modules = (Dictionary<string, string>)Field(fixture, "ownedModules");
                modules.Add("SyntheticOwnedModule", "owned-hash");
                var commands = new List<string>();
                fixture.Dispatch = (pid, request) =>
                {
                    var command = (string)((IDictionary<string, object>)request)["Command"];
                    commands.Add(command);
                    if (command == "read_module") return Reply(new Dictionary<string, object> { ["Sha256"] = "owned-hash" });
                    if (command == "component_properties" || command == "project_properties")
                        return Reply(new Dictionary<string, object> { ["Version"] = "reviewed-version" });
                    Assert.AreEqual("remove_component", command);
                    if (throwAfterDispatch) throw new IOException("synthetic mutation outcome unknown");
                    return null;
                };
                Assert.ThrowsException<AssertFailedException>(() => fixture.Dispose());
                CollectionAssert.AreEqual(new[] { "read_module", "component_properties", "project_properties", "remove_component" }, commands);
                AssertRetained(fixture, application, inspector, process, root);
                var lifecycle = Read(Path.Combine(root, "qualification-lifecycle.json"));
                Assert.AreEqual(true, lifecycle["BridgeUncertain"]);
                Assert.AreEqual(false, lifecycle["NativeExecutionUnsettled"]);
                fixture.Dispose();
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Response("remove_component"));
                Assert.AreEqual(4, commands.Count); Assert.AreEqual(0, inspector.CloseCount); Assert.AreEqual(0, application.QuitCount);
            });
        }

        [TestMethod]
        public void UncertainInspectorCloseRetainsOriginalReferencesBeforeAnyQuit()
        {
            WithFakeFixture((fixture, application, inspector, process, root) =>
            {
                inspector.CloseError = new InvalidOperationException("synthetic inspector close outcome unknown");
                Assert.ThrowsException<AssertFailedException>(() => fixture.Dispose());
                AssertRetained(fixture, application, inspector, process, root);
                Assert.AreEqual(1, inspector.CloseCount); Assert.AreEqual(0, application.QuitCount);
                fixture.Dispose(); Assert.AreEqual(1, inspector.CloseCount);
            });
        }

        [TestMethod]
        public void UncertainQuitRetainsOriginalReferencesAndQueryHandleWithoutRetry()
        {
            WithFakeFixture((fixture, application, inspector, process, root) =>
            {
                application.QuitError = new InvalidOperationException("synthetic quit outcome unknown");
                Assert.ThrowsException<AssertFailedException>(() => fixture.Dispose());
                AssertRetained(fixture, application, inspector, process, root);
                Assert.AreEqual(1, inspector.CloseCount); Assert.AreEqual(1, application.QuitCount);
                fixture.Dispose(); Assert.AreEqual(1, application.QuitCount);
            });
        }

        [TestMethod]
        public void ConfirmedCompletedRunClearsPendingWithoutRetainingSettledOwnership()
        {
            WithFakeFixture((fixture, application, inspector, process, root) =>
            {
                fixture.Dispatch = (pid, request) =>
                {
                    string command = (string)((IDictionary<string, object>)request)["Command"];
                    return command == "run_vba_tests"
                        ? Reply(new Dictionary<string, object> { ["Query"] = "synthetic-run" })
                        : Reply(new Dictionary<string, object>
                        {
                            ["Pending"] = false,
                            ["Report"] = new Dictionary<string, object> { ["uncertain"] = false }
                        });
                };
                fixture.Response("run_vba_tests"); Assert.AreEqual(true, Field(fixture, "runPending"));
                fixture.Response("vba_test_run_status"); Assert.AreEqual(false, Field(fixture, "runPending"));
                Assert.AreEqual(false, Field(fixture, "hostTeardownRefused"));
                Assert.AreEqual(0, inspector.CloseCount); Assert.AreEqual(0, application.QuitCount);
            });
        }

        private static IDictionary<string, object> Reply(object data)
            => new Dictionary<string, object> { ["Ok"] = true, ["Error"] = null, ["Data"] = data };

        private static void AssertRetained(OutlookVbaTestFixture fixture, FakeApplication application, FakeInspector inspector, Process process, string root)
        {
            Assert.AreSame(process, Field(fixture, "process")); Assert.AreNotEqual(IntPtr.Zero, process.Handle);
            Assert.AreSame(application, Field(fixture, "application")); Assert.AreSame(inspector, Field(fixture, "inspector"));
            Assert.IsNotNull(Field(fixture, "item")); Assert.IsNotNull(Field(fixture, "commandBars"));
            var retained = (IList)typeof(OutlookVbaTestFixture).GetField("retainedOutlookFixtures", BindingFlags.Static | BindingFlags.NonPublic).GetValue(null);
            lock (retained) Assert.IsTrue(retained.Contains(fixture));
            CollectionAssert.AreEqual(new byte[] { 1, 2, 3 }, File.ReadAllBytes(Path.Combine(root, "managed-fake.OTM")));
            var lifecycle = Read(Path.Combine(root, "qualification-lifecycle.json"));
            Assert.AreEqual(true, lifecycle["OwnershipRetained"]); Assert.AreEqual(true, lifecycle["ComReferencesRetained"]);
            Assert.AreEqual(true, lifecycle["ProcessHandleRetained"]); Assert.AreEqual(false, lifecycle["ExitedNormally"]);
            Assert.IsFalse(File.Exists(Path.Combine(root, "otm-cleanup.json")));
        }

        private static void WithFakeFixture(Action<OutlookVbaTestFixture, FakeApplication, FakeInspector, Process, string> action)
        {
            string root = Path.Combine(Path.GetTempPath(), "VBAi-Outlook-ManagedShutdown-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            var process = Process.GetCurrentProcess(); _ = process.Handle;
            var fixture = new OutlookVbaTestFixture(); var application = new FakeApplication(); var inspector = new FakeInspector();
            SetProperty(fixture, "Root", root); SetProperty(fixture, "ProcessId", process.Id);
            SetProperty(fixture, "Project", "ManagedFakeOnly"); SetProperty(fixture, "OtmPath", Path.Combine(root, "managed-fake.OTM"));
            File.WriteAllBytes(fixture.OtmPath, new byte[] { 1, 2, 3 });
            SetField(fixture, "process", process); SetField(fixture, "application", application); SetField(fixture, "inspector", inspector);
            SetField(fixture, "item", new object()); SetField(fixture, "commandBars", new object());
            fixture.VerifyOwnedProcess = () => { }; // All application operations below target only these managed fakes.
            try { action(fixture, application, inspector, process, root); }
            finally
            {
                process.Dispose(); // Release only this testhost query handle, never a native application process.
                var retained = (IList)typeof(OutlookVbaTestFixture).GetField("retainedOutlookFixtures", BindingFlags.Static | BindingFlags.NonPublic).GetValue(null);
                lock (retained) retained.Remove(fixture);
                Directory.Delete(root, true);
            }
        }

        private static object Field(object target, string name) => target.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(target);
        private static void SetField(object target, string name, object value) => target.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).SetValue(target, value);
        private static void SetProperty(object target, string name, object value) => target.GetType().GetProperty(name, BindingFlags.Instance | BindingFlags.NonPublic).SetValue(target, value);
        private static IDictionary<string, object> Read(string path) => (IDictionary<string, object>)new JavaScriptSerializer().DeserializeObject(File.ReadAllText(path));
        public sealed class FakeApplication { public int QuitCount; public Exception QuitError; public void Quit() { QuitCount++; if (QuitError != null) throw QuitError; } }
        public sealed class FakeInspector { public int CloseCount; public Exception CloseError; public void Close(int option) { Assert.AreEqual(1, option); CloseCount++; if (CloseError != null) throw CloseError; } }
    }
}