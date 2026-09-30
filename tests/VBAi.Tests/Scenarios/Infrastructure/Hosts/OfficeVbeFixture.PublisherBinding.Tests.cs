using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Exercises the actual startup fixture against separate managed wrappers, without Office activation or bridge emission.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class OfficePublisherStartupBindingTests
    {
        public sealed class Window { public long Hwnd { get; set; } = 100; public string Caption { get; set; } = "Disposable.pub - Publisher"; }
        public sealed class Document { public string FullName { get; set; } public bool Saved { get; set; } = true; public Window ActiveWindow { get; set; } = new Window(); }
        public sealed class Documents
        {
            public readonly List<Document> Items = new List<Document>();
            public int Count => Items.Count;
            public Document this[int index] => Items[index - 1];
        }
        public sealed class Application
        {
            public Window ActiveWindow { get; set; } = new Window();
            public Document ActiveDocument { get; set; }
            public Documents Documents { get; } = new Documents();
        }

        [TestMethod]
        public void DistinctWrappersForOneExactOwnedPublicationCanBindWithoutIUnknownEquality()
        {
            InFixture("valid", (fixture, projects, sends) => {
                fixture.BindStartupProject(projects); Assert.IsFalse(string.IsNullOrWhiteSpace(fixture.Project));
                string evidence = File.ReadAllText(Path.Combine(fixture.Root, "adapter-only-progress.json"));
                StringAssert.Contains(evidence, "\"ParentProcessId\":123");
                StringAssert.Contains(evidence, "\"AutomaticRebind\":false");
                Assert.AreEqual(123, fixture.ProcessId, "Observing a child must not change the retained PID.");
            });
        }

        [TestMethod]
        public void ReturnedNewDocumentRequiresWindowOwnershipBeforeBootstrapSaveButNotItsFuturePath()
        {
            InFixture("unsaved-new", (fixture, projects, sends) => {
                fixture.RequirePublisherPublication("BeforeBootstrap", false);
                Assert.IsNull(fixture.Project); Assert.AreEqual(0, sends[0]);
                Assert.ThrowsException<AssertFailedException>(() => fixture.RequirePublisherPublication("AfterBootstrap", true));
            });
        }

        [DataTestMethod]
        [DataRow("doc-pid"), DataRow("app-pid"), DataRow("window-mismatch"), DataRow("multiple-documents"), DataRow("wrong-document")]
        public void UnsafeReturnedDocumentRefusesBeforeAnyBootstrapSaveOrBridgeCommand(string state)
        {
            InFixture(state, (fixture, projects, sends) => {
                Assert.ThrowsException<AssertFailedException>(() => fixture.RequirePublisherPublication("BeforeBootstrap", false));
                Assert.IsNull(fixture.Project); Assert.AreEqual(0, sends[0]);
                Assert.IsTrue(File.ReadAllText(Path.Combine(fixture.Root, "adapter-only-progress.json")).Contains("\"Verified\":false"));
            });
        }

        [DataTestMethod]
        [DataRow("recovered"), DataRow("wrong-document"), DataRow("wrong-path"), DataRow("doc-pid"), DataRow("app-pid")]
        [DataRow("window-mismatch"), DataRow("multiple-documents"), DataRow("unavailable"), DataRow("unverified")]
        [DataRow("owner-mismatch"), DataRow("selected-project"), DataRow("selected-path"), DataRow("missing-selected"), DataRow("duplicate-project")]
        [DataRow("changed-during-readback"), DataRow("unsaved"), DataRow("active-window-pid"), DataRow("wrong-selected-projectpath")]
        [DataRow("wrong-host"), DataRow("run-mode"), DataRow("missing-owner"), DataRow("inventorypath-conflict")]
        [DataRow("retained-process-exited")]
        public void WrongWindowPathRecoveredProjectAndUnverifiedNativeAssociationRefuseBeforeBaseline(string state)
        {
            InFixture(state, (fixture, projects, sends) => {
                    Assert.ThrowsException<AssertFailedException>(() => fixture.BindStartupProject(projects), state);
                    Assert.IsNull(fixture.Project, "Refusal must leave the baseline unbound: " + state);
                    Assert.AreEqual(0, sends[0], "No baseline mutation may reach the bridge: " + state);
            });
        }

        private static void InFixture(string state, Action<OfficeVbeFixture, IDictionary<string, object>[], int[]> test)
        {
            string root = Path.Combine(Path.GetTempPath(), "vbai-publisher-bind-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try
            {
                string path = Path.Combine(root, "Disposable.pub"); File.WriteAllText(path, "synthetic fixture only");
                var retained = new Document { FullName = path };
                var active = new Document { FullName = path };
                var app = new Application { ActiveDocument = active }; app.Documents.Items.Add(active);
                var fixture = (OfficeVbeFixture)Activator.CreateInstance(typeof(OfficeVbeFixture), true);
                Set(fixture, "application", app); Set(fixture, "document", retained);
                foreach (var value in new Dictionary<string, object> { ["Kind"] = "Publisher", ["ProcessId"] = 123, ["Root"] = root, ["DocumentPath"] = path })
                    typeof(OfficeVbeFixture).GetProperty(value.Key, BindingFlags.Instance | BindingFlags.NonPublic).SetValue(fixture, value.Value);
                var project = new Dictionary<string, object> { ["Name"] = "Project", ["FileName"] = path, ["HostPath"] = path, ["Mode"] = 2 };
                var projects = new[] { (IDictionary<string, object>)project };
                var persistence = new Dictionary<string, object> { ["HostAvailable"] = true, ["IdentityVerified"] = true, ["OwnerProcessId"] = 123, ["HostPath"] = path, ["Host"] = "Publisher" };
                var selection = new Dictionary<string, object> { ["SelectedProject"] = "Project", ["SelectedProjectPath"] = path, ["SelectedHostPath"] = path, ["Mode"] = 2 };
                int[] mutations = { 0 };
                fixture.Dispatch = (pid, request) => {
                    var values = (IDictionary<string, object>)request;
                    string command = (string)values["Command"];
                    if (command != "project_persistence_status" && command != "debug_state") { mutations[0]++; Assert.Fail("Unexpected mutation: " + command); }
                    if (state == "changed-during-readback" && command == "debug_state") active.FullName = Path.Combine(root, "Switched.pub");
                    return new Dictionary<string, object> { ["Ok"] = true, ["Error"] = null, ["Data"] = command == "debug_state" ? selection : persistence };
                };
                if (state == "recovered") project["FileName"] = project["HostPath"] = Path.Combine(root, "pubRecovered.tmp");
                if (state == "wrong-document") active.FullName = Path.Combine(root, "OldRecovered.pub");
                if (state == "wrong-path") retained.FullName = Path.Combine(root, "Wrong.pub");
                if (state == "doc-pid") retained.ActiveWindow.Hwnd = 999;
                if (state == "app-pid") app.ActiveWindow.Hwnd = 999;
                if (state == "window-mismatch") retained.ActiveWindow.Hwnd = 101;
                if (state == "multiple-documents") app.Documents.Items.Add(new Document { FullName = Path.Combine(root, "Recovered.pub") });
                if (state == "unavailable") persistence["HostAvailable"] = false;
                if (state == "unverified") persistence["IdentityVerified"] = false;
                if (state == "owner-mismatch") persistence["OwnerProcessId"] = 456;
                if (state == "selected-project") selection["SelectedProject"] = "Other";
                if (state == "selected-path") selection["SelectedHostPath"] = Path.Combine(root, "Other.pub");
                if (state == "missing-selected") selection["SelectedProject"] = null;
                if (state == "duplicate-project") projects = new[] { project, project };
                if (state == "unsaved") retained.Saved = false;
                if (state == "unsaved-new") { retained.FullName = active.FullName = "Publication1"; retained.Saved = active.Saved = false; }
                if (state == "active-window-pid") active.ActiveWindow.Hwnd = 999;
                if (state == "wrong-selected-projectpath") selection["SelectedProjectPath"] = Path.Combine(root, "Other.pub");
                if (state == "wrong-host") persistence["Host"] = "Access";
                if (state == "run-mode") selection["Mode"] = 1;
                if (state == "missing-owner") persistence.Remove("OwnerProcessId");
                if (state == "inventorypath-conflict") project["FileName"] = Path.Combine(root, "Recovered.tmp");
                // The only injected seam is Win32 owner observation. The fake wrappers remain distinct objects.
                fixture.ReadPublisherWindowOwner = hwnd => hwnd.ToInt64() == 999 ? 456u : 123u;
                fixture.ReadPublisherProcessCanary = () => new object[] {
                    new { ProcessId = 456, ParentProcessId = 123, ObservedOnly = true, Adopted = false } };
                fixture.ReadPublisherOwnedProcessAlive = () => state != "retained-process-exited";
                test(fixture, projects, mutations);
            }
            finally { Directory.Delete(root, true); }
        }

        private static void Set(object target, string name, object value) => target.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).SetValue(target, value);
    }
}
