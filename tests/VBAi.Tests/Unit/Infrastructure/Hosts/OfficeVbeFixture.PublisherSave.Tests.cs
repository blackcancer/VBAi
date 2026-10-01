using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Exercises the actual fixture helper using managed objects and a query handle to the testhost; no Office is activated.</summary>
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class OfficeVbeFixturePublisherSaveTests
    {
        private const string ModuleName = "VBAiOfficeModule";
        private const string Source = "Public Sub Synthetic()\r\nEnd Sub\r\n";

        [TestMethod]
        public void ExactReviewedProjectSavesOnceAndRequiresNativeAndBridgeSavedReadbacks()
        {
            WithFixture((fixture, app, project, expected) => {
                fixture.SaveReviewedPublisherProject(ModuleName, expected);
                Assert.AreEqual(1, app.Editor.Control.Executions);
                Assert.IsTrue(project.Saved && app.Document.Saved);
                Assert.IsFalse(fixture.NativeExecutionUnsettled);
                Assert.IsFalse((bool)Field(fixture, "hostTeardownRefused"));
                Assert.AreEqual(1, app.BridgeReads);
                Assert.AreSame(project, app.Editor.ActiveVBProject);
                Assert.AreSame(project.Components.Component.Code.Pane, app.Editor.ActiveCodePane);
            });
        }

        [TestMethod]
        public void ChangedOwnershipSourceSelectionAndControlAreRefusedBeforeEmission()
        {
            Action<OfficeVbeFixture, FakeApplication, FakeProject>[] changes = {
                (fixture, app, project) => project.Mode = 1,
                (fixture, app, project) => project.Protection = 1,
                (fixture, app, project) => app.Document.ReadOnly = true,
                (fixture, app, project) => app.Document.FullName += ".other",
                (fixture, app, project) => app.Documents.Count = 2,
                (fixture, app, project) => app.Documents.Current = new FakeDocument(),
                (fixture, app, project) => app.ActiveDocument = new FakeDocument(),
                (fixture, app, project) => fixture.ReadPublisherWindowOwner = window => 0,
                (fixture, app, project) => project.Components.Component.Code.Source += "' unreviewed",
                (fixture, app, project) => project.Components.Component.Name = "OtherModule",
                (fixture, app, project) => app.Editor.Control = null,
                (fixture, app, project) => app.Editor.Control.Id = 4,
                (fixture, app, project) => app.Editor.Control.Type = 2,
                (fixture, app, project) => app.Editor.Control.BuiltIn = false,
                (fixture, app, project) => app.Editor.Control.Enabled = false,
                (fixture, app, project) => app.Editor.Control.OnAction = "UnreviewedMacro",
                (fixture, app, project) => app.Editor.Window.AfterFocus = () => app.Editor.ActiveCodePane = new FakePane(),
                (fixture, app, project) => app.Editor.Window.AfterFocus = () => app.Editor.ActiveVBProject = new FakeProject(),
                (fixture, app, project) => app.Editor.Window.AfterFocus = () => app.Editor.ActiveWindow = new FakeWindow(),
                (fixture, app, project) => app.Editor.Window.Visible = false,
                (fixture, app, project) => app.Editor.Window.Type = 1,
                (fixture, app, project) => app.Editor.Window.AfterFocus = () => project.Components.Component.Code.Pane.FirstLine = 2
            };
            foreach (var change in changes)
                WithFixture((fixture, app, project, expected) => {
                    var originalControl = app.Editor.Control;
                    change(fixture, app, project);
                    Assert.ThrowsException<AssertFailedException>(() => fixture.SaveReviewedPublisherProject(ModuleName, expected));
                    Assert.AreEqual(0, originalControl.Executions);
                    Assert.AreEqual(0, app.BridgeReads);
                    Assert.IsFalse(fixture.NativeExecutionUnsettled);
                });
        }

        [TestMethod]
        public void ControlGetterSideEffectsCannotRevokeReviewedAuthorityBeforeSave()
        {
            foreach (int affectedLookup in new[] { 1, 2 })
                WithFixture((fixture, app, project, expected) => {
                    app.Editor.Bars.AfterFind = lookup => {
                        if (lookup == affectedLookup) project.Components.Component.Code.Source += "' changed during lookup";
                    };
                    Assert.ThrowsException<AssertFailedException>(() => fixture.SaveReviewedPublisherProject(ModuleName, expected));
                    Assert.AreEqual(0, app.Editor.Control.Executions);
                    Assert.IsFalse(fixture.NativeExecutionUnsettled);
                });
            WithFixture((fixture, app, project, expected) => {
                app.Editor.Bars.AfterFind = lookup => { if (lookup == 2) app.Editor.Control.Caption = "Another action"; };
                Assert.ThrowsException<AssertFailedException>(() => fixture.SaveReviewedPublisherProject(ModuleName, expected));
                Assert.AreEqual(0, app.Editor.Control.Executions);
            });
            WithFixture((fixture, app, project, expected) => {
                var control = app.Editor.Control;
                app.Editor.Bars.AfterFind = lookup => { if (lookup == 1) app.Editor.Control = new FakeControl(); };
                Assert.ThrowsException<AssertFailedException>(() => fixture.SaveReviewedPublisherProject(ModuleName, expected));
                Assert.AreEqual(0, control.Executions);
                Assert.AreEqual(0, app.Editor.Control.Executions);
            });
        }

        [TestMethod]
        public void FailedOrUnverifiedEmissionRetainsOwnershipAndForbidsSaveAndQuitRetry()
        {
            Action<FakeApplication, FakeProject>[] changes = {
                (app, project) => app.Editor.Control.AfterExecute = () => { throw new InvalidOperationException("native save outcome unknown"); },
                (app, project) => app.Editor.Control.AfterExecute = () => project.Saved = false,
                (app, project) => app.Editor.Control.AfterExecute = () => app.Document.Saved = false,
                (app, project) => app.Editor.Control.AfterExecute = () => project.Components.Component.Code.Source += "' changed during save",
                (app, project) => app.Persistence["ProjectSaved"] = false,
                (app, project) => app.Persistence["HostSaved"] = false,
                (app, project) => app.Persistence["IdentityVerified"] = false,
                (app, project) => app.Persistence["OwnerProcessId"] = -1,
                (app, project) => app.Persistence["HostPath"] = "C:\\other.pub",
                (app, project) => app.AfterReadback = () => project.Saved = false,
                (app, project) => app.AfterReadback = () => project.Components.Component.Code.Source += "' changed during readback"
            };
            foreach (var change in changes)
                WithFixture((fixture, app, project, expected) => {
                    change(app, project);
                    Exception refusal = null;
                    try { fixture.SaveReviewedPublisherProject(ModuleName, expected); }
                    catch (Exception error) { refusal = error; }
                    Assert.IsNotNull(refusal, "An unverified save was accepted.");
                    Assert.IsTrue(refusal is AssertFailedException || refusal is InvalidOperationException, refusal.ToString());
                    if (refusal is InvalidOperationException) StringAssert.Contains(refusal.Message, "native save outcome unknown");
                    Assert.AreEqual(1, app.Editor.Control.Executions);
                    Assert.IsTrue(fixture.NativeExecutionUnsettled);
                    Assert.IsTrue((bool)Field(fixture, "hostTeardownRefused"));
                    Assert.IsTrue(((List<object>)Field(fixture, "retainedDiagnosticReferences")).Count > 0);
                    Assert.ThrowsException<InvalidOperationException>(() => fixture.SaveReviewedPublisherProject(ModuleName, expected));
                    Assert.ThrowsException<InvalidOperationException>(() => fixture.SaveNative());
                    fixture.Dispose();
                    Assert.AreEqual(0, app.QuitCalls);
                    Assert.AreEqual(1, app.Editor.Control.Executions);
                });
        }

        [TestMethod]
        public void UnsettledExecutionCannotStartAReviewedSave()
        {
            WithFixture((fixture, app, project, expected) => {
                fixture.NativeExecutionUnsettled = true;
                Assert.ThrowsException<AssertFailedException>(() => fixture.SaveReviewedPublisherProject(ModuleName, expected));
                Assert.AreEqual(0, app.Editor.Control.Executions);
                Assert.AreEqual(0, app.BridgeReads);
            });
        }

        [TestMethod]
        public void SourceReviewAllowsLineEndingNormalizationButRefusesInventoryAndContentChanges()
        {
            var expected = new Dictionary<string, string> { [ModuleName] = Source };
            OfficeVbeFixture.RequirePublisherSaveSources(expected, new Dictionary<string, string> { [ModuleName] = Source.Replace("\r\n", "\n").TrimEnd('\n') });
            Assert.ThrowsException<AssertFailedException>(() => OfficeVbeFixture.RequirePublisherSaveSources(expected, new Dictionary<string, string>()));
            Assert.ThrowsException<AssertFailedException>(() => OfficeVbeFixture.RequirePublisherSaveSources(expected, new Dictionary<string, string> { ["DifferentModule"] = Source }));
            Assert.ThrowsException<AssertFailedException>(() => OfficeVbeFixture.RequirePublisherSaveSources(expected, new Dictionary<string, string> { [ModuleName] = Source + "' changed" }));
        }

        private static void WithFixture(Action<OfficeVbeFixture, FakeApplication, FakeProject, IDictionary<string, string>> action)
        {
            string root = Path.Combine(Path.GetTempPath(), "VBAi-Publisher-save-mirror-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            string path = Path.Combine(root, "Disposable.pub"); File.WriteAllText(path, "Managed fixture only");
            var fixture = (OfficeVbeFixture)Activator.CreateInstance(typeof(OfficeVbeFixture), true);
            var process = Process.GetCurrentProcess();
            var app = new FakeApplication(); var project = new FakeProject();
            app.Document = new FakeDocument { FullName = path, VBProject = project }; app.ActiveDocument = app.Document;
            app.Documents = new FakeDocuments { Current = app.Document }; project.VBE = app.Editor;
            project.Components.Component.Code.Parent = project.Components.Component;
            var pane = project.Components.Component.Code.Pane;
            pane.CodeModule = project.Components.Component.Code; pane.Window = app.Editor.Window;
            pane.AfterShow = () => { app.Editor.ActiveCodePane = pane; app.Editor.ActiveWindow = app.Editor.Window; };
            app.Editor.Control.AfterExecute = () => { project.Saved = true; app.Document.Saved = true; };
            Set(fixture, "application", app); Set(fixture, "document", app.Document); Set(fixture, "owned", true); Set(fixture, "ownedProcess", process);
            Property(fixture, "Kind", "Publisher"); Property(fixture, "Root", root); Property(fixture, "DocumentPath", path);
            Property(fixture, "ProcessId", process.Id); Property(fixture, "Project", "SyntheticPublisher");
            fixture.ReadPublisherWindowOwner = window => (uint)process.Id;
            app.Persistence = new Dictionary<string, object> { ["Host"] = "Publisher", ["HostPath"] = path, ["OwnerProcessId"] = process.Id,
                ["IdentityVerified"] = true, ["ProjectSaved"] = true, ["HostSaved"] = true };
            fixture.Dispatch = (pid, request) => {
                var values = (IDictionary<string, object>)request;
                Assert.AreEqual("project_persistence_status", values["Command"]); Assert.AreEqual(path, values["Project"]);
                app.BridgeReads++; app.AfterReadback?.Invoke();
                return new Dictionary<string, object> { ["Ok"] = true, ["Error"] = null, ["Data"] = app.Persistence };
            };
            try { action(fixture, app, project, new Dictionary<string, string> { [ModuleName] = Source }); }
            finally
            {
                var retained = (List<OfficeVbeFixture>)typeof(OfficeVbeFixture).GetField("retainedOfficeFixtures", BindingFlags.Static | BindingFlags.NonPublic).GetValue(null);
                retained.Remove(fixture); process.Dispose(); Directory.Delete(root, true);
            }
        }

        private static object Field(object target, string name) => typeof(OfficeVbeFixture).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(target);
        private static void Set(object target, string name, object value) => typeof(OfficeVbeFixture).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).SetValue(target, value);
        private static void Property(object target, string name, object value) => typeof(OfficeVbeFixture).GetProperty(name, BindingFlags.Instance | BindingFlags.NonPublic).SetValue(target, value);

        public sealed class FakeApplication
        {
            public FakeDocument Document, ActiveDocument;
            public FakeDocuments Documents;
            public FakeEditor Editor = new FakeEditor();
            public FakeEditor VBE => Editor;
            public FakeWindow ActiveWindow => Editor.MainWindow;
            public IDictionary<string, object> Persistence;
            public int BridgeReads, QuitCalls;
            public Action AfterReadback;
            public void Quit() { QuitCalls++; }
        }
        public sealed class FakeDocument { public string FullName; public FakeProject VBProject; public bool ReadOnly, Saved; }
        public sealed class FakeDocuments { public int Count = 1; public FakeDocument Current; public FakeDocument this[int index] => Current; }
        public sealed class FakeProject
        {
            public int Mode = 2, Protection;
            public bool Saved;
            public FakeEditor VBE;
            public FakeComponents Components = new FakeComponents();
            public FakeComponents VBComponents => Components;
        }
        public sealed class FakeComponents
        {
            public int Count => 1;
            public FakeComponent Component = new FakeComponent();
            public FakeComponent Item(object key) => Component;
        }
        public sealed class FakeComponent { public string Name = ModuleName; public FakeCode Code = new FakeCode(); public FakeCode CodeModule => Code; }
        public sealed class FakeCode
        {
            public string Source = OfficeVbeFixturePublisherSaveTests.Source;
            public FakeComponent Parent;
            public FakePane Pane = new FakePane();
            public FakePane CodePane => Pane;
            public int CountOfLines => 2;
            public FakeLines Lines => new FakeLines { Code = this };
        }
        public sealed class FakeLines { public FakeCode Code; public string this[int start, int count] => Code.Source; }
        public sealed class FakePane
        {
            public FakeCode CodeModule;
            public FakeWindow Window;
            public Action AfterShow;
            public int FirstLine = 1;
            public void Show() { AfterShow?.Invoke(); }
            public void SetSelection(int firstLine, int firstColumn, int lastLine, int lastColumn) { FirstLine = firstLine; }
            public void GetSelection(out int firstLine, out int firstColumn, out int lastLine, out int lastColumn)
                { firstLine = FirstLine; firstColumn = lastLine = lastColumn = 1; }
        }
        public sealed class FakeWindow
        {
            public long hWnd = 123, HWnd = 124;
            public bool Visible = true;
            public int Type;
            public Action AfterFocus;
            public void SetFocus() { AfterFocus?.Invoke(); }
        }
        public sealed class FakeEditor
        {
            public FakeWindow Window = new FakeWindow(), MainWindow = new FakeWindow();
            public object ActiveVBProject, ActiveCodePane, ActiveWindow;
            public FakeControl Control = new FakeControl();
            public FakeBars Bars;
            public FakeEditor() { Bars = new FakeBars { Editor = this }; }
            public FakeBars CommandBars => Bars;
        }
        public sealed class FakeBars
        {
            public FakeEditor Editor;
            public int Lookups;
            public Action<int> AfterFind;
            public FakeControl FindControl(int type, int id) { var control = Editor.Control; AfterFind?.Invoke(++Lookups); return control; }
        }
        public sealed class FakeControl
        {
            public int Id = 3, Type = 1, Executions;
            public bool BuiltIn = true, Enabled = true;
            public string OnAction = "", Caption = "Save Host Document";
            public Action AfterExecute;
            public void Execute() { Executions++; AfterExecute?.Invoke(); }
        }
    }
}
