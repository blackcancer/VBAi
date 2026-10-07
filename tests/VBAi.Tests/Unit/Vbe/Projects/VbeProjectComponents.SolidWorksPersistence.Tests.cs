using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbeSolidWorksPersistenceTests
    {
        internal const string PendingSaveMessage = "A SOLIDWORKS save is already awaiting verification; no second save was invoked.";
        public sealed class HostProject : VbeProjectComponentsTests.FakeProject
        { public int Type { get; set; } = 100; public int Protection { get; set; } }
        internal sealed class Probe : VbeProjectComponents.ISolidWorksSaveProbe
        {
            internal readonly HostProject Project = new HostProject { Name = "P", FileName = @"C:\fixture\Owned.swp", Saved = false };
            internal readonly VbeProjectComponentsTests.FakeComponent Component = new VbeProjectComponentsTests.FakeComponent("M", 1);
            internal readonly VbeProjectComponents Service;
            internal bool Host = true, Owner = true, Identity = true, Selected = true, Exists = true, ReadOnly;
            internal long Bytes = 42;
            internal int Saves, Selections, Pid = 42;
            private readonly int ownerThread = Thread.CurrentThread.ManagedThreadId;
            internal Action OnSelect, OnSave;
            internal Probe()
            {
                var vbe = new VbeProjectComponentsTests.FakeVbe();
                vbe.VBProjects.Add(Project); Project.VBComponents.Add(Component);
                Service = new VbeProjectComponents(vbe, new VbeForms(vbe));
                Service.SolidWorksSaveProbe = () => this;
                Service.SolidWorksSaveVerificationTimeout = TimeSpan.FromMilliseconds(120);
            }
            internal Request Request()
            {
                dynamic state = Service.ProjectProperties("P");
                return new Request
                {
                    Project = "P",
                    ExpectedProjectVersion = state.Version,
                    ExpectedHostPath = @"C:\fixture\Owned.swp",
                    Path = @"C:\fixture\Different.swp"
                };
            }
            // Trace admission precedes VbeUiTask.Run, and the original task stays owned by this test STA.
            internal Task<object> SaveHostAsync(Request request) =>
                SolidWorksStaTestMethodAttribute.StartSave(() => Service.SaveHostDocumentAsync(request));
            internal Task<object> SaveAdapterAsync(Request request) =>
                SolidWorksStaTestMethodAttribute.StartSave(() => Service.SaveSolidWorksMacroAsync(request, this));
            public bool IsSolidWorks => Host;
            public int ProcessId => Pid;
            public VbeProjectComponents.SolidWorksSaveSelection Selection { get; } = new VbeProjectComponents.SolidWorksSaveSelection();
            public void RestoreSelection(object editor, object project, object component) { }
            public void RequireOwner(object editor) { Assert.AreEqual(ownerThread, Thread.CurrentThread.ManagedThreadId); if (!Host || !Owner) throw new InvalidOperationException("Wrong PID/thread"); }
            public bool SameProject(object first, object second) => Identity && ReferenceEquals(first, second);
            public object SelectComponent(object editor, object project) { Selections++; OnSelect?.Invoke(); return Component; }
            public bool SelectionMatches(object editor, object project, object component) => Selected && Identity && ReferenceEquals(project, Project) && ReferenceEquals(component, Component);
            public object SaveControl(object editor) => this;
            public void Save(object control) { Saves++; Project.Saved = true; OnSave?.Invoke(); }
            public bool FileExists(string path) => Exists;
            public bool FileReadOnly(string path) => ReadOnly;
            public long FileLength(string path) => Bytes;
        }
        public sealed class SelectionProject
        {
            public int Mode { get; set; } = 2;
            public System.Collections.Generic.List<SelectionComponent> VBComponents { get; } = new System.Collections.Generic.List<SelectionComponent>();
        }
        public sealed class SelectionComponent
        {
            public int Type => 1;
            public SelectionCode CodeModule { get; }
            public SelectionProject Owner { get; }
            public SelectionComponent(SelectionProject owner)
            { Owner = owner; CodeModule = new SelectionCode(this); owner.VBComponents.Add(this); }
        }
        public sealed class SelectionCode
        {
            public SelectionComponent Parent { get; }
            public SelectionPane CodePane { get; }
            public SelectionCode(SelectionComponent parent) { Parent = parent; CodePane = new SelectionPane(this); }
        }
        public sealed class SelectionPane
        {
            public SelectionCode CodeModule { get; }
            public SelectionWindow Window { get; } = new SelectionWindow();
            public int Shows;
            public SelectionPane(SelectionCode code) { CodeModule = code; }
            public void Show() { Shows++; }
        }
        public sealed class SelectionWindow { public int Focuses; public void SetFocus() { Focuses++; } }
        public sealed class SelectionEditor
        {
            private SelectionPane active;
            public SelectionProject ActiveVBProject { get; private set; }
            public SelectionPane ActiveCodePane { get => active; set { active = value; ActiveVBProject = value?.CodeModule.Parent.Owner; } }
        }
        private static VbeProjectComponents.ISolidWorksSaveProbe NativeSelectionProbe()
        {
            var type = typeof(VbeProjectComponents).GetNestedType("NativeSolidWorksSaveProbe", System.Reflection.BindingFlags.NonPublic);
            return (VbeProjectComponents.ISolidWorksSaveProbe)Activator.CreateInstance(type, true);
        }
        [TestMethod]
        public void NativeSelectionKeepsActiveSecondModuleWithoutShowingOrFocusingFirstModule()
        {
            var project = new SelectionProject();
            var first = new SelectionComponent(project); var second = new SelectionComponent(project);
            var editor = new SelectionEditor { ActiveCodePane = second.CodeModule.CodePane };
            var native = NativeSelectionProbe();
            Assert.AreSame(second, native.SelectComponent(editor, project));
            Assert.AreSame(second.CodeModule.CodePane, editor.ActiveCodePane);
            Assert.IsTrue(native.SelectionMatches(editor, project, second));
            Assert.IsFalse(native.Selection.Changed);
            Assert.AreEqual(0, first.CodeModule.CodePane.Shows); Assert.AreEqual(0, second.CodeModule.CodePane.Shows);
            Assert.AreEqual(0, second.CodeModule.CodePane.Window.Focuses);
        }
        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void NativeSelectionRestoresOnlyItsUnchangedSelectionWithoutForcingFocus(bool userChangedSelection)
        {
            var target = new SelectionProject(); var targetModule = new SelectionComponent(target);
            var previous = new SelectionComponent(new SelectionProject());
            var intervening = new SelectionComponent(new SelectionProject());
            var editor = new SelectionEditor { ActiveCodePane = previous.CodeModule.CodePane };
            var native = NativeSelectionProbe();
            Assert.AreSame(targetModule, native.SelectComponent(editor, target));
            if (userChangedSelection) editor.ActiveCodePane = intervening.CodeModule.CodePane;
            native.RestoreSelection(editor, target, targetModule);
            Assert.AreSame(userChangedSelection ? intervening.CodeModule.CodePane : previous.CodeModule.CodePane, editor.ActiveCodePane, native.Selection.Reason);
            Assert.AreEqual(!userChangedSelection, native.Selection.Restored);
            Assert.AreEqual(0, previous.CodeModule.CodePane.Shows); Assert.AreEqual(0, previous.CodeModule.CodePane.Window.Focuses);
        }
        [SolidWorksStaTestMethod]
        public void ExistingType100SaveUsesSeparateAdapterAndNeverClaimsReloadOrAllowsSaveAs()
        {
            var p = new Probe(); var request = p.Request();
            dynamic before = p.Service.PersistenceStatus("P");
            Assert.IsTrue((bool)before.HostAvailable);
            Assert.AreEqual("VBE.CommandBars.ID3", (string)before.SaveApi);
            Assert.ThrowsException<InvalidOperationException>(() => p.Service.SaveHostDocument(request));
            Assert.AreEqual(0, p.Saves, "Synchronous calls must refuse before the queued native mutation.");
            dynamic result = Complete(p.SaveHostAsync(request));
            Assert.IsTrue((bool)result.Verified); Assert.IsFalse((bool)result.Uncertain);
            Assert.IsTrue((bool)result.SaveInvoked); Assert.IsFalse((bool)result.SaveAsInvoked);
            Assert.IsFalse((bool)result.PersistenceReopenVerified);
            Assert.AreEqual(1, p.Saves);
            Assert.ThrowsException<InvalidOperationException>(() => p.Service.SaveHostDocumentAs(p.Request()));
            Assert.AreEqual(1, p.Saves);
        }
        [SolidWorksStaTestMethod]
        [DataRow("host")]
        [DataRow("owner")]
        [DataRow("type")]
        [DataRow("path")]
        [DataRow("extension")]
        [DataRow("missing")]
        [DataRow("readonly")]
        [DataRow("empty")]
        [DataRow("protected")]
        [DataRow("mode")]
        [DataRow("version")]
        [DataRow("selection changed")]
        [DataRow("pid changed")]
        [DataRow("code changed")]
        public void RefusalBeforeSaveInvokesNoMutation(string failure)
        {
            var p = new Probe(); var request = p.Request();
            if (failure == "host") p.Host = false;
            if (failure == "owner") p.Owner = false;
            if (failure == "type") p.Project.Type = 101;
            if (failure == "path") p.Project.FileName = @"C:\fixture\Other.swp";
            if (failure == "extension") p.Project.FileName = @"C:\fixture\Other.xlsm";
            if (failure == "missing") p.Exists = false;
            if (failure == "readonly") p.ReadOnly = true;
            if (failure == "empty") p.Bytes = 0;
            if (failure == "protected") p.Project.Protection = 1;
            if (failure == "mode") p.Project.Mode = 1;
            if (failure == "version") request.ExpectedProjectVersion = "stale";
            if (failure == "selection changed") p.OnSelect = () => p.Selected = false;
            if (failure == "pid changed") p.OnSelect = () => p.Pid++;
            if (failure == "code changed") p.OnSelect = () => p.Component.CodeModule.Source += "'changed";
            var refusal = Assert.ThrowsException<InvalidOperationException>(() => Complete(p.SaveAdapterAsync(request)));
            Assert.AreNotEqual(PendingSaveMessage, refusal.Message, "An unfinished earlier scenario cannot satisfy this preflight oracle.");
            Assert.AreEqual(0, p.Saves);
        }
        [SolidWorksStaTestMethod]
        [DataRow("native error")]
        [DataRow("unsaved")]
        [DataRow("missing")]
        [DataRow("empty")]
        [DataRow("path")]
        [DataRow("code")]
        [DataRow("identity")]
        [DataRow("owner")]
        [DataRow("mode")]
        public void PostInvocationFailureRemainsUncertainAndIsNeverRetried(string failure)
        {
            var p = new Probe(); var request = p.Request();
            p.OnSave = () =>
            {
                if (failure == "native error") throw new InvalidOperationException("Native save returned an error after invocation");
                if (failure == "unsaved") p.Project.Saved = false;
                if (failure == "missing") p.Exists = false;
                if (failure == "empty") p.Bytes = 0;
                if (failure == "path") p.Project.FileName = @"C:\fixture\Other.swp";
                if (failure == "code") p.Component.CodeModule.Source += "'changed";
                if (failure == "identity") p.Identity = false;
                if (failure == "owner") p.Owner = false;
                if (failure == "mode") p.Project.Mode = 1;
            };
            dynamic result = Complete(p.SaveAdapterAsync(request));
            Assert.AreEqual(1, p.Saves); Assert.IsTrue((bool)result.MutationInvoked);
            Assert.IsTrue((bool)result.Uncertain); Assert.IsFalse((bool)result.Verified);
            StringAssert.Contains((string)result.Next, "do not retry");
            if (failure == "unsaved") StringAssert.Contains((string)result.Reason, "saved flag");
            if (failure == "path") StringAssert.Contains((string)result.Reason, "host path");
            if (failure == "code") StringAssert.Contains((string)result.Reason, "live VBA source");
        }

        [SolidWorksStaTestMethod]
        public void QueuedSaveCompletesOnOwnerThreadWithOneMutationAndRejectsOverlappingSave()
        {
            var p = new Probe(); var request = p.Request();
            int owner = Thread.CurrentThread.ManagedThreadId;
            using (var ui = new Control())
            {
                var handle = ui.Handle;
                p.OnSave = () =>
                {
                    p.Project.Saved = false;
                    ui.BeginInvoke(new Action(() =>
                    {
                        Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                        p.Project.Saved = true;
                    }));
                };
                var pending = p.SaveHostAsync(request);
                Assert.IsFalse(pending.IsCompleted, "Native save must yield before observing completion.");
                Assert.AreEqual(1, p.Saves);
                Assert.AreEqual(PendingSaveMessage,
                    Assert.ThrowsException<InvalidOperationException>(() => Complete(p.SaveHostAsync(request))).Message);
                var anotherSession = new Probe();
                Assert.AreEqual(PendingSaveMessage,
                    Assert.ThrowsException<InvalidOperationException>(() => Complete(anotherSession.SaveHostAsync(anotherSession.Request()))).Message);
                Assert.AreEqual(0, anotherSession.Saves, "The bridge and chat must share the pending-save gate.");
                dynamic result = Complete(pending);
                Assert.IsTrue((bool)result.Verified); Assert.IsFalse((bool)result.Uncertain);
                Assert.AreEqual("DeferredOwnerThreadSavedReadback", (string)result.Verification);
                Assert.AreEqual(1, p.Saves);
            }
        }

        [SolidWorksStaTestMethod]
        [DataRow("code")]
        [DataRow("path")]
        [DataRow("identity")]
        [DataRow("selection")]
        [DataRow("mode")]
        [DataRow("owner")]
        [DataRow("pid")]
        [DataRow("metadata")]
        [DataRow("protected")]
        public void ChangesWhileYieldedNeverBecomeSuccessfulOrReplaySave(string change)
        {
            var p = new Probe();
            p.OnSave = () => p.Project.Saved = false;
            var pending = p.SaveHostAsync(p.Request());
            Assert.IsFalse(pending.IsCompleted);
            if (change == "code") p.Component.CodeModule.Source += "'intervening edit";
            if (change == "path") p.Project.FileName = @"C:\fixture\Changed.swp";
            if (change == "identity") p.Identity = false;
            if (change == "selection") p.Selected = false;
            if (change == "mode") p.Project.Mode = 1;
            if (change == "owner") p.Owner = false;
            if (change == "pid") p.Pid++;
            if (change == "metadata") p.Project.Description = "Changed during save";
            if (change == "protected") p.Project.Protection = 1;
            p.Project.Saved = true;
            dynamic result = Complete(pending);
            Assert.IsFalse((bool)result.Verified); Assert.IsTrue((bool)result.Uncertain);
            Assert.AreEqual(1, p.Saves);
        }

        internal static object Complete(Task<object> pending)
        {
            var watch = Stopwatch.StartNew();
            while (!pending.IsCompleted && watch.Elapsed.TotalSeconds < 5)
            { Application.DoEvents(); Thread.Sleep(1); }
            if (!pending.IsCompleted) Assert.Fail("Owner-thread save verification did not finish. " + SolidWorksStaTestMethodAttribute.Describe(pending));
            return pending.GetAwaiter().GetResult();
        }
    }
}
