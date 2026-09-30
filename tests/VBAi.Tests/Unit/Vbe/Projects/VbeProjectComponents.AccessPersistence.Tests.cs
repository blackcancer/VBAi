using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks deferred Access saved notifications with fake native objects on an owned STA.</summary>
    public sealed partial class VbeOtherHostPersistenceTests
    {
        /// <summary>Reproduces the single Save returning before Access processes its Saved notification.</summary>
        [STATestMethod]
        public void AccessSaveYieldsForDelayedOwnerThreadSavedReadbackWithoutInvokingAgain()
        {
            var f = new AsyncAccessFixture();
            int owner = Thread.CurrentThread.ManagedThreadId;
            using (var ui = new Control())
            {
                _ = ui.Handle;
                f.Probe.AfterInvocation = () => {
                    f.Probe.Observation.Format = 12;
                    f.Probe.Observation.Saved = null;
                    f.Probe.Project.Saved = false;
                    ui.BeginInvoke(new Action(() => {
                        Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                        f.Probe.Project.Saved = true;
                    }));
                };
                var pending = f.Service.SaveHostDocumentAsync(f.Probe.Request());
                Assert.IsFalse(pending.IsCompleted, "Access must yield after its single Save before reading the deferred saved flag.");
                dynamic result = CompleteAccessSave(pending);
                Assert.IsTrue((bool)result.Verified); Assert.IsFalse((bool)result.Uncertain);
                Assert.IsNull((bool?)result.HostSaved);
                Assert.AreEqual(1, f.Probe.Attempts);
                Assert.IsFalse((bool)result.PersistenceReopenVerified);
            }
        }

        /// <summary>Preserves exact context/revision guards throughout the yielded notification window.</summary>
        [STATestMethod]
        [DataRow("source"), DataRow("metadata"), DataRow("references"), DataRow("path")]
        [DataRow("projectpath"), DataRow("projectidentity"), DataRow("documentidentity"), DataRow("pane")]
        [DataRow("component"), DataRow("activeproject"), DataRow("mode"), DataRow("protected")]
        [DataRow("pid"), DataRow("owner"), DataRow("host"), DataRow("readonly"), DataRow("format")]
        [DataRow("missingfile"), DataRow("emptyfile")]
        public void AccessChangesDuringDeferredVerificationRemainUncertainWithoutAnotherSave(string change)
        {
            var f = new AsyncAccessFixture();
            f.Probe.AfterInvocation = () => { f.Probe.Observation.Format = 12; f.Probe.Observation.Saved = null; f.Probe.Project.Saved = false; };
            var pending = f.Service.SaveHostDocumentAsync(f.Probe.Request());
            Assert.IsFalse(pending.IsCompleted);
            switch (change)
            {
                case "source": f.Probe.Component.CodeModule.Source += "' intervening source"; break;
                case "metadata": f.Probe.Project.Description = "Intervening metadata"; break;
                case "references": f.Probe.Project.References.Add(new AsyncAccessReference()); break;
                case "path": f.Probe.Observation.Path = @"C:\fixture\Other.accdb"; break;
                case "projectpath": f.Probe.Project.FileName = @"C:\fixture\Other.accdb"; break;
                case "projectidentity": f.Editor.VBProjects[0] = new OtherProject { Name = "P" }; break;
                case "documentidentity": f.Probe.Items[0] = new object(); break;
                case "pane": f.Editor.ActiveCodePane = new AsyncAccessPane { CodeModule = new AsyncAccessCode { Parent = f.Probe.Component } }; break;
                case "component": f.Editor.ActiveCodePane.CodeModule.Parent = new object(); break;
                case "activeproject": f.Editor.ActiveVBProject = new object(); break;
                case "mode": f.Probe.Project.Mode = 1; break;
                case "protected": f.Probe.Project.Protection = 1; break;
                case "pid": f.Probe.ProcessId++; break;
                case "owner": f.Probe.Owner++; break;
                case "host": f.Probe.Kind = "Word"; break;
                case "readonly": f.Probe.Observation.ReadOnly = true; break;
                case "format": f.Probe.Observation.Format = 99; break;
                case "missingfile": f.Probe.Exists = false; break;
                case "emptyfile": f.Probe.Bytes = 0; break;
            }
            f.Probe.Project.Saved = true;
            dynamic result = CompleteAccessSave(pending);
            Assert.IsFalse((bool)result.Verified, change); Assert.IsTrue((bool)result.Uncertain, change);
            Assert.AreEqual(1, f.Probe.Attempts, change);
        }

        /// <summary>A failed Saved notification expires read-only, clears ownership, and never repeats native Save.</summary>
        [STATestMethod]
        public void AccessDeferredTimeoutAndCrossSessionPendingGuardNeverReplayNativeSave()
        {
            var f = new AsyncAccessFixture();
            f.Service.AccessSaveVerificationTimeout = TimeSpan.FromMilliseconds(80);
            f.Probe.AfterInvocation = () => { f.Probe.Observation.Format = 12; f.Probe.Project.Saved = false; };
            var pending = f.Service.SaveHostDocumentAsync(f.Probe.Request());
            Assert.ThrowsException<InvalidOperationException>(() => CompleteAccessSave(f.Service.SaveHostDocumentAsync(f.Probe.Request())));
            var second = new AsyncAccessFixture();
            Assert.ThrowsException<InvalidOperationException>(() => CompleteAccessSave(second.Service.SaveHostDocumentAsync(second.Probe.Request())));
            Assert.AreEqual(0, second.Probe.Attempts);
            dynamic expired = CompleteAccessSave(pending);
            Assert.IsFalse((bool)expired.Verified); Assert.IsTrue((bool)expired.Uncertain);
            StringAssert.Contains((string)expired.Reason, "timed out");
            Assert.AreEqual(1, f.Probe.Attempts);
            dynamic next = CompleteAccessSave(second.Service.SaveHostDocumentAsync(second.Probe.Request()));
            Assert.IsTrue((bool)next.Verified); Assert.AreEqual(1, second.Probe.Attempts);
        }

        /// <summary>Refuses unsafe approval snapshots before crossing the native mutation boundary.</summary>
        [STATestMethod]
        [DataRow("version"), DataRow("readonly"), DataRow("path"), DataRow("format")]
        [DataRow("owner"), DataRow("pane"), DataRow("selection"), DataRow("protected"), DataRow("mode")]
        public void AccessAsyncPreflightRefusesUnsafeRequestsBeforeSave(string failure)
        {
            var f = new AsyncAccessFixture(); var request = f.Probe.Request();
            switch (failure)
            {
                case "version": request.ExpectedProjectVersion = "stale"; break;
                case "readonly": f.Probe.Observation.ReadOnly = true; break;
                case "path": request.ExpectedHostPath = @"C:\fixture\Other.accdb"; break;
                case "format": f.Probe.Observation.Format = 99; break;
                case "owner": f.Probe.Owner = 999; break;
                case "pane": f.Editor.ActiveCodePane = null; break;
                case "selection": f.Editor.ActiveVBProject = null; break;
                case "protected": f.Probe.Project.Protection = 1; break;
                case "mode": f.Probe.Project.Mode = 1; break;
            }
            Assert.ThrowsException<InvalidOperationException>(() => CompleteAccessSave(f.Service.SaveHostDocumentAsync(request)), failure);
            Assert.AreEqual(0, f.Probe.Attempts, failure);
        }

        /// <summary>Preserves a native error unchanged and refuses the old synchronous mutation route.</summary>
        [STATestMethod]
        public void AccessSynchronousDispatchAndNativeFailuresCannotBypassDeferredContract()
        {
            var f = new AsyncAccessFixture();
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SaveHostDocument(f.Probe.Request()));
            Assert.AreEqual(0, f.Probe.Attempts);
            f.Probe.Failure = "native error";
            dynamic failed = CompleteAccessSave(f.Service.SaveHostDocumentAsync(f.Probe.Request()));
            Assert.IsFalse((bool)failed.Verified); Assert.IsTrue((bool)failed.Uncertain);
            StringAssert.Contains((string)failed.Reason, "Native save failed after invocation");
            Assert.AreEqual(1, f.Probe.Attempts);
        }

        /// <summary>Checks the final built-in command boundary against the exact approved pane and component.</summary>
        [DataTestMethod, DataRow("unchanged"), DataRow("pane"), DataRow("component"), DataRow("closed")]
        public void AccessNativeCommandRechecksApprovedSelectionImmediatelyBeforeItsSingleInvocation(string change)
        {
            var editor = new AsyncAccessEditor();
            var project = new PathProject { VBE = editor, FileName = @"C:\fixture\Owned.accdb" };
            object component = new object();
            editor.VBProjects.Add(project); editor.ActiveVBProject = project;
            editor.ActiveCodePane = new AsyncAccessPane { CodeModule = new AsyncAccessCode { Parent = component } };
            var app = new PathApplication(); var document = new PathDocument { Application = app, FullName = project.FileName };
            app.CurrentProject = document;
            var native = new VbeProjectComponents.NativeOtherHostProbe {
                ReadHostKind = () => "Access", ReadIdentity = ReferenceEquals,
                ReadOwner = window => (uint)Process.GetCurrentProcess().Id,
                ReadActiveApplication = name => app
            };
            native.BindProject(project); native.BindDocument(document);
            native.BindAccessSaveSelection(editor.ActiveCodePane, component);
            native.PrepareSave(document);
            if (change == "pane") editor.ActiveCodePane = new AsyncAccessPane { CodeModule = new AsyncAccessCode { Parent = component } };
            if (change == "component") editor.ActiveCodePane.CodeModule.Parent = new object();
            if (change == "closed") editor.ActiveCodePane = null;
            if (change == "unchanged")
            {
                native.Save(document, false, project.FileName, 12);
                Assert.AreEqual(1, editor.CommandBars.Control.Executions);
            }
            else
            {
                Assert.ThrowsException<InvalidOperationException>(() => native.Save(document, false, project.FileName, 12));
                Assert.AreEqual(0, editor.CommandBars.Control.Executions);
                Assert.IsFalse(native.SaveInvocationStarted);
            }
        }

        /// <summary>Pumps only the detached test dispatcher; no Office application or macro is created.</summary>
        private static object CompleteAccessSave(Task<object> pending)
        {
            var watch = Stopwatch.StartNew();
            while (!pending.IsCompleted && watch.Elapsed.TotalSeconds < 5)
            { Application.DoEvents(); Thread.Sleep(1); }
            Assert.IsTrue(pending.IsCompleted, "Deferred Access verification did not complete.");
            return pending.GetAwaiter().GetResult();
        }

        /// <summary>Retains real service revision checks while replacing the native Office/COM boundary.</summary>
        private sealed class AsyncAccessFixture
        {
            internal readonly Fixture Probe = new Fixture { Kind = "Access" };
            internal readonly AsyncAccessEditor Editor = new AsyncAccessEditor();
            internal readonly VbeProjectComponents Service;
            internal AsyncAccessFixture()
            {
                Probe.Observation.Path = @"C:\fixture\Owned.accdb";
                Probe.Project.FileName = Probe.Observation.Path;
                Probe.Observation.Format = 12; Probe.Observation.Saved = null; Probe.Project.Saved = false;
                Editor.VBProjects.Add(Probe.Project); Editor.ActiveVBProject = Probe.Project;
                Editor.ActiveCodePane = new AsyncAccessPane { CodeModule = new AsyncAccessCode { Parent = Probe.Component } };
                Service = new VbeProjectComponents(Editor, new VbeForms(Editor));
                Service.OtherHostProbe = () => Probe;
                Probe.AfterInvocation = () => { Probe.Observation.Format = 12; Probe.Observation.Saved = null; };
            }
        }

        /// <summary>Simulates selection identity without supplying a different revision implementation.</summary>
        public sealed class AsyncAccessEditor
        {
            public PathWindow MainWindow { get; } = new PathWindow();
            public PathCommands CommandBars { get; } = new PathCommands();
            public List<object> VBProjects { get; } = new List<object>();
            public object ActiveVBProject { get; set; }
            public AsyncAccessPane ActiveCodePane { get; set; }
        }
        /// <summary>Provides only the native pane identity and its selected component association.</summary>
        public sealed class AsyncAccessPane { public AsyncAccessCode CodeModule { get; set; } }
        /// <summary>Identifies the fake native code component selected for Save.</summary>
        public sealed class AsyncAccessCode { public object Parent { get; set; } }

        /// <summary>Changes a native reference revision without evaluating any external library.</summary>
        public sealed class AsyncAccessReference
        {
            public string GUID => "{420B2830-E718-11CF-893D-00A0C9054228}";
            public int Major => 1;
            public int Minor => 0;
            public bool IsBroken => false;
            public bool BuiltIn => false;
        }
    }
}
