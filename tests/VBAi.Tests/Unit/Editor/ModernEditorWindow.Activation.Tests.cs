using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks native code activation independently from mouse events and browser bootstrap.</summary>
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class ModernEditorActivationTests
    {
        private sealed class ActivationScope : IDisposable
        {
            internal readonly AddInModernEditorFixture Host = new AddInModernEditorFixture();
            internal readonly EditorVbeContract Native = new EditorVbeContract();
            internal readonly ModernEditorWindow Editor;
            internal ActivationScope()
            {
                Native.Vbe.MainWindow.HWnd = Host.Scope.Host.Owner.Handle.ToInt32();
                LlmBoundaryScope.Set(Host.Instance, "vbe", Native.Vbe);
                Editor = Host.Get();
                Editor.ScriptExecution = (method, values) => Task.FromResult(method == "open" ? "1" : "null");
                LlmBoundaryScope.Set(Editor, "<Ready>k__BackingField", true);
                Activate(Native.Original);
            }
            internal void Activate(EditorVbeContract.Component component)
            { Native.Vbe.ActiveCodePane = component.CodeModule.CodePane; Native.Vbe.ActiveWindow = component.CodeModule.CodePane.Window; }
            internal EditorVbeContract.Component Another()
            {
                var component = new EditorVbeContract.Component { Name = Native.Original.Name, Type = 2, Collection = Native.Project.VBComponents };
                component.CodeModule.Raw = "Option Explicit\nPrivate value As Long";
                Native.Project.VBComponents.Items.Add(component); return component;
            }
            internal void Follow() => ModernEditorDebugFixture.Wait(Editor.FollowNativeActivation());
            public void Dispose() => Host.Dispose();
        }

        [STATestMethod]
        public void NativeTransitionUsesComponentIdentityAndPreservesTheLatestRendererDraft()
        {
            using (var f = new ActivationScope())
            {
                f.Follow(); var first = f.Editor.Current;
                var second = f.Another(); string draft = first.Text + "\n' unsynchronized renderer draft";
                f.Editor.ScriptExecution = (method, values) => Task.FromResult(method == "snapshots" ?
                    new JavaScriptSerializer().Serialize(new[] { new { id = first.Id, version = 2, text = draft } }) : method == "open" ? "1" : "null");
                int nativeWrites = 0; f.Native.Original.CodeModule.BeforeOperation = operation => nativeWrites++;
                f.Activate(second); f.Follow();
                Assert.AreEqual(2, f.Editor.Documents.Count(), "Identical mutable display names must not collapse distinct native components.");
                Assert.AreEqual(EditorDocument.Normalize(second.CodeModule.Raw), f.Editor.Current.Text);
                Assert.AreEqual(draft, first.Text); Assert.AreEqual(0, nativeWrites);
                f.Native.Original.Name = "Renamed"; f.Activate(f.Native.Original); f.Follow();
                Assert.AreSame(first, f.Editor.Current); Assert.AreEqual(2, f.Editor.Documents.Count());
            }
        }

        [STATestMethod]
        public void UnchangedNativePaneDoesNotOverrideAnExplicitMonacoTabAndDesignerReturnDoesFollow()
        {
            using (var f = new ActivationScope())
            {
                f.Follow(); var first = f.Editor.Current; var second = f.Another();
                var manual = ModernEditorDebugFixture.Wait(f.Editor.OpenModule(new EditorVbeModule(f.Native.Vbe, f.Native.Project, second)));
                f.Follow(); Assert.AreSame(manual, f.Editor.Current);
                f.Native.Vbe.ActiveWindow = new EditorVbeContract.Window { Type = 1 }; f.Follow();
                f.Activate(f.Native.Original); f.Follow(); Assert.AreSame(first, f.Editor.Current);
            }
        }

        [STATestMethod]
        public void ActivationChangedWhileRendererCaptureIsPendingNeverOpensTheStaleTarget()
        {
            foreach (string change in new[] { "component", "designer", "protected", "break", "project", "window" })
                using (var f = new ActivationScope())
                {
                    var second = f.Another(); var capture = new TaskCompletionSource<string>();
                    int reads = 0; f.Native.Original.CodeModule.BeforeRead = () => reads++;
                    f.Editor.ScriptExecution = (method, values) => method == "snapshots" ? capture.Task : Task.FromResult("1");
                    var follow = f.Editor.FollowNativeActivation(); Application.DoEvents();
                    Assert.IsFalse(follow.IsCompleted, change);
                    if (change == "component") f.Activate(second);
                    if (change == "designer") f.Native.Vbe.ActiveWindow = new EditorVbeContract.Window { Type = 1 };
                    if (change == "protected") f.Native.Project.Protection = 1;
                    if (change == "break") f.Native.Project.Mode = 1;
                    if (change == "project") f.Native.Vbe.ActiveVBProject = new object();
                    if (change == "window") f.Native.Vbe.ActiveWindow = new EditorVbeContract.Window();
                    capture.SetResult("null"); ModernEditorDebugFixture.Wait(follow);
                    Assert.IsNull(f.Editor.Current, change); Assert.AreEqual(0, reads, change);
                }
        }

        [STATestMethod]
        public void PendingCapturePreventsOverlappingFollowAndReadFailureCanRecoverOnTheNextObservation()
        {
            using (var f = new ActivationScope())
            {
                var capture = new TaskCompletionSource<string>(); int captures = 0;
                f.Editor.ScriptExecution = (method, values) => { if (method == "snapshots") { captures++; return capture.Task; } return Task.FromResult("1"); };
                var first = f.Editor.FollowNativeActivation(); Application.DoEvents();
                ModernEditorDebugFixture.Wait(f.Editor.FollowNativeActivation()); Assert.AreEqual(1, captures);
                capture.SetResult("null"); ModernEditorDebugFixture.Wait(first); Assert.IsNotNull(f.Editor.Current);
            }
            using (var f = new ActivationScope())
            {
                f.Native.Original.CodeModule.BeforeRead = () => { throw new InvalidOperationException("Detached transient read failure"); };
                f.Follow(); Assert.IsNull(f.Editor.Current);
                f.Native.Original.CodeModule.BeforeRead = null; f.Follow(); Assert.IsNotNull(f.Editor.Current);
            }
        }

        [STATestMethod]
        public void PassiveActivationRejectsNonCodeModeProtectionAndForeignIdentityBeforeReadingSource()
        {
            foreach (string state in new[] { "designer", "browser", "tool", "break", "running", "protected", "foreign-project", "foreign-window", "component-type", "removed", "not-ready", "busy", "closing", "hidden", "debug-gate" })
                using (var f = new ActivationScope())
                {
                    int reads = 0; f.Native.Original.CodeModule.BeforeRead = () => reads++;
                    if (state == "designer") f.Native.Original.CodeModule.CodePane.Window.Type = 1;
                    if (state == "browser") f.Native.Original.CodeModule.CodePane.Window.Type = 2;
                    if (state == "tool") f.Native.Original.CodeModule.CodePane.Window.Type = 6;
                    if (state == "break") f.Native.Project.Mode = 1;
                    if (state == "running") f.Native.Project.Mode = 0;
                    if (state == "protected") f.Native.Project.Protection = 1;
                    if (state == "foreign-project") f.Native.Vbe.ActiveVBProject = new object();
                    if (state == "foreign-window") f.Native.Vbe.ActiveWindow = new EditorVbeContract.Window();
                    if (state == "component-type") f.Native.Original.Type = 11;
                    if (state == "removed") f.Native.Project.VBComponents.Items.Clear();
                    if (state == "not-ready") LlmBoundaryScope.Set(f.Editor, "<Ready>k__BackingField", false);
                    if (state == "busy") LlmBoundaryScope.Set(f.Editor, "busy", true);
                    if (state == "closing") LlmBoundaryScope.Set(f.Editor, "closing", true);
                    if (state == "hidden") f.Editor.Hide();
                    var gate = UiInvoke.Field<SemaphoreSlim>(f.Editor, "debugCommands"); if (state == "debug-gate") gate.Wait();
                    try { f.Follow(); Assert.IsNull(f.Editor.Current, state); Assert.AreEqual(0, reads, state); }
                    finally { if (state == "debug-gate") gate.Release(); LlmBoundaryScope.Set(f.Editor, "closing", false); }
                }
        }

        /// <summary>Reproduces an empty ready Monaco shell after a later programmatic native code activation.</summary>
        [STATestMethod]
        public void ReadyEmptyWorkspaceFollowsLateProgrammaticCodeActivationWithoutATreeDoubleClick()
        {
            using (var fixture = new AddInModernEditorFixture())
            {
                var native = new EditorVbeContract();
                native.Vbe.MainWindow.HWnd = fixture.Scope.Host.Owner.Handle.ToInt32();
                native.Vbe.ActiveCodePane = null;
                LlmBoundaryScope.Set(fixture.Instance, "vbe", native.Vbe);
                var editor = fixture.Get();
                editor.ScriptExecution = (method, values) => System.Threading.Tasks.Task.FromResult(method == "open" ? "1" : "null");
                LlmBoundaryScope.Set(editor, "<Ready>k__BackingField", true);
                Assert.IsNull(editor.Current);
                native.Vbe.ActiveCodePane = native.Original.CodeModule.CodePane;
                native.Vbe.ActiveWindow = native.Original.CodeModule.CodePane.Window;
                int selectionWrites = native.Vbe.ActiveCodePaneSetCount;
                int visibilityWrites = 0;
                native.Original.CodeModule.CodePane.Window.OnVisible = () => visibilityWrites++;
                UiInvoke.Call(typeof(ModernEditorWindow), "TimerTick", editor, null, EventArgs.Empty);
                var deadline = Stopwatch.StartNew();
                while (editor.Current == null && deadline.ElapsedMilliseconds < 3000)
                { Application.DoEvents(); Thread.Sleep(1); }
                Assert.IsNotNull(editor.Current, "Native code activation must load Monaco after empty bootstrap; no mouse or key event is required.");
                Assert.AreEqual(EditorDocument.Normalize(native.Original.CodeModule.Raw), editor.Current.Text);
                Assert.AreEqual(1, editor.Documents.Count());
                Assert.AreEqual(selectionWrites, native.Vbe.ActiveCodePaneSetCount);
                Assert.AreEqual(0, native.Original.CodeModule.CodePane.Shows);
                Assert.AreEqual(0, native.Original.CodeModule.CodePane.Window.Focuses);
                Assert.AreEqual(0, visibilityWrites, "Following an already active native pane must not show or activate it again.");
            }
        }
    }
}
