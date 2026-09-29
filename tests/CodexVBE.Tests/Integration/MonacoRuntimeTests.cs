using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Integration
{
    [TestClass, TestCategory("MonacoRuntime")]
    public sealed class MonacoRuntimeTests : EditorUiTestFixture
    {
        internal static void Wait(Func<bool> complete, int seconds = 40)
        { var clock = Stopwatch.StartNew(); while (!complete() && clock.Elapsed.TotalSeconds < seconds) { Application.DoEvents(); EditorUiTestFixture.ThrowIfUiFailed(); Thread.Sleep(15); } EditorUiTestFixture.ThrowIfUiFailed(); Assert.IsTrue(complete(), "Timed out waiting for real WebView2/Monaco."); }
        internal static T Wait<T>(Task<T> task) { Wait(() => task.IsCompleted); return task.GetAwaiter().GetResult(); }
        internal static void Wait(Task task) { Wait(() => task.IsCompleted); task.GetAwaiter().GetResult(); }
        [STATestMethod]
        public void ClosingDuringBrowserInitializationPreservesDraftWithoutNativeWrites()
        {
            using (var host = new EditorFixture())
            using (var window = new ModernEditorWindow { Drafts = new EditorDraftStore(host.Root) })
            {
                var document = Wait(window.OpenModule(host));
                document.Edit(document.Text + "\n' recovery during initialization");
                window.Show();
                window.Close();
                Wait(() => window.IsDisposed);
                Assert.AreEqual(0, host.Writes);
                Assert.AreEqual(document.Text, new EditorDraftStore(host.Root).Recover(host.Key).Text);
            }
        }

        [STATestMethod]
        public void ClosingDuringConflictButtonHandleCreationWaitsForStatusLayout()
        {
            using (var host = new EditorFixture())
            using (var window = new ModernEditorWindow { Drafts = new EditorDraftStore(host.Root) })
            {
                window.Show(); Wait(() => window.Ready);
                ((System.Windows.Forms.Timer)typeof(ModernEditorWindow).GetField("timer", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).GetValue(window)).Stop();
                var document = Wait(window.OpenModule(host));
                var field = typeof(ModernEditorWindow).GetField("compare", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
                var compare = (Button)field.GetValue(window);
                Assert.IsFalse(compare.IsHandleCreated, "The conflict button must still be hidden and uncreated.");
                bool closeRequested = false;
                compare.HandleCreated += (sender, args) => { closeRequested = true; window.Close(); };
                Wait(window.Script("insert", "' local draft\n"));
                Wait(() => document.Dirty);
                host.Code += "\n' concurrent native change";
                Wait(window.ProcessDocuments(false));
                Wait(() => closeRequested);
                Wait(() => window.IsDisposed);
                Assert.AreEqual(0, host.Writes);
                Assert.AreEqual(document.Text, new EditorDraftStore(host.Root).Recover(host.Key).Text);
            }
        }

        [STATestMethod]
        public void DefinitionNavigationSelectsTheTargetModuleTab()
        {
            using (var first = new EditorFixture { DisplayName = "Caller", Code = "Public Sub Main()\n ExternalProc\nEnd Sub" })
            using (var second = new EditorFixture { DisplayName = "Library", Code = "Public Sub ExternalProc()\nEnd Sub" })
            using (var window = new ModernEditorWindow())
            {
                window.Drafts = new EditorDraftStore(first.Root);
                var caller = Wait(window.OpenModule(first)); var target = Wait(window.OpenModule(second));
                window.Show(); Wait(() => window.Ready); Wait(() => Wait(window.Script("snapshots")).Contains(target.Id));
                Wait(window.OpenModule(first));
                Wait(window.Script("reveal", 2, 5)); Wait(window.Script("command", "editor.action.revealDefinition"));
                Wait(() => window.Current == target);
                Wait(() => Wait(window.Script("position")).Contains("\"line\":1"));
                window.Close(); Wait(() => window.IsDisposed);
            }
        }
        [STATestMethod]
        public void LanguageProvidersReadTheCurrentUnsavedSnapshot()
        {
            using (var host = new EditorFixture())
            using (var window = new ModernEditorWindow())
            {
                host.Code = "Option Explicit\nPublic Function Compute(ByVal value As Long) As Long\n Dim localValue As Long\n Compute = value\nEnd Function\nPublic Sub Elsewhere()\n Dim hiddenLocal As String\nEnd Sub";
                window.Drafts = new EditorDraftStore(host.Root);
                var doc = Wait(window.OpenModule(host)); window.Show(); Wait(() => window.Ready);
                Wait(() => Wait(window.Script("snapshots")).Contains(doc.Id));
                UiInvoke.Field<System.Windows.Forms.Timer>(window, "timer").Stop();
                Wait(window.Browser.CoreWebView2.ExecuteScriptAsync("window.languageResult=null; window.vbai.languageInspect('" + doc.Id + "',4,2).then(value=>window.languageResult=value);"));
                Wait(() => Wait(window.Browser.CoreWebView2.ExecuteScriptAsync("window.languageResult")) != "null");
                string result = Wait(window.Browser.CoreWebView2.ExecuteScriptAsync("window.languageResult"));
                StringAssert.Contains(result, "localValue"); StringAssert.Contains(result, "value As Long");
                Assert.IsFalse(result.Contains("hiddenLocal"));
                Wait(window.Script("reveal", 4, 14));
                Wait(window.Script("command", "editor.action.revealDefinition"));
                Wait(() => Wait(window.Script("position")).Contains("\"line\":2"));
                window.Close(); Wait(() => window.IsDisposed);
            }
        }
        [STATestMethod]
        public void RealRendererEditsSynchronizesRejectsStaleReplacementsAndDisplaysDiff()
        {
            using (var host = new EditorFixture())
            using (var window = new ModernEditorWindow())
            {
                window.Drafts = new EditorDraftStore(host.Root);
                var doc = Wait(window.OpenModule(host)); window.Show();
                Wait(() => window.Ready); UiInvoke.Field<System.Windows.Forms.Timer>(window, "timer").Stop();
                Wait(() => Wait(window.Script("snapshots")).Contains(doc.Id)); UiInvoke.Field<System.Windows.Forms.Timer>(window, "timer").Stop();
                var snapshot = new System.Web.Script.Serialization.JavaScriptSerializer().Deserialize<dynamic[]>(Wait(window.Script("snapshots")));
                int version = (int)snapshot[0]["version"];
                var markers = new[] { new { message = "Fixture compile diagnostic", startLineNumber = 3, startColumn = 1, endLineNumber = 3, endColumn = 2, severity = 8 } };
                Wait(window.Script("diagnostics", doc.Id, version - 1, markers));
                StringAssert.Contains(Wait(window.Script("testInfo")), "\"markers\":0");
                Wait(window.Script("diagnostics", doc.Id, version, markers));
                StringAssert.Contains(Wait(window.Script("testInfo")), "\"markers\":1");
                string edited = host.Code.Replace("Print 1", "Print 42");
                string result = Wait(window.Script("apply", doc.Id, version, edited)); Assert.AreNotEqual("0", result);
                StringAssert.Contains(Wait(window.Script("testInfo")), "\"markers\":0");
                Assert.AreEqual("0", Wait(window.Script("apply", doc.Id, version, "stale text")));
                Wait(window.ProcessDocuments(true)); Assert.AreEqual(edited, host.Code); Assert.IsFalse(doc.Dirty);
                Wait(window.Script("theme", true, false)); StringAssert.Contains(Wait(window.Script("testInfo")), "30, 34, 42");
                Wait(window.Script("compare", "Option Explicit")); StringAssert.Contains(Wait(window.Script("testInfo")), "\"diff\":true");
                Wait(window.Script("hideDiff"));
                Wait(window.Script("reveal", 4, 8));
                Wait(window.Script("insert", "\n' actual editor input"));
                Wait(() => doc.Text.Contains("actual editor input"));
                Wait(window.ProcessDocuments(true)); StringAssert.Contains(host.Code, "actual editor input");
                Wait(window.Script("theme", false, false)); StringAssert.Contains(Wait(window.Script("testInfo")), "255, 255, 255");
                string actual = Wait(window.Browser.CoreWebView2.ExecuteScriptAsync("window.vbai.testInfo().language")); Assert.AreEqual("\"vba\"", actual);
                string action = Wait(window.Browser.CoreWebView2.ExecuteScriptAsync("window.vbai.command('actions.find'); true")); Assert.AreEqual("true", action);
                var folder = Environment.GetEnvironmentVariable("VBAI_EDITOR_EVIDENCE");
                if (!string.IsNullOrEmpty(folder))
                {
                    Directory.CreateDirectory(folder);
                    using (var file = File.Create(Path.Combine(folder, "monaco-light.png"))) Wait(window.Browser.CoreWebView2.CapturePreviewAsync(CoreImageFormat(), file));
                    Wait(window.Script("theme", true, false));
                    using (var file = File.Create(Path.Combine(folder, "monaco-dark.png"))) Wait(window.Browser.CoreWebView2.CapturePreviewAsync(CoreImageFormat(), file));
                }

                UiInvoke.Field<System.Windows.Forms.Timer>(window, "timer").Start();
                Wait(window.Script("insert", "\n' continuous synchronization")); Wait(() => host.Code.Contains("continuous synchronization"));
                UiInvoke.Field<System.Windows.Forms.Timer>(window, "timer").Stop();
                Wait(window.Script("insert", "\n' resolve local")); Wait(() => doc.Dirty);
                host.Code += "\n' native concurrent"; Wait(window.ProcessDocuments(false)); Assert.IsTrue(doc.Conflict);
                 Wait(() => UiInvoke.Field<Button>(window, "compare").Visible && !UiInvoke.Field<bool>(window, "busy")); UiInvoke.Field<Button>(window, "compare").PerformClick();
                Wait(() => UiInvoke.Field<Button>(window, "resolve").Enabled);
                Wait(() => UiInvoke.Field<Button>(window, "resolve").Visible && !UiInvoke.Field<bool>(window, "busy")); UiInvoke.Field<Button>(window, "resolve").PerformClick();  Wait(() => !doc.Dirty && !UiInvoke.Field<bool>(window, "busy"));

                StringAssert.Contains(host.Code, "resolve local"); Assert.IsFalse(doc.Conflict);
                Wait(window.Script("insert", "\n' reload recovery")); Wait(() => doc.Dirty);
                host.Code += "\n' another native change"; Wait(window.ProcessDocuments(false));
                Wait(() => UiInvoke.Field<Button>(window, "reload").Visible && !UiInvoke.Field<bool>(window, "busy")); UiInvoke.Field<Button>(window, "reload").PerformClick(); Wait(() => !doc.Dirty);
                Wait(window.ProcessDocuments(false));
                StringAssert.Contains(window.Drafts.Recover(doc.RecoveryKey).Text, "reload recovery");
                Wait(() => UiInvoke.Field<Button>(window, "restore").Visible && !UiInvoke.Field<bool>(window, "busy")); UiInvoke.Field<Button>(window, "restore").PerformClick(); Wait(() => doc.Text.Contains("reload recovery") && !UiInvoke.Field<bool>(window, "busy"));

                Wait(() => !UiInvoke.Field<bool>(window, "busy")); UiInvoke.Call(typeof(ModernEditorWindow), "CloseModuleClick", window, window, EventArgs.Empty); Wait(() => !System.Linq.Enumerable.Any(window.Documents));
                StringAssert.Contains(window.Drafts.Recover(doc.RecoveryKey).Text, "reload recovery");

                window.Close(); Wait(() => window.IsDisposed);
            }
        }
        [STATestMethod]
        public void BreakpointRequestsToggleAndClearOnSourceChanges()
        {
            using (var host = new EditorFixture())
            using (var window = new ModernEditorWindow())
            {
                window.Drafts = new EditorDraftStore(host.Root);
                var doc = Wait(window.OpenModule(host)); window.Show(); Wait(() => window.Ready);
                Wait(() => Wait(window.Script("snapshots")).Contains(doc.Id));
                Wait(window.Script("breakpointRequested", doc.Id, 3));
                Wait(window.Script("breakpointRequested", doc.Id, 3));
                StringAssert.Contains(Wait(window.Script("testInfo")), "\"pendingBreakpoints\":0");
                Wait(window.Script("breakpointRequested", doc.Id, 3));
                StringAssert.Contains(Wait(window.Script("testInfo")), "\"pendingBreakpoints\":1");
                Wait(window.Script("insert", "'changed\n"));
                StringAssert.Contains(Wait(window.Script("testInfo")), "\"pendingBreakpoints\":0");
                window.Close(); Wait(() => window.IsDisposed);
            }
        }
        [STATestMethod]
        public void ClosingATabKeepsItsOriginalTargetAcrossAnAwait()
        {
            using (var first = new EditorFixture())
            using (var second = new EditorFixture())
            using (var window = new ModernEditorWindow())
            {
                window.Drafts = new EditorDraftStore(first.Root);
                var a = Wait(window.OpenModule(first)); window.Show(); Wait(() => window.Ready);
                var b = Wait(window.OpenModule(second)); Wait(window.OpenModule(first));
                UiInvoke.Field<System.Windows.Forms.Timer>(window, "timer").Stop();
                UiInvoke.Call(typeof(ModernEditorWindow), "CloseModuleClick", window, window, EventArgs.Empty);
                Wait(window.OpenModule(second));
                Wait(() => !UiInvoke.Field<bool>(window, "busy"));
                Assert.AreSame(b, window.Current);
                Assert.AreEqual(1, System.Linq.Enumerable.Count(window.Documents));
                window.Close(); Wait(() => window.IsDisposed);
            }
        }
        public sealed class WorkspaceVbe
        {
            public WorkspaceMain MainWindow { get; set; }
            public WorkspaceWindow ActiveWindow { get; } = new WorkspaceWindow();
        }
        public sealed class WorkspaceMain { public int HWnd { get; set; } }
        public sealed class WorkspaceWindow { public int Type { get; set; } }
        [STATestMethod]
        public void WorkspaceFillsDocumentAreaAndLeavesNativeDesignersAccessible()
        {
            using (var fixture = new EditorFixture())
            using (var desktop = new Form { IsMdiContainer = true, Width = 1000, Height = 700 })
            using (var window = new ModernEditorWindow())
            {
                desktop.Show();
                var vbe = new WorkspaceVbe { MainWindow = new WorkspaceMain { HWnd = desktop.Handle.ToInt32() } };
                using (var host = new EditorWorkspaceHost(vbe, window))
                {
                    host.Show(); Wait(() => window.Ready);
                    var mdi = System.Linq.Enumerable.Single(System.Linq.Enumerable.OfType<MdiClient>(desktop.Controls));
                    Assert.AreEqual(mdi.ClientSize, window.Size);
                    desktop.Width = 800; Wait(() => window.Size == mdi.ClientSize);
                    vbe.ActiveWindow.Type = 1; Wait(() => !window.Visible);
                    vbe.ActiveWindow.Type = 7; UiInvoke.Call(typeof(EditorWorkspaceHost), "Resize", host); Assert.IsFalse(window.Visible);
                    vbe.ActiveWindow.Type = 0; Wait(() => window.Visible);
                    window.Close(); Assert.IsFalse(window.IsDisposed);
                }
            }
        }
        private static Microsoft.Web.WebView2.Core.CoreWebView2CapturePreviewImageFormat CoreImageFormat() => Microsoft.Web.WebView2.Core.CoreWebView2CapturePreviewImageFormat.Png;
    }
}
