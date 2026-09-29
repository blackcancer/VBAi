using System;
using System.Collections.Generic;
using System.Linq;
using VBAi;
using VBAi.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit.Editor
{
    [TestClass]
    public sealed class ModernEditorDebugTests
    {
        [STATestMethod]
        public void ObservationCoversAbsentBusyManagedStableRunningDirtyConflictedAndCleanDocuments()
        {
            foreach (string state in new[] { "absent", "busy", "managed", "stable", "running", "dirty", "conflict", "clean" })
            using (var f = new ModernEditorDebugFixture())
            {
                f.Native.Project.Mode = 1; f.Native.Adapter.ShowNative(3, 2);
                f.Command(1813, "Show Next Statement");
                if (state == "absent") f.Set("selected", null);
                if (state == "busy") f.Set("busy", true);
                if (state == "managed") ModernEditorDebugFixture.Wait(f.Window.OpenModule(f.Storage));
                if (state == "stable") { f.Set("lastDebugMode", 1); f.Set("lastDebugDocument", f.Document.Id); f.Set("lastDebugPosition", "Module1:3:2:3:2"); f.Set("lastExecutionVersion", 1); }
                if (state == "running") f.Native.Project.Mode = 0;
                if (state == "dirty") f.Document.Edit(f.Document.Text + "\n' draft");
                if (state == "conflict") { f.Document.Edit(f.Document.Text + "\n' draft"); f.Native.Original.CodeModule.Raw += "\n' external"; f.Document.Observe(); f.Document.Edit(f.Document.Baseline); Assert.IsTrue(f.Document.Conflict); Assert.IsFalse(f.Document.Dirty); }
                f.Scripts.Clear(); f.Observe();
                Assert.AreEqual(state == "clean" || state == "running" || state == "dirty" || state == "conflict" ? 1 : 0, f.Scripts.Count(x => (x.Item1 == "execution" || x.Item1 == "executionBatch")), state);
            }
        }

        [STATestMethod]
        public void CommandGuardsRefuseBusyMissingManagedStaleDraftAndReformattedBuffers()
        {
            foreach (string state in new[] { "busy", "missing", "null-id", "managed", "stale", "dirty", "conflict", "reformatted" })
            using (var f = new ModernEditorDebugFixture())
            {
                string id = f.Document.Id;
                if (state == "busy") { f.Set("busy", true); f.Window.BeginInvoke(new Action(() => { f.Set("closing", true); f.Set("busy", false); })); }
                if (state == "missing") id = "missing";
                if (state == "managed") id = ModernEditorDebugFixture.Wait(f.Window.OpenModule(f.Storage)).Id;
                if (state == "dirty") f.Document.Edit(f.Document.Text + "\n' draft");
                if (state == "conflict") { f.Document.Edit(f.Document.Text + "\n' draft"); f.Native.Original.CodeModule.Raw += "\n' concurrent"; f.Document.Observe(); }
                if (state == "reformatted")
                {
                    f.Native.Original.CodeModule.Raw += "\n' normalized"; f.Ready(true);
                    f.Rendering = (method, values) => method == "apply" ? "2" : "null";
                }
                if (state == "busy" || state == "missing" || state == "null-id" || state == "managed") f.Send("compile", id: id, nullId: state == "null-id");
                else Assert.ThrowsException<InvalidOperationException>(() => f.Send("toggle_breakpoint", id: id, version: state == "stale" ? 0 : 1), state);
                Assert.AreEqual(0, f.Preflight);
            }
        }

        [STATestMethod]
        public void CompilationCoversWrongModePreflightFailureCommandFailureAndNoDiagnostic()
        {
            foreach (string state in new[] { "run", "break", "preflight", "execute", "null", "whitespace", "reformatted" })
            using (var f = new ModernEditorDebugFixture())
            {
                if (state == "run") f.Native.Project.Mode = 0;
                if (state == "break") f.Native.Project.Mode = 1;
                if (state == "preflight") f.Window.EnsureCompileDialogAbsent = pid => throw new InvalidOperationException("Existing diagnostic");
                var command = f.Command(578, "Compile Project1", () => { if (state == "execute") throw new InvalidOperationException("Compile failed"); });
                if (state == "whitespace") f.Diagnostic = "  ";
                if (state == "reformatted") { f.Native.Original.CodeModule.Raw += "\n' normalized"; f.Ready(true); f.Rendering = (method, values) => method == "apply" ? "2" : "null"; }
                if (state == "run" || state == "break" || state == "preflight" || state == "execute") Assert.ThrowsException<InvalidOperationException>(() => f.Send("compile", version: 1));
                else
                {
                    f.Send("compile", version: 1);
                    Assert.AreEqual(UiText.Get("Compilation finished: no native diagnostics observed. Macros were not executed."), f.Get<System.Windows.Forms.Label>("status").Text, state);
                    Assert.AreEqual(1, command.ExecuteCount); Assert.AreEqual(1, f.Observations);
                    Assert.IsTrue(f.Scripts.Any(x => x.Item1 == "diagnostics" && ((object[])x.Item2[2]).Length == 0));
                }
                Assert.IsFalse(f.Get<bool>("busy"), state);
            }
        }

        [STATestMethod]
        public void CompilationDiagnosticsRejectForeignSelectionsAndNeverUnderlineUnsynchronizedDrafts()
        {
            foreach (string state in new[] { "clean", "dirty", "conflict", "changed", "foreign" })
            using (var f = new ModernEditorDebugFixture())
            {
                f.Diagnostic = "Compile error: fixture"; f.Native.Adapter.ShowNative(3, 2);
                f.Command(578, "Compile Project1", () =>
                {
                    f.Native.Original.CodeModule.CodePane.SetSelection(3, 2, 3, 2);
                    if (state == "dirty") f.Document.Edit(f.Document.Text + "\n' draft");
                    if (state == "conflict") { f.Document.Edit(f.Document.Text + "\n' draft"); f.Native.Original.CodeModule.Raw += "\n' external"; f.Document.Observe(); }
                    if (state == "changed") f.Native.Original.CodeModule.Raw += "\n' compiler change";
                    if (state == "foreign") f.Native.Vbe.ActiveCodePane = new EditorVbeContract.Component { Name = "Module1" }.CodeModule.CodePane;
                });
                if (state == "foreign") Assert.ThrowsException<InvalidOperationException>(() => f.Send("compile"));
                else
                {
                    f.Send("compile");
                    var diagnostics = f.Scripts.Where(x => x.Item1 == "diagnostics").ToArray();
                    Assert.AreEqual(state == "clean" ? 2 : 1, diagnostics.Length, state);
                    if (state == "clean")
                    {
                        Assert.AreEqual(1, ((Array)diagnostics[1].Item2[2]).Length);
                        Assert.IsTrue(f.Scripts.Any(x => x.Item1 == "reveal" && (int)x.Item2[0] == 3 && (int)x.Item2[1] == 2));
                    }
                    else Assert.AreEqual(f.Diagnostic, f.Get<System.Windows.Forms.Label>("status").Text);
                }
                Assert.IsFalse(f.Get<bool>("busy"));
            }
        }

        [STATestMethod]
        public void ExecutionNavigationMarksOnlySelectedDocumentAndRefusesForeignIdentity()
        {
            foreach (bool foreign in new[] { false, true })
            using (var f = new ModernEditorDebugFixture())
            {
                f.Native.Project.Mode = 1; f.Native.Adapter.ShowNative(3, 2);
                var other = new EditorVbeContract.Component { Name = "Other", Type = 1, Collection = f.Native.Project.VBComponents };
                other.CodeModule.Raw = "Debug.Print 2"; f.Native.Project.VBComponents.Items.Add(other);
                ModernEditorDebugFixture.Wait(f.Window.OpenModule(f.Native.Adapter.Sibling("Other")));
                f.Command(1813, "Show Next Statement", () => { if (foreign) f.Native.Vbe.ActiveCodePane = new EditorVbeContract.Component { Name = "Module1" }.CodeModule.CodePane; });
                if (foreign) Assert.ThrowsException<InvalidOperationException>(() => f.Send("show_next_statement"));
                else
                {
                    f.Send("show_next_statement"); var markers = f.Scripts.Where(x => (x.Item1 == "execution" || x.Item1 == "executionBatch")).ToArray();
                    Assert.AreEqual(1, markers.Length); Assert.AreEqual("executionBatch", markers[0].Item1); Assert.AreEqual(f.Document.Id, markers[0].Item2[0]); Assert.AreEqual(3, markers[0].Item2[1]);
                }
                Assert.IsFalse(f.Get<bool>("busy"));
            }
        }

        [STATestMethod]
        public void PaginatedDebugCommandsCoverDisabledWrongCaptionMissingAndEverySupportedAction()
        {
            foreach (string action in new[] { "unknown", "missing", "exhausted", "toggle_breakpoint", "step_into", "step_over", "step_out" })
            using (var f = new ModernEditorDebugFixture())
            {
                f.Native.Project.Mode = action == "toggle_breakpoint" ? 2 : 1;
                int fillers = action == "exhausted" ? 2000 : action == "missing" ? 3 : 200;
                for (int i = 0; i < fillers; i++) f.Command(1000 + i, "Unrelated " + i, enabled: i % 2 == 0);
                string caption = action == "toggle_breakpoint" ? "Toggle Breakpoint" : action == "step_into" ? "Step Into" : action == "step_over" ? "Step Over" : "Step Out";
                var command = action == "unknown" || action == "missing" || action == "exhausted" ? null : f.Command(action == "toggle_breakpoint" ? 51 : 9000, caption);
                f.Set("lastDebugMode", 1);
                if (action == "step_out") { var browser = new Microsoft.Web.WebView2.WinForms.WebView2(); typeof(ModernEditorWindow).GetProperty("Browser", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).SetValue(f.Window, browser); f.Window.Controls.Add(browser); }
                if (action == "unknown") Assert.ThrowsException<ArgumentException>(() => f.Send(action));
                else if (command == null) Assert.ThrowsException<InvalidOperationException>(() => f.Send("step_into"));
                else
                {
                    f.Send(action); Assert.AreEqual(1, command.ExecuteCount);
                    if (action == "toggle_breakpoint") Assert.IsTrue(f.Scripts.Any(x => x.Item1 == "breakpointRequested"));
                    else { Assert.AreEqual(-1, f.Get<int>("lastDebugMode")); Assert.IsTrue(f.Scripts.Any(x => (x.Item1 == "execution" || x.Item1 == "executionBatch") && (int)x.Item2[1] == 0)); }
                }
                Assert.IsFalse(f.Get<bool>("busy"));
            }
        }
        [STATestMethod]
        public void StableExecutionPositionRefreshesOnlyCleanRevisedCachedDocument()
        {
            foreach (string state in new[] { "none", "missing", "dirty", "conflict", "equal", "revised", "null-pane", "unchanged-run" })
            using (var f = new ModernEditorDebugFixture())
            {
                f.Native.Project.Mode = state == "unchanged-run" ? 0 : 1;
                f.Native.Adapter.ShowNative(3, 2);
                f.Set("lastDebugMode", f.Native.Project.Mode);
                f.Set("lastDebugPosition", "Module1:3:2:3:2");
                f.Set("lastDebugDocument", state == "none" ? null : state == "missing" ? "missing" : f.Document.Id);
                f.Set("lastExecutionLine", 3); f.Set("lastExecutionVersion", state == "revised" ? 0 : 1);
                if (state == "dirty") f.Document.Edit(f.Document.Text + "\n' draft");
                if (state == "conflict") { f.Document.Edit(f.Document.Text + "\n' draft"); f.Native.Original.CodeModule.Raw += "\n' external"; f.Document.Observe(); f.Document.Edit(f.Document.Baseline); }
                if (state == "null-pane") { f.Native.Vbe.ActiveCodePane = null; f.Set("lastDebugPosition", ""); }
                f.Scripts.Clear(); f.Observe();
                Assert.AreEqual(state == "revised" ? 1 : 0, f.Scripts.Count, state);
                if (state == "revised") { Assert.AreEqual(3, f.Scripts[0].Item2[1]); Assert.AreEqual(1, f.Get<int>("lastExecutionVersion")); }
            }
        }

        [STATestMethod]
        public void CommandWaitHonorsLateClosingDisposalAndSuccessfulRelease()
        {
            foreach (string state in new[] { "closing", "disposed", "late-closing", "late-disposed", "released" })
            using (var f = new ModernEditorDebugFixture())
            {
                if (state == "closing") f.Set("closing", true);
                if (state == "disposed") f.Window.Dispose();
                if (state.StartsWith("late-", StringComparison.Ordinal) || state == "released")
                {
                    f.Set("busy", true);
                    f.Window.BeginInvoke(new Action(() => { if (state == "late-closing") f.Set("closing", true); if (state == "late-disposed") f.Window.Dispose(); f.Set("busy", false); }));
                }
                var command = f.Command(578, "Compile Project1");
                f.Send("compile"); Assert.AreEqual(state == "released" ? 1 : 0, command.ExecuteCount, state);
            }
        }

        [STATestMethod]
        public void ExecutionNavigationReusesCurrentTabAndOpensPreviouslyUnseenSelection()
        {
            foreach (bool unopened in new[] { false, true })
            using (var f = new ModernEditorDebugFixture())
            {
                f.Native.Project.Mode = 1; f.Native.Adapter.ShowNative(3, 2);
                if (unopened)
                {
                    ModernEditorDebugFixture.Wait(f.Window.OpenModule(f.Storage));
                    var other = new EditorVbeContract.Component { Name = "Unopened", Type = 1, Collection = f.Native.Project.VBComponents };
                    other.CodeModule.Raw = "Debug.Print 2"; f.Native.Project.VBComponents.Items.Add(other);
                    other.CodeModule.CodePane.SetSelection(1, 1, 1, 1);
                    f.Command(1813, "Show Next Statement", () => f.Native.Vbe.ActiveCodePane = other.CodeModule.CodePane);
                }
                else f.Command(1813, "Show Next Statement");
                f.Send("show_next_statement");
                Assert.AreEqual(unopened ? 3 : 1, f.Get<Dictionary<string, EditorDocument>>("documents").Count);
                Assert.AreEqual(unopened ? 1 : 3, f.Get<int>("lastExecutionLine"));
                Assert.AreEqual(1, f.Scripts.Count(x => (x.Item1 == "execution" || x.Item1 == "executionBatch") && (int)x.Item2[1] != 0));
            }
        }
    }
}
