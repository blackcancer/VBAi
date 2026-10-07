using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Threading.Tasks;
using VBAi.Tests.Infrastructure;
namespace VBAi.Tests.Unit.Editor
{
    [TestClass, TestCategory("Unit")]
    public sealed class ModernEditorToolTests
    {
        private static void Wait(Task task) => ModernEditorDebugFixture.Wait(task);
        private static T Wait<T>(Task<T> task) => ModernEditorDebugFixture.Wait(task);
        [STATestMethod]
        public void CaptureAndRevisionGuardsRejectEachUnavailableOwnedStateAndReleaseBusyOnFailure()
        {
            foreach (string state in new[] { "busy", "closing", "disposed", "not-ready", "error", "late-closing", "late-disposed", "late-not-ready" })
                using (var f = new ModernEditorToolFixture())
                {
                    if (state == "busy" || state == "closing") f.Base.Set(state, true);
                    if (state == "disposed") f.Window.Dispose(); if (state == "not-ready") f.Base.Ready(false);
                    f.Override = (method, values) =>
                    {
                        if (method != "snapshots") return null;
                        if (state == "error") throw new InvalidOperationException("owned capture refusal");
                        if (state == "late-closing") f.Base.Set("closing", true);
                        if (state == "late-disposed") f.Window.Dispose();
                        if (state == "late-not-ready") f.Base.Ready(false);
                        return null;
                    };
                    Assert.ThrowsException<InvalidOperationException>(() => Wait(f.Window.CaptureForTool()), state);
                    if (state != "busy") Assert.IsFalse(f.Base.Get<bool>("busy"));
                }
            foreach (string state in new[] { "closing", "disposed", "not-ready", "missing", "stale", "valid" })
                using (var f = new ModernEditorToolFixture())
                {
                    if (state == "closing") f.Base.Set("closing", true); if (state == "disposed") f.Window.Dispose(); if (state == "not-ready") f.Base.Ready(false);
                    var doc = state == "missing" ? new EditorDocument(f.Module) : f.Document;
                    if (state == "valid") f.Private("RequireVersion", doc, 1);
                    else Assert.ThrowsException<InvalidOperationException>(() => f.Private("RequireVersion", doc, state == "stale" ? 0 : 1));
                }
        }
        [STATestMethod]
        public void ExactModuleLookupRecognizesOwnedNativeIdentityAndReportsMissingModules()
        {
            using (var f = new ModernEditorToolFixture())
            using (var missing = new EditorFixture())
            {
                Assert.AreSame(f.Document, f.Window.FindToolDocument(f.Module));
                Assert.AreSame(f.Base.Document, f.Window.FindToolDocument(new EditorVbeModule(f.Base.Native.Vbe, f.Base.Native.Project, f.Base.Native.Original)));
                var foreign = new EditorVbeContract(); Assert.ThrowsException<InvalidOperationException>(() => f.Window.FindToolDocument(foreign.Adapter));
                Assert.ThrowsException<InvalidOperationException>(() => f.Window.FindToolDocument(missing));
                Assert.IsFalse(f.Window.HasPendingEditorDraft); f.Document.Edit(f.Document.Text + "\n' draft"); Assert.IsTrue(f.Window.HasPendingEditorDraft);
                f.Module.Code += "\n' external"; f.Document.Observe(); f.Document.Edit(f.Document.Baseline); Assert.IsTrue(f.Window.HasPendingEditorDraft);
            }
        }
        [STATestMethod]
        public void ToolReadKeepsStaleRendererTextAndReportsActualNativeConflict()
        {
            foreach (string state in new[] { "null", "stale", "clean", "dirty", "conflict" })
                using (var f = new ModernEditorToolFixture())
                {
                    string draft = f.Document.Text + "\n' draft";
                    if (state == "conflict") f.Module.Code += "\n' native";
                    f.Override = (method, values) => method == "read" ? state == "null" ? "null" : f.Json.Serialize(new { text = state == "clean" ? f.Document.Text : draft, version = state == "stale" ? 0 : 1, selection = new { startLineNumber = 1 } }) : null;
                    if (state == "null") Assert.ThrowsException<InvalidOperationException>(() => Wait(f.Window.ReadForTool(f.Document)));
                    else
                    {
                        var result = f.Result(Wait(f.Window.ReadForTool(f.Document))); Assert.AreEqual(state != "clean", result["Dirty"]); Assert.AreEqual(state == "conflict", result["Conflict"]);
                        Assert.AreEqual(state == "stale" || state == "clean" ? f.Document.Baseline : draft, f.Document.Text); Assert.AreEqual(EditorDocument.Hash(f.Module.Code), result["NativeSha256"]);
                    }
                    Assert.IsFalse(f.Base.Get<bool>("busy"));
                }
        }
        [STATestMethod]
        public void NavigationValidatesEachSelectionBoundBeforeApplyingOwnedRange()
        {
            int[][] invalid = { new[] { 0, 1, 1, 1 }, new[] { 2, 1, 1, 1 }, new[] { 1, 1, 99, 1 }, new[] { 1, 0, 1, 1 }, new[] { 1, 1, 1, 0 }, new[] { 1, 999, 1, 999 }, new[] { 1, 1, 1, 999 }, new[] { 1, 2, 1, 1 } };
            using (var f = new ModernEditorToolFixture())
            {
                foreach (var range in invalid) Assert.ThrowsException<ArgumentException>(() => Wait(f.Window.NavigateForTool(f.Document, 1, range[0], range[1], range[2], range[3])));
                f.Override = (method, values) => method == "selectRange" ? "false" : null;
                Assert.ThrowsException<InvalidOperationException>(() => Wait(f.Window.NavigateForTool(f.Document, 1, 1, 1, 1, 2)));
                f.Override = (method, values) => method == "selectRange" ? "true" : null;
                var result = f.Result(Wait(f.Window.NavigateForTool(f.Document, 1, 1, 1, 2, 2))); Assert.AreEqual(f.Document.Id, result["DocumentId"]);
            }
        }
        [STATestMethod]
        public void EditRejectsReadOnlyDivergenceAndInvalidRendererReplacements()
        {
            foreach (string state in new[] { "readonly", "conflict", "native-changed", "not-number", "zero" })
                using (var f = new ModernEditorToolFixture())
                {
                    if (state == "readonly") f.Module.CanWrite = false;
                    if (state == "conflict") f.Document.Edit(f.Document.Text + "\n' existing draft");
                    if (state == "conflict" || state == "native-changed") f.Module.Code += "\n' native";
                    if (state == "not-number" || state == "zero") f.Override = (method, values) => method == "apply" ? state == "zero" ? "0" : "invalid" : null;
                    Assert.ThrowsException<InvalidOperationException>(() => Wait(f.Window.EditForTool(f.Document, 1, f.Document.Text + "\n' proposed")), state);
                    Assert.AreEqual(0, f.Module.Writes); Assert.IsFalse(f.Base.Get<bool>("busy"));
                }
        }
        [STATestMethod]
        public void EditPreservesNewerTypingAndRejectsChangedPreparedSnapshots()
        {
            foreach (string state in new[] { "newer", "already-newer", "text-changed", "baseline-changed" })
                using (var f = new ModernEditorToolFixture())
                {
                    string requested = f.Document.Text + "\n' requested";
                    f.Override = (method, values) =>
                    {
                        if (method == "apply" && f.Applies == 1 && state == "already-newer") { f.Versions[f.Document.Id] = 3; f.Document.Edit(f.Document.Text + "\n' real newer"); return "2"; }
                        if (method == "snapshots" && f.Captures == 3)
                        {
                            if (state == "newer") return f.Json.Serialize(new[] { f.Snapshot(requested + "\n' real newer", 3) });
                            if (state == "text-changed") f.Document.Edit(requested + "\n' concurrent");
                            if (state == "baseline-changed") f.Document.AcceptRemote("changed baseline");
                        }
                        return null;
                    };
                    if (state == "text-changed" || state == "baseline-changed") Assert.ThrowsException<InvalidOperationException>(() => Wait(f.Window.EditForTool(f.Document, 1, requested)));
                    else { var result = f.Result(Wait(f.Window.EditForTool(f.Document, 1, requested))); Assert.AreEqual(false, result["Synchronized"]); Assert.AreEqual(3, result["Version"]); Assert.IsTrue(f.Document.Text.Contains("real newer")); }
                    Assert.AreEqual(0, f.Module.Writes);
                }
        }
        [STATestMethod]
        public void EditReportsNativeFailureAndReconcilesOwnedSuccessWithoutDroppingTyping()
        {
            foreach (string state in new[] { "failure", "plain", "callback", "newer", "invalid", "zero" })
                using (var f = new ModernEditorToolFixture())
                {
                    f.Module.Fail = state == "failure"; int callbacks = 0;
                    Action callback = state == "plain" || state == "failure" ? null : (Action)(() => { callbacks++; if (state == "newer") f.Document.Edit(f.Document.Text + "\n' post-write typing"); });
                    if (state == "invalid" || state == "zero") f.Override = (method, values) => method == "apply" && f.Applies == 2 ? state == "zero" ? "0" : "invalid" : null;
                    var result = f.Result(Wait(f.Window.EditForTool(f.Document, 1, f.Document.Text + "\n' edit", callback)));
                    Assert.AreEqual(state != "failure", result["Synchronized"]); Assert.AreEqual(state == "failure" ? 0 : 1, f.Module.Writes);
                    Assert.AreEqual(callback == null ? 0 : 1, callbacks); Assert.AreEqual(state == "failure" || state == "newer", f.Document.Dirty);
                    if (state == "newer") Assert.IsTrue(f.Document.Text.Contains("post-write typing"));
                }
        }
        [STATestMethod]
        public void SynchronizeChecksOwnedNativeShaAndPreparedSnapshotsAtBothBoundaries()
        {
            foreach (string state in new[] { "wrong-sha", "text-changed", "baseline-changed", "native-changed" })
                using (var f = new ModernEditorToolFixture())
                {
                    f.Document.Edit(f.Document.Text + "\n' draft"); string hash = EditorDocument.Hash(f.Module.Code);
                    f.Override = (method, values) =>
                    {
                        if (method == "snapshots" && f.Captures == 2)
                        {
                            if (state == "text-changed") f.Document.Edit(f.Document.Text + "\n' concurrent");
                            if (state == "baseline-changed") f.Document.AcceptRemote("new baseline");
                            if (state == "native-changed") f.Module.Code += "\n' native concurrent";
                        }
                        return null;
                    };
                    Assert.ThrowsException<InvalidOperationException>(() => Wait(f.Window.SynchronizeForTool(f.Document, 1, state == "wrong-sha" ? new string('a', 64) : hash)), state);
                    Assert.AreEqual(0, f.Module.Writes); Assert.IsFalse(f.Base.Get<bool>("busy"));
                }
        }
        [STATestMethod]
        public void SynchronizeReconcilesOwnedWritesCallbacksAndInvalidOrNewerRendererResults()
        {
            foreach (string state in new[] { "plain", "callback", "newer", "invalid", "zero" })
                using (var f = new ModernEditorToolFixture())
                {
                    f.Document.Edit(f.Document.Text + "\n' draft"); int callbacks = 0;
                    Action callback = state == "plain" ? null : (Action)(() => { callbacks++; if (state == "newer") f.Document.Edit(f.Document.Text + "\n' newer typing"); });
                    if (state == "invalid" || state == "zero") f.Override = (method, values) => method == "apply" ? state == "zero" ? "0" : "invalid" : null;
                    var result = f.Result(Wait(f.Window.SynchronizeForTool(f.Document, 1, EditorDocument.Hash(f.Module.Code).ToUpperInvariant(), callback)));
                    Assert.AreEqual(true, result["Synchronized"]); Assert.AreEqual(1, f.Module.Writes); Assert.AreEqual(callback == null ? 0 : 1, callbacks);
                    Assert.AreEqual(state == "newer", f.Document.Dirty); Assert.AreEqual(EditorDocument.Hash(f.Module.Code), result["NativeSha256"]);
                }
        }
    }
}