using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Unit.Editor
{
    [TestClass]
    public sealed class ModernEditorSaveTests
    {
        /// <summary>Synchronized changes stay recoverable until native file and host persistence are both verified.</summary>
        [STATestMethod]
        public void SynchronizedRecoveryDraftSurvivesFailedAndUnknownSaveAndRetiresAfterVerifiedSave()
        {
            foreach (string state in new[] { "failed", "host-unsaved", "host-unknown", "late-edit", "saved" })
                using (var fixture = new ModernEditorDebugFixture())
                {
                    fixture.Ready(true);
                    fixture.Window.NativeSaveObservationTimeout = TimeSpan.FromMilliseconds(80);
                    Directory.CreateDirectory(fixture.Storage.Root);
                    string path = Path.Combine(fixture.Storage.Root, "owned.bas");
                    File.WriteAllText(path, "owned baseline");
                    fixture.Native.Project.Path = path;
                    fixture.Document.Observe();
                    string changed = fixture.Document.Text + "\n' unsaved synchronized edit";
                    fixture.Document.Edit(changed);
                    ModernEditorDebugFixture.Wait(fixture.Window.ProcessDocuments(true));
                    Assert.IsFalse(fixture.Document.Dirty);
                    Assert.AreEqual(changed, fixture.Window.Drafts.Recover(fixture.Document.RecoveryKey).Text);
                    int saveCalls = 0;
                    fixture.Window.NativeSave = native =>
                    {
                        saveCalls++;
                        if (state == "failed") throw new IOException("Synthetic save failure");
                        if (state == "late-edit") fixture.Document.Edit(changed + "\n' late typing");
                    };
                    fixture.Window.NativeHostSaved = native => state == "host-unknown" ? (bool?)null : state != "host-unsaved";
                    if (state == "failed")
                        Assert.ThrowsException<IOException>(() => ModernEditorDebugFixture.Wait(fixture.Window.SaveDocument(fixture.Document.Id)));
                    else if (state == "host-unsaved" || state == "late-edit")
                        Assert.ThrowsException<InvalidOperationException>(() => ModernEditorDebugFixture.Wait(fixture.Window.SaveDocument(fixture.Document.Id)));
                    else ModernEditorDebugFixture.Wait(fixture.Window.SaveDocument(fixture.Document.Id));
                    Assert.AreEqual(1, saveCalls);
                    var retained = fixture.Window.Drafts.Recover(fixture.Document.RecoveryKey);
                    if (state == "saved") Assert.IsNull(retained);
                    else Assert.AreEqual(state == "late-edit" ? changed + "\n' late typing" : changed, retained.Text);
                }
        }
        /// <summary>One native command may complete later; only its unchanged revision retires the draft.</summary>
        [STATestMethod]
        public void SaveObservesDeferredPersistenceAndRetainsDraftOnScopeOrLifetimeChanges()
        {
            foreach (string state in new[] { "late-project", "late-file", "late-host", "save-as", "cancelled", "missing", "unknown", "closing", "disposed", "version", "source", "component", "project", "path", "unknown-path", "edit" })
                using (var f = new ModernEditorDebugFixture())
                {
                    Directory.CreateDirectory(f.Storage.Root); f.Ready(true);
                    f.Window.NativeSaveObservationTimeout = TimeSpan.FromMilliseconds(150);
                    string path = Path.Combine(f.Storage.Root, "owned.bas");
                    f.Native.Project.Path = state == "save-as" ? null : path;
                    if (state != "late-file" && state != "missing") File.WriteAllText(path, "owned");
                    f.Document.Observe();
                    f.Document.Edit(f.Document.Text + "\n' synchronized recovery");
                    ModernEditorDebugFixture.Wait(f.Window.ProcessDocuments(true));
                    string key = f.Document.RecoveryKey, code = f.Document.Text;
                    if (state == "unknown-path") f.Native.Project.FailFileName = true;
                    int saves = 0, reads = 0, capturesAfterSave = 0;
                    f.Rendering = (method, values) =>
                    {
                        if (saves > 0 && method == "snapshots") capturesAfterSave++;
                        return "null";
                    };
                    f.Window.NativeSave = native =>
                    {
                        saves++;
                        if (state == "late-project" || state == "cancelled") f.Native.Project.Saved = false;
                    };
                    f.Window.NativeHostSaved = native =>
                    {
                        reads++;
                        if (reads == 1)
                            f.Window.BeginInvoke(new Action(() =>
                            {
                                if (state == "late-project") f.Native.Project.Saved = true;
                                if (state == "late-file") File.WriteAllText(path, "owned");
                                if (state == "save-as") f.Native.Project.Path = path;
                                if (state == "closing") f.Set("closing", true);
                                if (state == "disposed") f.Window.Dispose();
                                if (state == "version") f.Get<Dictionary<string, int>>("versions")[f.Document.Id]++;
                                if (state == "source") f.Native.Original.CodeModule.Raw += "\n' external revision";
                                if (state == "component") f.Native.Project.VBComponents.Items.Clear();
                                if (state == "project") f.Native.Vbe.VBProjects.Clear();
                                if (state == "unknown-path") f.Native.Project.FailFileName = false;
                                if (state == "edit") f.Document.Edit(code + "\n' late revision");
                                if (state == "path") f.Native.Project.Path = Path.Combine(f.Storage.Root, "other.bas");
                            }));
                        if (state == "unknown") return null;
                        if (state == "closing" || state == "disposed" || state == "version" || state == "source" || state == "component" || state == "project" || state == "path" || state == "unknown-path" || state == "edit") return false;
                        return state != "late-host" || reads > 1;
                    };
                    bool success = state == "late-project" || state == "late-file" || state == "late-host" || state == "save-as";
                    if (success || state == "unknown") ModernEditorDebugFixture.Wait(f.Window.SaveDocument(f.Document.Id));
                    else Assert.ThrowsException<InvalidOperationException>(() => ModernEditorDebugFixture.Wait(f.Window.SaveDocument(f.Document.Id)), state);
                    Assert.AreEqual(1, saves, state);
                    Assert.IsFalse(f.Get<bool>("busy"), state);
                    var draft = f.Window.Drafts.Recover(key);
                    if (success) { Assert.IsNull(draft, state); Assert.IsTrue(reads >= 2, state); }
                    else Assert.AreEqual(state == "edit" ? code + "\n' late revision" : code, draft.Text, state);
                    if (state == "closing" || state == "disposed") Assert.AreEqual(0, capturesAfterSave, state);
                    if (state == "unknown") StringAssert.Contains(f.Get<string>("lastSaveError"), "could not be verified");
                }
        }
        /// <summary>A runner or host context cannot move deferred native readback away from the editor STA.</summary>
        [STATestMethod]
        public void SavePinsDeferredReadbackToOwnerUnderNonWindowsFormsContext()
        {
            using (var f = new ModernEditorDebugFixture())
            {
                Directory.CreateDirectory(f.Storage.Root); f.Ready(true);
                string path = Path.Combine(f.Storage.Root, "owned.bas");
                File.WriteAllText(path, "owned"); f.Native.Project.Path = path;
                f.Document.Observe(); f.Document.Edit(f.Document.Text + "\n' recovery");
                ModernEditorDebugFixture.Wait(f.Window.ProcessDocuments(true));
                int owner = Thread.CurrentThread.ManagedThreadId, saves = 0, reads = 0;
                f.Window.NativeSave = native => { Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId); saves++; f.Native.Project.Saved = false; };
                f.Window.NativeHostSaved = native =>
                {
                    Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                    if (++reads == 2) f.Native.Project.Saved = true;
                    return true;
                };
                var previous = SynchronizationContext.Current;
                var foreignContext = new SynchronizationContext();
                try
                {
                    SynchronizationContext.SetSynchronizationContext(foreignContext);
                    ModernEditorDebugFixture.Wait(f.Window.SaveDocument(f.Document.Id));
                    Assert.AreSame(foreignContext, SynchronizationContext.Current);
                }
                finally { SynchronizationContext.SetSynchronizationContext(previous); }
                Assert.AreEqual(1, saves); Assert.IsTrue(reads >= 2);
                Assert.IsNull(f.Window.Drafts.Recover(f.Document.RecoveryKey));
            }
        }
        [STATestMethod]
        public void SaveGuardsWaitForOwnedOperationsAndRejectUnavailableDocuments()
        {
            foreach (string state in new[] { "not-ready", "closing", "disposed", "null", "missing", "managed", "late-closing", "late-disposed", "release" })
                using (var f = new ModernEditorDebugFixture())
                {
                    Directory.CreateDirectory(f.Storage.Root); f.Ready(state != "not-ready"); int saves = 0;
                    f.Window.NativeSave = native => saves++;
                    f.Native.Project.Path = Path.Combine(f.Storage.Root, "owned.bas"); File.WriteAllText(f.Native.Project.Path, "owned"); f.Window.NativeHostSaved = native => true;
                    string id = state == "null" ? null : state == "missing" ? "missing" : f.Document.Id;
                    if (state == "managed") id = ModernEditorDebugFixture.Wait(f.Window.OpenModule(f.Storage)).Id;
                    if (state == "closing") f.Set("closing", true);
                    if (state == "disposed") f.Window.Dispose();
                    if (state.StartsWith("late-", StringComparison.Ordinal) || state == "release")
                    {
                        f.Set("busy", true); f.Window.BeginInvoke(new Action(() => { if (state == "late-closing") f.Set("closing", true); if (state == "late-disposed") f.Window.Dispose(); f.Set("busy", false); }));
                    }
                    if (state == "managed") Assert.ThrowsException<InvalidOperationException>(() => ModernEditorDebugFixture.Wait(f.Window.SaveDocument(id)));
                    else ModernEditorDebugFixture.Wait(f.Window.SaveDocument(id));
                    Assert.AreEqual(state == "release" ? 1 : 0, saves, state);
                }
        }

        [STATestMethod]
        public void SaveVerifiesNativeFileProjectAndHostAndPreservesEveryFailure()
        {
            foreach (string state in new[] { "native-failure", "filename-failure", "null-path", "blank-path", "missing-file", "project-unsaved", "host-unsaved", "host-unknown", "saved" })
                using (var f = new ModernEditorDebugFixture())
                {
                    Directory.CreateDirectory(f.Storage.Root); f.Ready(true); f.Window.NativeSaveObservationTimeout = TimeSpan.FromMilliseconds(80); int saves = 0;
                    string path = Path.Combine(f.Storage.Root, "owned.bas"); File.WriteAllText(path, "owned"); f.Native.Project.Path = path;
                    var errors = f.Get<Dictionary<string, string>>("documentSynchronizationErrors");
                    errors[f.Document.Id] = "old synchronization failure"; f.Set("lastSaveError", "old save failure");
                    f.Window.NativeSave = native => { Assert.AreSame(f.Native.Adapter, native); saves++; if (state == "native-failure") throw new IOException("owned native save failure"); };
                    f.Window.NativeHostSaved = native => state == "host-unsaved" ? false : state == "host-unknown" ? (bool?)null : true;
                    if (state == "filename-failure") f.Native.Project.FailFileName = true;
                    if (state == "null-path") f.Native.Project.Path = null;
                    if (state == "blank-path") f.Native.Project.Path = " ";
                    if (state == "missing-file") File.Delete(path);
                    if (state == "project-unsaved") f.Native.Project.Saved = false;
                    if (state == "saved" || state == "host-unknown")
                    {
                        ModernEditorDebugFixture.Wait(f.Window.SaveDocument(f.Document.Id));
                        Assert.IsFalse(errors.ContainsKey(f.Document.Id));
                        Assert.AreEqual(UiText.Get(state == "saved" ? "Saved." : "The native Save command finished, but the host document's saved state could not be verified."), f.Get<Label>("status").Text);
                        Assert.AreEqual(state == "saved", f.Get<string>("lastSaveError") == null);
                    }
                    else
                    {
                        if (state == "native-failure") Assert.ThrowsException<IOException>(() => ModernEditorDebugFixture.Wait(f.Window.SaveDocument(f.Document.Id)));
                        else Assert.ThrowsException<InvalidOperationException>(() => ModernEditorDebugFixture.Wait(f.Window.SaveDocument(f.Document.Id)));
                        Assert.IsFalse(string.IsNullOrEmpty(f.Get<string>("lastSaveError"))); Assert.AreEqual(f.Document.Text, EditorDocument.Normalize(f.Native.Adapter.Read()));
                    }
                    Assert.AreEqual(1, saves); Assert.IsFalse(f.Get<bool>("busy"));
                }
        }

        [STATestMethod]
        public void SaveRejectsPendingBuffersInItsProjectAndHonorsLateLifetimeChanges()
        {
            foreach (string state in new[] { "dirty", "conflict", "foreign-project", "managed", "closing", "disposed" })
                using (var f = new ModernEditorDebugFixture())
                {
                    Directory.CreateDirectory(f.Storage.Root); f.Ready(true); f.Window.NativeSaveObservationTimeout = TimeSpan.FromMilliseconds(80); int saves = 0, captures = 0;
                    f.Native.Project.Path = Path.Combine(f.Storage.Root, "owned.bas"); File.WriteAllText(f.Native.Project.Path, "owned");
                    f.Window.NativeSave = native => saves++; f.Window.NativeHostSaved = native => true;
                    if (state == "foreign-project") { var other = new EditorVbeContract(); ModernEditorDebugFixture.Wait(f.Window.OpenModule(other.Adapter)); }
                    if (state == "managed") ModernEditorDebugFixture.Wait(f.Window.OpenModule(f.Storage));
                    f.Rendering = (method, values) =>
                    {
                        if (method == "snapshots" && ++captures == 2)
                        {
                            if (state == "dirty") f.Document.Edit(f.Document.Text + "\n' late draft");
                            if (state == "conflict") { f.Document.Edit(f.Document.Text + "\n' late draft"); f.Native.Original.CodeModule.Raw += "\n' external"; f.Document.Observe(); f.Document.Edit(f.Document.Baseline); }
                            if (state == "closing") f.Set("closing", true);
                            if (state == "disposed") f.Window.Dispose();
                        }
                        return "null";
                    };
                    if (state == "dirty" || state == "conflict") Assert.ThrowsException<InvalidOperationException>(() => ModernEditorDebugFixture.Wait(f.Window.SaveDocument(f.Document.Id)));
                    else ModernEditorDebugFixture.Wait(f.Window.SaveDocument(f.Document.Id));
                    Assert.AreEqual(state == "foreign-project" || state == "managed" ? 1 : 0, saves, state); Assert.IsFalse(f.Get<bool>("busy"));
                }
        }

        [STATestMethod]
        public void NativeSaveFocusesOwnedTargetAndRejectsChangedMissingOrDisabledCommands()
        {
            foreach (string state in new[] { "changed", "missing", "disabled", "executed" })
                using (var f = new ModernEditorDebugFixture())
                {
                    var command = state == "missing" ? null : f.Command(3, "Save", enabled: state != "disabled");
                    if (state == "changed") f.Native.Original.CodeModule.CodePane.Window.OnFocus = () => f.Native.Vbe.ActiveCodePane = new EditorVbeContract.Component { Name = "Other" }.CodeModule.CodePane;
                    if (state == "executed") f.Window.NativeSave(f.Native.Adapter);
                    else Assert.ThrowsException<InvalidOperationException>(() => f.Window.NativeSave(f.Native.Adapter));
                    Assert.AreEqual(1, f.Native.Original.CodeModule.CodePane.Window.Focuses);
                    Assert.AreEqual(state == "executed" ? 1 : 0, command?.ExecuteCount ?? 0);
                }
        }

        [STATestMethod]
        public void DocumentHostStatusHandlesManagedNativeUnavailableAndBoundaryFailure()
        {
            using (var f = new ModernEditorDebugFixture())
            {
                var method = typeof(ModernEditorWindow).GetMethod("ReadDocumentHostSaved", BindingFlags.NonPublic | BindingFlags.Instance);
                var managed = ModernEditorDebugFixture.Wait(f.Window.OpenModule(f.Storage)); Assert.IsNull(method.Invoke(f.Window, new object[] { managed }));
                f.Window.NativeHostSaved = native => throw new IOException("owned host unavailable"); Assert.IsNull(method.Invoke(f.Window, new object[] { f.Document }));
                f.Window.NativeHostSaved = native => true; Assert.AreEqual(true, method.Invoke(f.Window, new object[] { f.Document }));
            }
            using (var f = new ModernEditorDebugFixture())
            {
                f.Native.Project.Type = 100; Assert.IsNull(f.Window.NativeHostSaved(f.Native.Adapter));
                Directory.CreateDirectory(f.Storage.Root); f.Native.Project.Type = 101; f.Native.Project.Path = Path.Combine(f.Storage.Root, "owned.swp"); File.WriteAllText(f.Native.Project.Path, "owned");
                Assert.AreEqual(true, f.Window.NativeHostSaved(f.Native.Adapter));
            }
        }
    }
}
