using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit.Editor
{
    [TestClass]
    public sealed class ModernEditorSaveTests
    {
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
                Directory.CreateDirectory(f.Storage.Root); f.Ready(true); int saves = 0;
                string path = Path.Combine(f.Storage.Root, "owned.bas"); File.WriteAllText(path, "owned"); f.Native.Project.Path = path;
                f.Set("synchronizationError", "old synchronization failure"); f.Set("lastSaveError", "old save failure");
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
                    Assert.IsNull(f.Get<string>("synchronizationError"));
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
                Directory.CreateDirectory(f.Storage.Root); f.Ready(true); int saves = 0, captures = 0;
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