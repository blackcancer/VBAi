using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace CodexVBE
{
    internal sealed partial class ModernEditorWindow
    {
        // Native boundary permits testing cancellation without replacing synchronization.
        internal Action<EditorVbeModule> NativeSave = SaveInVbe;
        internal Func<EditorVbeModule, bool?> NativeHostSaved = ReadHostSaved;
        private string lastSaveError;

        private bool? ReadDocumentHostSaved(EditorDocument document)
        {
            try { return document.Module is EditorVbeModule native ? NativeHostSaved(native) : null; }
            catch { return null; }
        }

        internal async Task SaveDocument(string id)
        {
            while (busy && !closing && !IsDisposed) await Task.Delay(15);
            if (closing || IsDisposed || !Ready || !documents.TryGetValue(id ?? "", out var document)) return;
            if (!(document.Module is EditorVbeModule native)) throw new InvalidOperationException("This document has no native host to save.");
            busy = true;
            try
            {
                await ProcessDocumentsCore(true);
                // Unwind the WebView callback before the native Save/Save As modal dialog.
                await Task.Yield();
                await CaptureDocuments();
                if (closing || IsDisposed) return;
                var projectDocuments = documents.Values.Where(item => item.Module is EditorVbeModule module &&
                    ReferenceEquals(module.Project, native.Project)).ToArray();
                if (projectDocuments.Any(item => item.Dirty || item.Conflict))
                    throw new InvalidOperationException("The document was not saved because some edits are still pending synchronization.");
                NativeSave(native);
                string path = null;
                try { path = (string)((dynamic)native.Project).FileName; } catch { }
                if (!(bool)((dynamic)native.Project).Saved || string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                    throw new InvalidOperationException("Saving was cancelled or not completed. Your code remains in the editor and host document.");
                bool? hostSaved = NativeHostSaved(native);
                if (hostSaved == false)
                    throw new InvalidOperationException("The host document was not saved. Your code remains in the editor and host document.");
                synchronizationError = null;
                lastSaveError = hostSaved.HasValue ? null : "The native Save command finished, but the host document's saved state could not be verified.";
                status.Text = UiText.Get(lastSaveError ?? "Saved.");
            }
            catch (Exception error)
            {
                lastSaveError = error.Message;
                PreserveDrafts();
                throw;
            }
            finally { busy = false; }
        }

        private static bool? ReadHostSaved(EditorVbeModule native)
        {
            dynamic state = new VbeProjectComponents(native.Vbe, null).PersistenceStatus(native.ProjectName);
            return (bool)state.HostAvailable ? (bool?)state.HostSaved : null;
        }

        private static void SaveInVbe(EditorVbeModule native)
        {
            dynamic vbe = native.Vbe;
            // Save must target the tab's project even if a different workbook is active.
            native.ShowNative(1, 1);
            vbe.ActiveCodePane.Window.SetFocus();
            if (!native.IsComponent((object)vbe.ActiveCodePane.CodeModule.Parent))
                throw new InvalidOperationException("The native save target changed. Select the module and retry.");
            dynamic save = vbe.CommandBars.FindControl(1, 3);
            if (save == null || !(bool)save.Enabled)
                throw new InvalidOperationException("The native Save command is unavailable.");
            save.Execute();
        }
    }
}
