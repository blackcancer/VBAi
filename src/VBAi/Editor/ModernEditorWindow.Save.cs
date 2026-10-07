using System;
using System.IO;
using System.Diagnostics;
using System.Threading;
using System.Linq;
using System.Threading.Tasks;

namespace VBAi
{

    /// <summary>Synchronizes editor buffers and verifies native host persistence after saving VBE-backed modules.</summary>
    internal sealed partial class ModernEditorWindow
    {
        // Native boundary permits testing cancellation without replacing synchronization.
        /// <summary>Replaceable native Save command route for the selected project's VBE module.</summary>
        internal Action<EditorVbeModule> NativeSave = SaveInVbe;

        /// <summary>Replaceable readback of the host document's persistence state after Save.</summary>
        internal Func<EditorVbeModule, bool?> NativeHostSaved = ReadHostSaved;

        /// <summary>Last save or persistence-verification error shown in the editor status area.</summary>
        private string lastSaveError;

        /// <summary>Bounds read-only persistence observation after the one native Save command.</summary>
        internal TimeSpan NativeSaveObservationTimeout = TimeSpan.FromSeconds(5);

        /// <summary>Reads host persistence status for native documents while treating unavailable status as unknown.</summary>
        /// <param name="document">Editor document whose adapter may expose native host persistence state.</param>
        /// <returns><see langword="true"/> or <see langword="false"/> when available; otherwise <see langword="null"/>.</returns>
        private bool? ReadDocumentHostSaved(EditorDocument document)
        {
            try { return document.Module is EditorVbeModule native ? NativeHostSaved(native) : null; }
            catch { return null; }
        }

        /// <summary>Saves the active editor document through its registered native save route.</summary>
        /// <param name="id">Editor document identifier selected for native host saving.</param>
        /// <returns>A task completed after synchronization, native Save, and available host persistence readback.</returns>
        internal Task SaveDocument(string id)
        {
            if (closing || IsDisposed || !Ready) return Task.CompletedTask;
            if (InvokeRequired) throw new InvalidOperationException("Native Save must start on the editor's owning VBE thread.");
            return VbeUiTask.Run(async () => { await SaveDocumentCore(id); return true; });
        }

        /// <summary>Runs synchronization and persistence observation with all continuations pinned to the owning STA.</summary>
        /// <param name="id">Document identifier captured by the original Save request.</param>
        /// <returns>A task completed after the one native command and its guarded readback.</returns>
        private async Task SaveDocumentCore(string id)
        {
            while (busy && !closing && !IsDisposed) await Task.Delay(15);
            if (closing || IsDisposed || !Ready || !documents.TryGetValue(id ?? "", out var document)) return;
            if (!(document.Module is EditorVbeModule native)) throw new InvalidOperationException("This document has no native host to save.");
            busy = true;
            try
            {
                await ProcessDocumentsCore(true);
                // Yield before invoking the native command; persistence still requires readback.
                await Task.Yield();
                await CaptureDocuments();
                if (closing || IsDisposed) return;
                var projectDocuments = documents.Values.Where(item => item.Module is EditorVbeModule module &&
                    EditorVbeModule.SameIdentity(module.Project, native.Project)).ToArray();
                if (projectDocuments.Any(item => item.Dirty || item.Conflict))
                    throw new InvalidOperationException("The document was not saved because some edits are still pending synchronization.");
                var savedRevisions = projectDocuments.ToDictionary(item => item.Id, item => new
                {
                    Document = item, Native = (EditorVbeModule)item.Module,
                    Component = ((EditorVbeModule)item.Module).Component,
                    ModuleName = ((EditorVbeModule)item.Module).ModuleName,
                    item.Text, item.Baseline, item.RecoveryKey, Version = versions[item.Id]
                });
                object originalProject = native.Project;
                int ownerThread = Thread.CurrentThread.ManagedThreadId, ownerProcess = native.HostProcessId;
                bool pathAvailable = false;
                string ReadProjectPath()
                {
                    string value = null;
                    pathAvailable = false;
                    try { value = (string)((dynamic)originalProject).FileName; pathAvailable = true; } catch { }
                    return !string.IsNullOrWhiteSpace(value) && Path.IsPathRooted(value) ? Path.GetFullPath(value) : null;
                }
                // A previously unsaved project may acquire one Save As path; an existing path is fixed.
                string savedPath = ReadProjectPath();
                bool allowSaveAsPath = pathAvailable && savedPath == null;
                void RequireSaveScope()
                {
                    if (Thread.CurrentThread.ManagedThreadId != ownerThread)
                        throw new InvalidOperationException("Native Save observation must remain on its original VBE thread.");
                    if (closing || IsDisposed || !Ready)
                        throw new InvalidOperationException("The editor closed before saving could be verified. Your draft is preserved.");
                    if (!EditorVbeModule.SameIdentity(native.Project, originalProject) || native.HostProcessId != ownerProcess)
                        throw new InvalidOperationException("The native save target changed. Select the module and retry.");
                    var currentProjectDocuments = documents.Values.Where(item => item.Module is EditorVbeModule module &&
                        EditorVbeModule.SameIdentity(module.Project, originalProject)).ToArray();
                    if (currentProjectDocuments.Length != savedRevisions.Count || currentProjectDocuments.Any(item =>
                        !savedRevisions.TryGetValue(item.Id, out var revision) || !ReferenceEquals(item, revision.Document)))
                        throw new InvalidOperationException("The project changed before saving could be verified. Your draft is preserved.");
                    foreach (var revision in savedRevisions.Values)
                    {
                        if (!revision.Native.IsComponent(revision.Component) || revision.Native.ModuleName != revision.ModuleName ||
                            !EditorVbeModule.SameIdentity(revision.Native.Project, originalProject) ||
                            !revision.Native.CanWrite || EditorDocument.Normalize(revision.Native.Read()) != revision.Text ||
                            revision.Document.Dirty || revision.Document.Conflict || revision.Document.Text != revision.Text ||
                            revision.Document.Baseline != revision.Baseline || revision.Document.RecoveryKey != revision.RecoveryKey ||
                            !versions.TryGetValue(revision.Document.Id, out int version) || version != revision.Version)
                            throw new InvalidOperationException("The document changed before saving could be verified. Your draft is preserved.");
                    }
                    string currentPath = ReadProjectPath();
                    if (savedPath == null && allowSaveAsPath) savedPath = currentPath;
                    else if (!string.Equals(savedPath, currentPath, StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException("The project path changed before saving could be verified. Your draft is preserved.");
                }
                RequireSaveScope();
                NativeSave(native); // Exactly one delivery; only read-only observation follows.
                var observation = Stopwatch.StartNew();
                bool persisted = false;
                bool? hostSaved = null;
                while (true)
                {
                    RequireSaveScope();
                    hostSaved = NativeHostSaved(native);
                    // COM/adapter reads may pump callbacks; revalidate before accepting their result.
                    RequireSaveScope();
                    persisted = (bool)((dynamic)originalProject).Saved && pathAvailable && savedPath != null && File.Exists(savedPath);
                    if (persisted && hostSaved == true) break;
                    if (observation.Elapsed >= NativeSaveObservationTimeout) break;
                    await Task.Delay(Math.Max(1, Math.Min(50, (int)(NativeSaveObservationTimeout - observation.Elapsed).TotalMilliseconds)));
                    RequireSaveScope();
                    await CaptureDocuments();
                }
                if (!persisted)
                    throw new InvalidOperationException("Saving was cancelled or not completed. Your code remains in the editor and host document.");
                if (hostSaved == false)
                    throw new InvalidOperationException("The host document was not saved. Your code remains in the editor and host document.");
                RequireSaveScope();
                foreach (var savedDocument in projectDocuments)
                {
                    documentSynchronizationErrors.Remove(savedDocument.Id);
                    // An unknown host state cannot retire the recovery copy.
                    var savedRevision = savedRevisions[savedDocument.Id];
                    if (hostSaved == true && !savedDocument.Dirty && !savedDocument.Conflict
                        && savedDocument.Text == savedRevision.Text && savedDocument.Baseline == savedRevision.Baseline
                        && versions.TryGetValue(savedDocument.Id, out int currentVersion) && currentVersion == savedRevision.Version)
                        Drafts.ClearOwn(savedDocument);
                }
                lastSaveError = hostSaved.HasValue ? null : "The native Save command finished, but the host document's saved state could not be verified.";
                SetResultStatus(UiText.Get(lastSaveError ?? "Saved."));
            }
            catch (Exception error)
            {
                lastSaveError = error.Message;
                PreserveDrafts();
                throw;
            }
            finally { busy = false; }
        }

        /// <summary>Reads the adapter's host-saved status for the module's project.</summary>
        /// <param name="native">Native module identifying the VBE and project to inspect.</param>
        /// <returns>The host's saved flag when persistence status is available; otherwise <see langword="null"/>.</returns>
        private static bool? ReadHostSaved(EditorVbeModule native)
        {
            dynamic state = new VbeProjectComponents(native.Vbe, null).PersistenceStatus(native.ProjectName);
            return (bool)state.HostAvailable ? (bool?)state.HostSaved : null;
        }

        /// <summary>Writes the editor buffer back to the corresponding VBE code module.</summary>
        /// <param name="native">Native module whose project must be saved by its own VBE instance.</param>
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
