using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace VBAi
{
    internal sealed partial class VbeProjectComponents
    {
        /// <summary>Native boundaries for an existing SOLIDWORKS host-project SWP save.</summary>
        internal interface ISolidWorksSaveProbe
        {
            bool IsSolidWorks { get; }
            int ProcessId { get; }
            void RequireOwner(object editor);
            bool SameProject(object first, object second);
            object SelectComponent(object editor, object project);
            SolidWorksSaveSelection Selection { get; }
            void RestoreSelection(object editor, object project, object component);
            bool SelectionMatches(object editor, object project, object component);
            object SaveControl(object editor);
            void Save(object control);
            bool FileExists(string path);
            bool FileReadOnly(string path);
            long FileLength(string path);
        }
        /// <summary>Reports only selection changes owned by this save operation.</summary>
        internal sealed class SolidWorksSaveSelection
        {
            public bool Changed { get; set; }
            public bool Restored { get; set; }
            public string Reason { get; set; }
        }
        /// <summary>Creates only an in-process probe; never selects another host from the ROT.</summary>
        internal Func<ISolidWorksSaveProbe> SolidWorksSaveProbe = () => new NativeSolidWorksSaveProbe();
        internal TimeSpan SolidWorksSaveVerificationTimeout = TimeSpan.FromSeconds(3);
        // Bridge, editor and chat have separate sessions on the same VBE owner thread.
        [ThreadStatic] private static bool solidWorksSavePending;

        /// <summary>Yields only for the SOLIDWORKS adapter; other hosts retain their native save contract.</summary>
        internal Task<object> SaveHostDocumentAsync(Request request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.ExpectedHostPath) ||
                !Path.IsPathRooted(request.ExpectedHostPath))
                throw new ArgumentException("ExpectedHostPath must be the absolute path read from project_persistence_status.");
            if (!host.IsExcel)
            {
                var native = SolidWorksSaveProbe();
                if (native.IsSolidWorks && (int)GetProject(request.Project).Type == 100)
                    return SaveSolidWorksMacroAsync(request, native);
            }
            return Task.FromResult(SaveHostDocument(request));
        }

        /// <summary>Preserves project metadata/references across Save, allowing only Saved to transition.</summary>
        private string SolidWorksSaveRevision(string selector)
        {
            dynamic state = ProjectProperties(selector);
            var properties = ((System.Collections.Generic.IEnumerable<VbePropertyInfo>)state.Properties)
                .Where(property => !string.Equals(property.Name, "Saved", StringComparison.OrdinalIgnoreCase)).ToArray();
            return Hash(json.Serialize(new { Mode = (int)state.Mode, Properties = properties,
                Components = (object)state.Components, References = (object)state.References }));
        }

        private sealed class NativeSolidWorksSaveProbe : ISolidWorksSaveProbe
        {
            [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
            private readonly bool isSolidWorks;
            private readonly int processId;
            private object previousPane, selectedPane;
            public SolidWorksSaveSelection Selection { get; } = new SolidWorksSaveSelection();
            internal NativeSolidWorksSaveProbe()
            {
                using (var process = Process.GetCurrentProcess())
                { processId = process.Id; isSolidWorks = string.Equals(process.ProcessName, "SLDWORKS", StringComparison.OrdinalIgnoreCase); }
            }
            public bool IsSolidWorks => isSolidWorks;
            public int ProcessId => processId;
            public void RequireOwner(object editor)
            {
                if (!IsSolidWorks) throw new InvalidOperationException("This save adapter requires SOLIDWORKS.");
                IntPtr window = new IntPtr(Convert.ToInt64(((dynamic)editor).MainWindow.HWnd));
                uint owner;
                uint thread = GetWindowThreadProcessId(window, out owner);
                if (window == IntPtr.Zero || owner != (uint)ProcessId || thread != GetCurrentThreadId())
                    throw new InvalidOperationException("The VBE window must belong to this SOLIDWORKS process and owning UI thread.");
            }
            public bool SameProject(object first, object second)
            {
                if (first == null || second == null) return false;
                IntPtr a = Marshal.GetIUnknownForObject(first), b = IntPtr.Zero;
                try { b = Marshal.GetIUnknownForObject(second); return a == b; }
                finally { if (b != IntPtr.Zero) Marshal.Release(b); Marshal.Release(a); }
            }
            public object SelectComponent(object editor, object project)
            {
                previousPane = (object)((dynamic)editor).ActiveCodePane;
                if (previousPane != null)
                    foreach (dynamic component in ((dynamic)project).VBComponents)
                        if (SameProject((object)component.CodeModule, (object)((dynamic)previousPane).CodeModule))
                            return component; // Preserve the existing pane, caret and focus.
                foreach (dynamic component in ((dynamic)project).VBComponents)
                {
                    int type = (int)component.Type;
                    if (type != 1 && type != 2 && type != 3 && type != 100) continue;
                    dynamic pane = component.CodeModule.CodePane;
                    selectedPane = pane; Selection.Changed = true;
                    pane.Show(); ((dynamic)editor).ActiveCodePane = pane; pane.Window.SetFocus();
                    return component;
                }
                throw new InvalidOperationException("The SWP project has no selectable code component.");
            }
            public void RestoreSelection(object editor, object project, object component)
            {
                if (!Selection.Changed) return;
                try
                {
                    if (previousPane == null) { Selection.Reason = "No previous code pane to restore."; return; }
                    if ((int)((dynamic)project).Mode != 2 ||
                        !SameProject(selectedPane, (object)((dynamic)editor).ActiveCodePane))
                    { Selection.Reason = "The selection or execution mode changed; user context was left untouched."; return; }
                    // Do not Show, SetFocus or reset the user's selection/caret.
                    ((dynamic)editor).ActiveCodePane = (dynamic)previousPane;
                    Selection.Restored = SameProject(previousPane, (object)((dynamic)editor).ActiveCodePane);
                    if (!Selection.Restored) Selection.Reason = "The previous code pane was not restored.";
                }
                catch (Exception error) { Selection.Reason = "Save selection restoration failed: " + error.Message; }
            }
            public bool SelectionMatches(object editor, object project, object component)
            {
                dynamic pane = ((dynamic)editor).ActiveCodePane;
                if (pane == null || !SameProject(project, (object)((dynamic)editor).ActiveVBProject) ||
                    !SameProject(component, (object)pane.CodeModule.Parent)) return false;
                foreach (dynamic item in ((dynamic)project).VBComponents)
                    if (SameProject(component, (object)item)) return true;
                return false;
            }
            public object SaveControl(object editor)
            {
                dynamic control = ((dynamic)editor).CommandBars.FindControl(1, 3);
                if (control == null || (int)control.Id != 3 || !(bool)control.BuiltIn || !(bool)control.Enabled)
                    throw new InvalidOperationException("The built-in VBE Save command is unavailable.");
                return control;
            }
            public void Save(object control) { ((dynamic)control).Execute(); }
            public bool FileExists(string path) => File.Exists(path);
            public bool FileReadOnly(string path) => (File.GetAttributes(path) & FileAttributes.ReadOnly) != 0;
            public long FileLength(string path) => new FileInfo(path).Length;
        }
        /// <summary>Reads existing Type100 SWP state without implying that disk reload was tested.</summary>
        private object SolidWorksPersistence(string selector, object projectObject, ISolidWorksSaveProbe native)
        {
            dynamic project = projectObject;
            try
            {
                native.RequireOwner((object)vbe);
                string path = SolidWorksMacroPath(projectObject);
                bool exists = native.FileExists(path);
                return new { Project = selector, ProjectSaved = (bool)project.Saved, Host = "SOLIDWORKS",
                    HostAvailable = true, HostPath = path, HostSaved = exists ? (bool?)(bool)project.Saved : null,
                    HostReadOnly = exists ? (bool?)native.FileReadOnly(path) : null, HostHasPath = (bool?)true,
                    FileExists = exists, OwnerProcessId = native.ProcessId, SaveApi = "VBE.CommandBars.ID3",
                    NativeQualification = "NOT_RUN", PersistenceReopenVerified = false, Reason = (string)null };
            }
            catch (Exception error)
            {
                return new { Project = selector, ProjectSaved = (bool)project.Saved, Host = "SOLIDWORKS",
                    HostAvailable = false, HostPath = (string)null, HostSaved = (bool?)null,
                    HostReadOnly = (bool?)null, HostHasPath = (bool?)null, Reason = error.Message };
            }
        }
        private static string SolidWorksMacroPath(object projectObject)
        {
            dynamic project = projectObject;
            if ((int)project.Type != 100) throw new InvalidOperationException("This adapter requires a SOLIDWORKS host project (Type=100).");
            string path = (string)project.FileName;
            if (string.IsNullOrWhiteSpace(path) || !Path.IsPathRooted(path) ||
                !string.Equals(Path.GetExtension(path), ".swp", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("An existing absolute SWP path is required; host-project SaveAs is not supported.");
            return Path.GetFullPath(path);
        }
        /// <summary>Invokes the selected native Save once, preserving failures after invocation as uncertain.</summary>
        internal Task<object> SaveSolidWorksMacroAsync(Request request, ISolidWorksSaveProbe native)
        {
            return VbeUiTask.Run(() => SaveSolidWorksMacroCoreAsync(request, native));
        }

        private async Task<object> SaveSolidWorksMacroCoreAsync(Request request, ISolidWorksSaveProbe native)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.ExpectedProjectVersion)) throw new ArgumentException("ExpectedProjectVersion is required.");
            if (native == null || !native.IsSolidWorks) throw new InvalidOperationException("This save adapter requires SOLIDWORKS.");
            native.RequireOwner((object)vbe);
            if (solidWorksSavePending) throw new InvalidOperationException("A SOLIDWORKS save is already awaiting verification; no second save was invoked.");
            string selector = request.Project;
            object projectObject = GetDesignProject(selector);
            dynamic project = projectObject;
            string path = RequireAbsolutePath(request.ExpectedHostPath);
            RequireSolidWorksSaveState(request, projectObject, path, native);
            int processId = native.ProcessId;
            string sourceSha = OtherHostSourceSha(projectObject);
            string revision = SolidWorksSaveRevision(selector);
            object component = null;
            solidWorksSavePending = true;
            try
            {
                component = native.SelectComponent((object)vbe, projectObject);
                object control = native.SaveControl((object)vbe);
                RequireSolidWorksSaveState(request, projectObject, path, native);
                if (native.ProcessId != processId || !native.SelectionMatches((object)vbe, projectObject, component) || OtherHostSourceSha(projectObject) != sourceSha)
                    throw new InvalidOperationException("The selected SWP project, host process or live source changed before Save.");
                try
                {
                    native.Save(control); // No fallback or retry beyond this mutation boundary.
                    var elapsed = Stopwatch.StartNew();
                    // CommandBars.Execute can return before SOLIDWORKS processes Save.
                    // Return to the owning message loop; never pump it reentrantly or replay Save.
                    while (true)
                    {
                        await Task.Delay(25);
                        native.RequireOwner((object)vbe);
                        string mismatch = (int)project.Mode != 2 ? "execution mode changed" :
                        native.ProcessId != processId ? "host process changed" :
                        !native.SameProject(projectObject, (object)GetProject(selector)) ? "project identity changed" :
                        !native.SelectionMatches((object)vbe, projectObject, component) ? "native selection changed" :
                        !string.Equals(SolidWorksMacroPath(projectObject), path, StringComparison.OrdinalIgnoreCase) ? "host path changed" :
                        OtherHostSourceSha(projectObject) != sourceSha ? "live VBA source changed" :
                        SolidWorksSaveRevision(selector) != revision ? "project metadata or references changed" : null;
                        if (mismatch != null)
                            throw new InvalidOperationException("Native SWP Save verification failed: " + mismatch + ".");
                        if (!(bool)project.Saved)
                        {
                            if (elapsed.Elapsed < SolidWorksSaveVerificationTimeout) continue;
                            throw new InvalidOperationException("Native SWP Save verification timed out: project saved flag is still false.");
                        }
                        mismatch =
                        !native.FileExists(path) ? "SWP file is absent" :
                        native.FileLength(path) < 1 ? "SWP file is empty" : null;
                        if (mismatch != null)
                            throw new InvalidOperationException("Native SWP Save verification failed: " + mismatch + ".");
                        break;
                    }
                    return new { Project = selector, Host = "SOLIDWORKS", HostPath = path, SaveApi = "VBE.CommandBars.ID3",
                        SaveInvoked = true, SaveAsInvoked = false, MutationInvoked = true, Verified = true, Uncertain = false,
                        ProjectSaved = true, HostSaved = true, Bytes = native.FileLength(path), OwnerProcessId = processId,
                        SourceSha256 = sourceSha, CodePreserved = true, NativeSelection = native.Selection, NativeQualification = "NOT_RUN", PersistenceReopenVerified = false,
                        Verification = "DeferredOwnerThreadSavedReadback", VerificationMilliseconds = elapsed.ElapsedMilliseconds,
                        Limit = "Native command, saved flags, file presence and unchanged live code were observed. Reopen the SWP to verify persisted code, resources and signatures." };
                }
                catch (Exception error)
                {
                    return new { Project = selector, Host = "SOLIDWORKS", HostPath = path, SaveApi = "VBE.CommandBars.ID3",
                        SaveInvoked = true, SaveAsInvoked = false, MutationInvoked = true, Verified = false, Uncertain = true,
                        NativeSelection = native.Selection, Reason = error.Message, Next = "Inspect project_persistence_status and the SWP file; do not retry automatically." };
                }
            }
            finally
            {
                try { native.RestoreSelection((object)vbe, projectObject, component); }
                finally { solidWorksSavePending = false; }
            }
        }
        private void RequireSolidWorksSaveState(Request request, object projectObject, string path, ISolidWorksSaveProbe native)
        {
            native.RequireOwner((object)vbe);
            dynamic project = projectObject;
            if (!native.SameProject(projectObject, (object)GetDesignProject(request.Project)) || (int)project.Protection != 0)
                throw new InvalidOperationException("The SWP project identity or protection changed.");
            AssertProjectVersion(request, project);
            if (!string.Equals(SolidWorksMacroPath(projectObject), path, StringComparison.OrdinalIgnoreCase) ||
                !native.FileExists(path) || native.FileLength(path) < 1 || native.FileReadOnly(path))
                throw new InvalidOperationException("The approved SWP path is changed, absent, empty or read-only.");
        }
    }
}
