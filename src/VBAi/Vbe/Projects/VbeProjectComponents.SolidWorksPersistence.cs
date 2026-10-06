using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace VBAi
{

    /// <summary>Owns the vbe project components state and operations.</summary>
    internal sealed partial class VbeProjectComponents
    {

        /// <summary>Native boundaries for an existing SOLIDWORKS host-project SWP save.</summary>
        internal interface ISolidWorksSaveProbe
        {

            /// <summary>Gets the is solid works.</summary>
            /// <value>Current is solid works exposed by i solid works save probe.</value>
            bool IsSolidWorks { get; }

            /// <summary>Gets the process id.</summary>
            /// <value>Current process id exposed by i solid works save probe.</value>
            int ProcessId { get; }

            /// <summary>Requires owner for i solid works save probe.</summary>
            /// <param name="editor">object that supplies the editor for this operation.</param>
            void RequireOwner(object editor);

            /// <summary>Compares project for i solid works save probe.</summary>
            /// <param name="first">object that supplies the first for this operation.</param>
            /// <param name="second">object that supplies the second for this operation.</param>
            /// <returns>Boolean indicating the result of the check for same project on i solid works save probe.</returns>
            bool SameProject(object first, object second);

            /// <summary>Handles select component for i solid works save probe.</summary>
            /// <param name="editor">object that supplies the editor for this operation.</param>
            /// <param name="project">object that supplies the project for this operation.</param>
            /// <returns>object produced by the operation for select component on i solid works save probe.</returns>
            object SelectComponent(object editor, object project);

            /// <summary>Gets the selection.</summary>
            /// <value>Current selection exposed by i solid works save probe.</value>
            SolidWorksSaveSelection Selection { get; }

            /// <summary>Handles restore selection for i solid works save probe.</summary>
            /// <param name="editor">object that supplies the editor for this operation.</param>
            /// <param name="project">object that supplies the project for this operation.</param>
            /// <param name="component">object that supplies the component for this operation.</param>
            void RestoreSelection(object editor, object project, object component);

            /// <summary>Handles selection matches for i solid works save probe.</summary>
            /// <param name="editor">object that supplies the editor for this operation.</param>
            /// <param name="project">object that supplies the project for this operation.</param>
            /// <param name="component">object that supplies the component for this operation.</param>
            /// <returns>Boolean indicating the result of the check for selection matches on i solid works save probe.</returns>
            bool SelectionMatches(object editor, object project, object component);

            /// <summary>Saves control for i solid works save probe.</summary>
            /// <param name="editor">object that supplies the editor for this operation.</param>
            /// <returns>object produced by the operation for save control on i solid works save probe.</returns>
            object SaveControl(object editor);

            /// <summary>Saves  for i solid works save probe.</summary>
            /// <param name="control">object that supplies the control for this operation.</param>
            void Save(object control);

            /// <summary>Handles file exists for i solid works save probe.</summary>
            /// <param name="path">Path used for the path being processed.</param>
            /// <returns>Boolean indicating the result of the check for file exists on i solid works save probe.</returns>
            bool FileExists(string path);

            /// <summary>Handles file read only for i solid works save probe.</summary>
            /// <param name="path">Path used for the path being processed.</param>
            /// <returns>Boolean indicating the result of the check for file read only on i solid works save probe.</returns>
            bool FileReadOnly(string path);

            /// <summary>Handles file length for i solid works save probe.</summary>
            /// <param name="path">Path used for the path being processed.</param>
            /// <returns>long produced by the operation for file length on i solid works save probe.</returns>
            long FileLength(string path);
        }

        /// <summary>Reports only selection changes owned by this save operation.</summary>
        internal sealed class SolidWorksSaveSelection
        {

            /// <summary>Gets or sets the changed.</summary>
            /// <value>Current changed exposed by solid works save selection.</value>
            public bool Changed { get; set; }

            /// <summary>Gets or sets the restored.</summary>
            /// <value>Current restored exposed by solid works save selection.</value>
            public bool Restored { get; set; }

            /// <summary>Gets or sets the reason.</summary>
            /// <value>Current reason exposed by solid works save selection.</value>
            public string Reason { get; set; }
        }

        /// <summary>Creates only an in-process probe; never selects another host from the ROT.</summary>
        internal Func<ISolidWorksSaveProbe> SolidWorksSaveProbe = () => new NativeSolidWorksSaveProbe();

        /// <summary>Maintains the solid works save verification timeout state for vbe project components.</summary>
        internal TimeSpan SolidWorksSaveVerificationTimeout = TimeSpan.FromSeconds(3);
        // Bridge, editor and chat have separate sessions on the same VBE owner thread.
        /// <summary>Maintains the solid works save pending state for vbe project components.</summary>
        [ThreadStatic] private static bool solidWorksSavePending;

        /// <summary>Yields for SOLIDWORKS and Access delayed Saved notifications on their owning VBE thread.</summary>
        /// <param name="request">request that supplies the request for this operation.</param>
        /// <returns>task&lt;object&gt; produced by the operation for save host document async on vbe project components.</returns>
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
                var other = OtherHostProbe();
                if (other.HostKind == "Access") return SaveAccessDocumentAsync(request, other);
            }
            return Task.FromResult(SaveHostDocument(request));
        }

        /// <summary>Preserves project metadata/references across Save, allowing only Saved to transition.</summary>
        /// <param name="selector">Text that supplies the selector value. Use the format required by the calling operation.</param>
        /// <returns>Text produced by the operation for solid works save revision on vbe project components.</returns>
        private string SolidWorksSaveRevision(string selector)
        {
            dynamic state = ProjectProperties(selector);
            var properties = ((System.Collections.Generic.IEnumerable<VbePropertyInfo>)state.Properties)
                .Where(property => !string.Equals(property.Name, "Saved", StringComparison.OrdinalIgnoreCase)).ToArray();
            return Hash(json.Serialize(new { Mode = (int)state.Mode, Properties = properties,
                Components = (object)state.Components, References = (object)state.References }));
        }

        /// <summary>Owns the native solid works save probe state and operations.</summary>
        private sealed class NativeSolidWorksSaveProbe : ISolidWorksSaveProbe
        {

            /// <summary>Returns current thread id for native solid works save probe.</summary>
            /// <returns>uint produced by the operation for get current thread id on native solid works save probe.</returns>
            [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();

            /// <summary>Tracks the is solid works state of native solid works save probe.</summary>
            private readonly bool isSolidWorks;

            /// <summary>Identifies the process id associated with native solid works save probe.</summary>
            private readonly int processId;

            /// <summary>Maintains the previous pane and selected pane state for native solid works save probe.</summary>
            private object previousPane, selectedPane;

            /// <summary>Gets the selection.</summary>
            /// <value>Current selection exposed by native solid works save probe.</value>
            public SolidWorksSaveSelection Selection { get; } = new SolidWorksSaveSelection();

            /// <summary>Initializes a NativeSolidWorksSaveProbe instance with the supplied state.</summary>
            internal NativeSolidWorksSaveProbe()
            {
                using (var process = Process.GetCurrentProcess())
                { processId = process.Id; isSolidWorks = string.Equals(process.ProcessName, "SLDWORKS", StringComparison.OrdinalIgnoreCase); }
            }

            /// <summary>Gets the is solid works.</summary>
            /// <value>Current is solid works exposed by native solid works save probe.</value>
            public bool IsSolidWorks => isSolidWorks;

            /// <summary>Gets the process id.</summary>
            /// <value>Current process id exposed by native solid works save probe.</value>
            public int ProcessId => processId;

            /// <summary>Requires owner for native solid works save probe.</summary>
            /// <param name="editor">object that supplies the editor for this operation.</param>
            public void RequireOwner(object editor)
            {
                if (!IsSolidWorks) throw new InvalidOperationException("This save adapter requires SOLIDWORKS.");
                IntPtr window = new IntPtr(Convert.ToInt64(((dynamic)editor).MainWindow.HWnd));
                uint owner;
                uint thread = GetWindowThreadProcessId(window, out owner);
                if (window == IntPtr.Zero || owner != (uint)ProcessId || thread != GetCurrentThreadId())
                    throw new InvalidOperationException("The VBE window must belong to this SOLIDWORKS process and owning UI thread.");
            }

            /// <summary>Compares project for native solid works save probe.</summary>
            /// <param name="first">object that supplies the first for this operation.</param>
            /// <param name="second">object that supplies the second for this operation.</param>
            /// <returns>Boolean indicating the result of the check for same project on native solid works save probe.</returns>
            public bool SameProject(object first, object second)
            {
                if (first == null || second == null) return false;
                IntPtr a = Marshal.GetIUnknownForObject(first), b = IntPtr.Zero;
                try { b = Marshal.GetIUnknownForObject(second); return a == b; }
                finally { if (b != IntPtr.Zero) Marshal.Release(b); Marshal.Release(a); }
            }

            /// <summary>Handles select component for native solid works save probe.</summary>
            /// <param name="editor">object that supplies the editor for this operation.</param>
            /// <param name="project">object that supplies the project for this operation.</param>
            /// <returns>object produced by the operation for select component on native solid works save probe.</returns>
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

            /// <summary>Handles restore selection for native solid works save probe.</summary>
            /// <param name="editor">object that supplies the editor for this operation.</param>
            /// <param name="project">object that supplies the project for this operation.</param>
            /// <param name="component">object that supplies the component for this operation.</param>
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

            /// <summary>Handles selection matches for native solid works save probe.</summary>
            /// <param name="editor">object that supplies the editor for this operation.</param>
            /// <param name="project">object that supplies the project for this operation.</param>
            /// <param name="component">object that supplies the component for this operation.</param>
            /// <returns>Boolean indicating the result of the check for selection matches on native solid works save probe.</returns>
            public bool SelectionMatches(object editor, object project, object component)
            {
                dynamic pane = ((dynamic)editor).ActiveCodePane;
                if (pane == null || !SameProject(project, (object)((dynamic)editor).ActiveVBProject) ||
                    !SameProject(component, (object)pane.CodeModule.Parent)) return false;
                foreach (dynamic item in ((dynamic)project).VBComponents)
                    if (SameProject(component, (object)item)) return true;
                return false;
            }

            /// <summary>Saves control for native solid works save probe.</summary>
            /// <param name="editor">object that supplies the editor for this operation.</param>
            /// <returns>object produced by the operation for save control on native solid works save probe.</returns>
            public object SaveControl(object editor)
            {
                dynamic control = ((dynamic)editor).CommandBars.FindControl(1, 3);
                if (control == null || (int)control.Id != 3 || !(bool)control.BuiltIn || !(bool)control.Enabled)
                    throw new InvalidOperationException("The built-in VBE Save command is unavailable.");
                return control;
            }

            /// <summary>Saves  for native solid works save probe.</summary>
            /// <param name="control">object that supplies the control for this operation.</param>
            public void Save(object control) { ((dynamic)control).Execute(); }

            /// <summary>Handles file exists for native solid works save probe.</summary>
            /// <param name="path">Path used for the path being processed.</param>
            /// <returns>Boolean indicating the result of the check for file exists on native solid works save probe.</returns>
            public bool FileExists(string path) => File.Exists(path);

            /// <summary>Handles file read only for native solid works save probe.</summary>
            /// <param name="path">Path used for the path being processed.</param>
            /// <returns>Boolean indicating the result of the check for file read only on native solid works save probe.</returns>
            public bool FileReadOnly(string path) => (File.GetAttributes(path) & FileAttributes.ReadOnly) != 0;

            /// <summary>Handles file length for native solid works save probe.</summary>
            /// <param name="path">Path used for the path being processed.</param>
            /// <returns>long produced by the operation for file length on native solid works save probe.</returns>
            public long FileLength(string path) => new FileInfo(path).Length;
        }

        /// <summary>Reads existing Type100 SWP state without implying that disk reload was tested.</summary>
        /// <param name="selector">Text that supplies the selector value. Use the format required by the calling operation.</param>
        /// <param name="projectObject">object that supplies the project object for this operation.</param>
        /// <param name="native">i solid works save probe that supplies the native for this operation.</param>
        /// <returns>object produced by the operation for solid works persistence on vbe project components.</returns>
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

        /// <summary>Handles solid works macro path for vbe project components.</summary>
        /// <param name="projectObject">object that supplies the project object for this operation.</param>
        /// <returns>Text produced by the operation for solid works macro path on vbe project components.</returns>
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
        /// <param name="request">request that supplies the request for this operation.</param>
        /// <param name="native">i solid works save probe that supplies the native for this operation.</param>
        /// <returns>task&lt;object&gt; produced by the operation for save solid works macro async on vbe project components.</returns>
        internal Task<object> SaveSolidWorksMacroAsync(Request request, ISolidWorksSaveProbe native)
        {
            return VbeUiTask.Run(() => SaveSolidWorksMacroCoreAsync(request, native));
        }

        /// <summary>Saves solid works macro core async for vbe project components.</summary>
        /// <param name="request">request that supplies the request for this operation.</param>
        /// <param name="native">i solid works save probe that supplies the native for this operation.</param>
        /// <returns>task&lt;object&gt; produced by the operation for save solid works macro core async on vbe project components.</returns>
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

        /// <summary>Requires solid works save state for vbe project components.</summary>
        /// <param name="request">request that supplies the request for this operation.</param>
        /// <param name="projectObject">object that supplies the project object for this operation.</param>
        /// <param name="path">Path used for the path being processed.</param>
        /// <param name="native">i solid works save probe that supplies the native for this operation.</param>
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
