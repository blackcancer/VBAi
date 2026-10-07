using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace VBAi
{

    /// <summary>Reports persistence state and coordinates existing SOLIDWORKS/Access host-project saves.</summary>
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

            /// <summary>Requires the native host window, process, and owner-thread identity to remain current.</summary>
            /// <param name="editor">VBE editor whose native host window is checked.</param>
            void RequireOwner(object editor);

            /// <summary>Compares canonical COM identity for a project across host transitions.</summary>
            /// <param name="first">Previously captured project COM object.</param>
            /// <param name="second">Current project COM object.</param>
            /// <returns>True when both references identify the same project.</returns>
            bool SameProject(object first, object second);

            /// <summary>Selects the component needed to route the VBE Save command to the host project.</summary>
            /// <param name="editor">VBE editor that owns the active code pane.</param>
            /// <param name="project">Canonical SOLIDWORKS Type100 project to save.</param>
            /// <returns>Selected component handle retained for verification and selection restoration.</returns>
            object SelectComponent(object editor, object project);

            /// <summary>Gets the selection.</summary>
            /// <value>Current selection exposed by i solid works save probe.</value>
            SolidWorksSaveSelection Selection { get; }

            /// <summary>Restores the editor's prior code-pane selection after save verification.</summary>
            /// <param name="editor">VBE editor whose selection was changed by this operation.</param>
            /// <param name="project">Project containing the selected component.</param>
            /// <param name="component">Component whose code pane was temporarily selected.</param>
            void RestoreSelection(object editor, object project, object component);

            /// <summary>Checks that the active code pane still belongs to the selected project component.</summary>
            /// <param name="editor">VBE editor with the current active code pane.</param>
            /// <param name="project">Expected project identity.</param>
            /// <param name="component">Expected selected component identity.</param>
            /// <returns>True only when pane, project, and component identities all match.</returns>
            bool SelectionMatches(object editor, object project, object component);

            /// <summary>Resolves the built-in VBE Save command and requires it to be enabled.</summary>
            /// <param name="editor">VBE editor whose CommandBars collection is queried.</param>
            /// <returns>Built-in command control with ID 3.</returns>
            object SaveControl(object editor);

            /// <summary>Executes the selected native Save command once.</summary>
            /// <param name="control">Built-in VBE Save command control returned by <see cref="SaveControl"/>.</param>
            void Save(object control);

            /// <summary>Checks whether the host project file currently exists.</summary>
            /// <param name="path">Absolute SOLIDWORKS macro project path.</param>
            /// <returns>File existence at the time of the query.</returns>
            bool FileExists(string path);

            /// <summary>Checks the read-only attribute on an existing host project file.</summary>
            /// <param name="path">Absolute SOLIDWORKS macro project path.</param>
            /// <returns>Whether the file has the ReadOnly attribute.</returns>
            bool FileReadOnly(string path);

            /// <summary>Reads the current byte length of the host project file.</summary>
            /// <param name="path">Absolute SOLIDWORKS macro project path.</param>
            /// <returns>File length in bytes.</returns>
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

        /// <summary>Maximum wait for the host's delayed Saved notification after one SOLIDWORKS Save command.</summary>
        internal TimeSpan SolidWorksSaveVerificationTimeout = TimeSpan.FromSeconds(3);
        // Bridge, editor and chat have separate sessions on the same VBE owner thread.
        /// <summary>Thread-local guard preventing a second SWP save while the current mutation outcome is unsettled.</summary>
        [ThreadStatic] private static bool solidWorksSavePending;

        /// <summary>Yields for SOLIDWORKS and Access delayed Saved notifications on their owning VBE thread.</summary>
        /// <param name="request">Request carrying the absolute host path previously observed from persistence status.</param>
        /// <returns>Asynchronous host save result for supported SOLIDWORKS/Access adapters, or a completed ordinary save result.</returns>
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
        /// <param name="selector">Project selector used by the current project-property reader.</param>
        /// <returns>Hash of mode, properties except the changing Saved flag, components, and references.</returns>
        private string SolidWorksSaveRevision(string selector)
        {
            dynamic state = ProjectProperties(selector);
            var properties = ((System.Collections.Generic.IEnumerable<VbePropertyInfo>)state.Properties)
                .Where(property => !string.Equals(property.Name, "Saved", StringComparison.OrdinalIgnoreCase)).ToArray();
            return Hash(json.Serialize(new
            {
                Mode = (int)state.Mode,
                Properties = properties,
                Components = (object)state.Components,
                References = (object)state.References
            }));
        }

        /// <summary>Performs existing-project SOLIDWORKS Save checks on the owning VBE thread.</summary>
        private sealed class NativeSolidWorksSaveProbe : ISolidWorksSaveProbe
        {

            /// <summary>Reads the native thread ID used to enforce host-window ownership.</summary>
            /// <returns>Current Win32 thread ID.</returns>
            [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();

            /// <summary>Whether construction resolved this probe inside a SOLIDWORKS process.</summary>
            private readonly bool isSolidWorks;

            /// <summary>Host process ID captured by this native probe.</summary>
            private readonly int processId;

            /// <summary>Prior editor pane and temporary save-target pane retained for restoring user selection.</summary>
            private object previousPane, selectedPane;

            /// <summary>Gets the selection.</summary>
            /// <value>Current selection exposed by native solid works save probe.</value>
            public SolidWorksSaveSelection Selection { get; } = new SolidWorksSaveSelection();

            /// <summary>Captures current host process and VBE owner-thread identity for the native save probe.</summary>
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

            /// <summary>Requires the current process and native VBE window to remain owned by the captured thread.</summary>
            /// <param name="editor">VBE editor whose main-window identity is checked.</param>
            public void RequireOwner(object editor)
            {
                if (!IsSolidWorks) throw new InvalidOperationException("This save adapter requires SOLIDWORKS.");
                IntPtr window = new IntPtr(Convert.ToInt64(((dynamic)editor).MainWindow.HWnd));
                uint thread = GetWindowThreadProcessId(window, out uint owner);
                if (window == IntPtr.Zero || owner != (uint)ProcessId || thread != GetCurrentThreadId())
                    throw new InvalidOperationException("The VBE window must belong to this SOLIDWORKS process and owning UI thread.");
            }

            /// <summary>Compares canonical project COM identity.</summary>
            /// <param name="first">Previously captured project object.</param>
            /// <param name="second">Current project object.</param>
            /// <returns>Whether both objects identify the same project.</returns>
            public bool SameProject(object first, object second)
            {
                if (first == null || second == null) return false;
                IntPtr a = Marshal.GetIUnknownForObject(first), b = IntPtr.Zero;
                try { b = Marshal.GetIUnknownForObject(second); return a == b; }
                finally { if (b != IntPtr.Zero) Marshal.Release(b); Marshal.Release(a); }
            }

            /// <summary>Selects an existing project component so the VBE Save command targets the host project.</summary>
            /// <param name="editor">VBE editor whose active code pane is saved and later restored.</param>
            /// <param name="project">Canonical SOLIDWORKS Type100 project.</param>
            /// <returns>Selected component object retained for selection verification.</returns>
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

            /// <summary>Restores the prior VBE code-pane selection after verification.</summary>
            /// <param name="editor">VBE editor whose pane selection changed.</param>
            /// <param name="project">Project containing the temporary save target.</param>
            /// <param name="component">Component whose pane was temporarily selected.</param>
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

            /// <summary>Requires the active code pane, project, and component COM identities to match the save target.</summary>
            /// <param name="editor">VBE editor with the current active code pane.</param>
            /// <param name="project">Expected project identity.</param>
            /// <param name="component">Expected selected component.</param>
            /// <returns>True only when the current pane and project collection confirm that component.</returns>
            public bool SelectionMatches(object editor, object project, object component)
            {
                dynamic pane = ((dynamic)editor).ActiveCodePane;
                if (pane == null || !SameProject(project, (object)((dynamic)editor).ActiveVBProject) ||
                    !SameProject(component, (object)pane.CodeModule.Parent)) return false;
                foreach (dynamic item in ((dynamic)project).VBComponents)
                    if (SameProject(component, (object)item)) return true;
                return false;
            }

            /// <summary>Resolves the enabled built-in VBE Save command control with ID 3.</summary>
            /// <param name="editor">VBE editor whose CommandBars are queried.</param>
            /// <returns>The enabled built-in Save control.</returns>
            public object SaveControl(object editor)
            {
                dynamic control = ((dynamic)editor).CommandBars.FindControl(1, 3);
                if (control == null || (int)control.Id != 3 || !(bool)control.BuiltIn || !(bool)control.Enabled)
                    throw new InvalidOperationException("The built-in VBE Save command is unavailable.");
                return control;
            }

            /// <summary>Executes the selected VBE Save control.</summary>
            /// <param name="control">Built-in Save command control resolved immediately before delivery.</param>
            public void Save(object control) { ((dynamic)control).Execute(); }

            /// <summary>Checks for the host project file at the supplied path.</summary>
            /// <param name="path">Absolute SWP path.</param>
            /// <returns>Whether the file currently exists.</returns>
            public bool FileExists(string path) => File.Exists(path);

            /// <summary>Checks the ReadOnly attribute on the existing host project file.</summary>
            /// <param name="path">Absolute SWP path.</param>
            /// <returns>Whether the file is marked read-only.</returns>
            public bool FileReadOnly(string path) => (File.GetAttributes(path) & FileAttributes.ReadOnly) != 0;

            /// <summary>Reads the current host file size.</summary>
            /// <param name="path">Absolute SWP path.</param>
            /// <returns>Current file length in bytes.</returns>
            public long FileLength(string path) => new FileInfo(path).Length;
        }

        /// <summary>Reads existing Type100 SWP state without implying that disk reload was tested.</summary>
        /// <param name="selector">Project selector used in the returned persistence report.</param>
        /// <param name="projectObject">Live Type100 project whose Saved flag and FileName are observed.</param>
        /// <param name="native">Owner-bound probe used for host path, process, and file-state reads.</param>
        /// <returns>Host persistence status; native qualification and reopen verification remain explicitly unclaimed.</returns>
        private object SolidWorksPersistence(string selector, object projectObject, ISolidWorksSaveProbe native)
        {
            dynamic project = projectObject;
            try
            {
                native.RequireOwner((object)vbe);
                string path = SolidWorksMacroPath(projectObject);
                bool exists = native.FileExists(path);
                return new
                {
                    Project = selector,
                    ProjectSaved = (bool)project.Saved,
                    Host = "SOLIDWORKS",
                    HostAvailable = true,
                    HostPath = path,
                    HostSaved = exists ? (bool?)(bool)project.Saved : null,
                    HostReadOnly = exists ? (bool?)native.FileReadOnly(path) : null,
                    HostHasPath = (bool?)true,
                    FileExists = exists,
                    OwnerProcessId = native.ProcessId,
                    SaveApi = "VBE.CommandBars.ID3",
                    NativeQualification = "NOT_RUN",
                    PersistenceReopenVerified = false,
                    Reason = (string)null
                };
            }
            catch (Exception error)
            {
                return new
                {
                    Project = selector,
                    ProjectSaved = (bool)project.Saved,
                    Host = "SOLIDWORKS",
                    HostAvailable = false,
                    HostPath = (string)null,
                    HostSaved = (bool?)null,
                    HostReadOnly = (bool?)null,
                    HostHasPath = (bool?)null,
                    Reason = error.Message
                };
            }
        }

        /// <summary>Reads and canonicalizes the existing absolute .swp path of a Type100 host project.</summary>
        /// <param name="projectObject">Live host project whose Type and FileName are checked.</param>
        /// <returns>Full host file path; missing, relative, wrong-extension, or non-Type100 paths throw.</returns>
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
        /// <param name="request">Expected project revision and exact host path approved for this existing-project save.</param>
        /// <param name="native">Native host probe used for owner checks, selection, Save delivery, and file verification.</param>
        /// <returns>Task completing with verified or uncertain save status; an uncertain save is never retried.</returns>
        internal Task<object> SaveSolidWorksMacroAsync(Request request, ISolidWorksSaveProbe native)
        {
            return VbeUiTask.Run(() => SaveSolidWorksMacroCoreAsync(request, native));
        }

        /// <summary>Saves solid works macro core async for vbe project components.</summary>
        /// <param name="request">Design project revision and expected absolute SWP path.</param>
        /// <param name="native">Owner-thread probe that selects the project, executes one Save, and checks host file state.</param>
        /// <returns>Save report after delayed Saved/file verification; post-invocation failures are returned as uncertain.</returns>
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
                    return new
                    {
                        Project = selector,
                        Host = "SOLIDWORKS",
                        HostPath = path,
                        SaveApi = "VBE.CommandBars.ID3",
                        SaveInvoked = true,
                        SaveAsInvoked = false,
                        MutationInvoked = true,
                        Verified = true,
                        Uncertain = false,
                        ProjectSaved = true,
                        HostSaved = true,
                        Bytes = native.FileLength(path),
                        OwnerProcessId = processId,
                        SourceSha256 = sourceSha,
                        CodePreserved = true,
                        NativeSelection = native.Selection,
                        NativeQualification = "NOT_RUN",
                        PersistenceReopenVerified = false,
                        Verification = "DeferredOwnerThreadSavedReadback",
                        VerificationMilliseconds = elapsed.ElapsedMilliseconds,
                        Limit = "Native command, saved flags, file presence and unchanged live code were observed. Reopen the SWP to verify persisted code, resources and signatures."
                    };
                }
                catch (Exception error)
                {
                    return new
                    {
                        Project = selector,
                        Host = "SOLIDWORKS",
                        HostPath = path,
                        SaveApi = "VBE.CommandBars.ID3",
                        SaveInvoked = true,
                        SaveAsInvoked = false,
                        MutationInvoked = true,
                        Verified = false,
                        Uncertain = true,
                        NativeSelection = native.Selection,
                        Reason = error.Message,
                        Next = "Inspect project_persistence_status and the SWP file; do not retry automatically."
                    };
                }
            }
            finally
            {
                try { native.RestoreSelection((object)vbe, projectObject, component); }
                finally { solidWorksSavePending = false; }
            }
        }

        /// <summary>Revalidates canonical project identity, protection, revision, and existing writable nonempty SWP path.</summary>
        /// <param name="request">Approved project selector and expected project revision.</param>
        /// <param name="projectObject">Frozen canonical Type100 project COM object.</param>
        /// <param name="path">Expected canonical absolute SWP path that must still match the project.</param>
        /// <param name="native">Owner-bound probe used to verify native ownership and file state.</param>
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
