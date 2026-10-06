using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VBAi
{

    /// <summary>Native live-project execution, independent of the host's Application.Run API.</summary>
    internal sealed class VbaNativeTestExecutionHost : IVbaTestExecutionHost, IDisposable
    {

        /// <summary>Defines the i probe contract.</summary>
        internal interface IProbe
        {

            /// <summary>Rejects access when the supplied VBE object is not being used on its owning thread.</summary>
            /// <param name="vbe">VBE automation object whose thread ownership is checked.</param>
            void RequireOwner(object vbe);

            /// <summary>Finds and prepares the native editor pane used to execute the approved procedure source.</summary>
            /// <param name="vbe">Owning VBE automation object.</param>
            /// <param name="project">Resolved project whose active code pane is prepared.</param>
            /// <param name="source">Expected module source used to validate the selected procedure.</param>
            /// <returns>Opaque preparation record required by subsequent revalidation and execution.</returns>
            object Prepare(object vbe, object project, string source);

            /// <summary>Rechecks selection, project, source, and pane identity immediately before native dispatch.</summary>
            /// <param name="vbe">Owning VBE automation object.</param>
            /// <param name="project">Project resolved for this attempt.</param>
            /// <param name="prepared">Preparation evidence returned by <see cref="Prepare"/>.</param>
            void Revalidate(object vbe, object project, object prepared);

            /// <summary>Performs the single native execution action after revalidation.</summary>
            /// <param name="prepared">Validated editor preparation record.</param>
            void Execute(object prepared);

            /// <summary>Reads the host's current execution mode to detect completion of the VBA call.</summary>
            /// <param name="project">Project whose host execution state is observed.</param>
            /// <returns>Host-reported mode value.</returns>
            int ReadMode(object project);
        }

        /// <summary>VBE automation object used only from the captured owning thread.</summary>
        private readonly object vbe;

        /// <summary>Correlates native result callbacks with this run's single outstanding call.</summary>
        private readonly VbaTestResultSink sink;

        /// <summary>Resolves the catalog's project selector to the live VBIDE project immediately before use.</summary>
        private readonly Func<VbaTestCatalog, object> resolveProject;

        /// <summary>Rechecks source revision, project identity, and authorization for a catalog.</summary>
        private readonly Action<VbaTestCatalog> validateCatalog;

        /// <summary>Checks current execution policy before each native attempt.</summary>
        private readonly Action executionGuard;

        /// <summary>Provides the expected support-module signature used to reject stale result messages.</summary>
        private readonly Func<VbaTestCatalog, string> signature;

        /// <summary>Provides the active runner correlation identifier for native result receipt matching.</summary>
        private readonly Func<string> runId;

        /// <summary>Managed thread identifier captured at construction and required for all COM operations.</summary>
        private readonly int owner = Thread.CurrentThread.ManagedThreadId;

        /// <summary>The only native procedure call currently awaiting a completion receipt.</summary>
        private Call active;

        /// <summary>Tracks shutdown and whether an unverified attempt permanently disabled this transport instance.</summary>
        private bool disposed, uncertain;

        /// <summary>Native editor adapter; replaceable internally for isolated execution-boundary verification.</summary>
        internal IProbe Probe = new NativeProbe();

        /// <summary>Posts the dispatch action to the owning WinForms thread.</summary>
        internal Action<Action> Post;

        /// <summary>Starts bounded completion polling and returns its disposable timer.</summary>
        internal Func<Action, IDisposable> StartPolling;

        /// <summary>Maximum time allowed to obtain positive native completion evidence.</summary>
        internal TimeSpan VerificationTimeout = TimeSpan.FromSeconds(30);

        /// <summary>Captures all identities and resources for one one-shot native invocation attempt.</summary>
        private sealed class Call
        {

            /// <summary>Catalog whose project and revision were validated before dispatch.</summary>
            internal VbaTestCatalog Catalog;

            /// <summary>Discovered descriptor expected in the result receipt.</summary>
            internal VbaTestDescriptor Test;

            /// <summary>Lifecycle phase recorded in the completion result.</summary>
            internal string Phase;

            /// <summary>Live project object resolved for this attempt.</summary>
            internal object Project;

            /// <summary>Tracks whether result correlation was armed and the irreversible native call was dispatched.</summary>
            internal bool Armed, Invoked;

            /// <summary>Elapsed-time source started when the attempt record is created.</summary>
            internal Stopwatch Clock = Stopwatch.StartNew();

            /// <summary>Temporary result exposure and completion timer, both released when the call settles.</summary>
            internal IDisposable Exposure, Polling;

            /// <summary>Asynchronous receipt completed once with verified success/failure or transport uncertainty.</summary>
            internal TaskCompletionSource<VbaTestResult> Completion = new TaskCompletionSource<VbaTestResult>(TaskCreationOptions.RunContinuationsAsynchronously);
        }

        /// <summary>Initializes a VbaNativeTestExecutionHost instance with the supplied state.</summary>
        /// <param name="vbe">Live VBE automation object captured on the owning thread.</param>
        /// <param name="dispatcher">Control with a created handle for posting native work back to that thread.</param>
        /// <param name="sink">Result sink that validates callback correlation.</param>
        /// <param name="resolveProject">Resolver for the project represented by the catalog.</param>
        /// <param name="validateCatalog">Revision and identity check repeated before dispatch.</param>
        /// <param name="executionGuard">Policy check that must permit each native call.</param>
        /// <param name="supportSignature">Expected signature of the injected result support module.</param>
        /// <param name="runId">Current run identifier used to reject stale callbacks.</param>
        /// <exception cref="ArgumentNullException">A required dependency is null.</exception>
        internal VbaNativeTestExecutionHost(object vbe, Control dispatcher, VbaTestResultSink sink,
            Func<VbaTestCatalog, object> resolveProject, Action<VbaTestCatalog> validateCatalog,
            Action executionGuard, Func<VbaTestCatalog, string> supportSignature, Func<string> runId)
        {
            this.vbe = vbe ?? throw new ArgumentNullException(nameof(vbe));
            this.sink = sink ?? throw new ArgumentNullException(nameof(sink));
            this.resolveProject = resolveProject ?? throw new ArgumentNullException(nameof(resolveProject));
            this.validateCatalog = validateCatalog ?? throw new ArgumentNullException(nameof(validateCatalog));
            this.executionGuard = executionGuard ?? throw new ArgumentNullException(nameof(executionGuard));
            signature = supportSignature ?? throw new ArgumentNullException(nameof(supportSignature));
            this.runId = runId ?? throw new ArgumentNullException(nameof(runId));
            if (dispatcher == null) throw new ArgumentNullException(nameof(dispatcher));
            Post = action => dispatcher.BeginInvoke(action);
            StartPolling = tick => { var timer = new System.Windows.Forms.Timer { Interval = 50 }; timer.Tick += (_, __) => tick(); timer.Start(); return timer; };
        }

        /// <summary>Checks thread affinity, live transport state, policy, VBE ownership, and project revision.</summary>
        /// <param name="catalog">Catalog whose captured identity and source revision must still be current.</param>
        public void Validate(VbaTestCatalog catalog)
        {
            RequireOwner();
            if (disposed || uncertain) throw new VbaTestInvocationException("The live native transport is disconnected or its preceding outcome is uncertain.", false);
            executionGuard();
            Probe.RequireOwner(vbe);
            validateCatalog(catalog);
        }

        /// <summary>Queues one serialized call and returns its completion receipt; a dispatched call is never retried.</summary>
        /// <param name="catalog">Validated catalog defining project and revision identity.</param>
        /// <param name="procedure">Discovered test or fixture to execute.</param>
        /// <param name="phase">Lifecycle phase attached to the result.</param>
        /// <returns>Task completed by the correlated native result sink or with a safe/uncertain failure.</returns>
        public Task<VbaTestResult> InvokeAsync(VbaTestCatalog catalog, VbaTestDescriptor procedure, string phase)
        {
            RequireOwner();
            if (active != null) throw new VbaTestInvocationException("A native test attempt is already pending.", false);
            if (catalog?.Project == null || procedure == null) throw new ArgumentException("A discovered test plan is required.");
            Validate(catalog);
            var call = new Call { Catalog = catalog, Test = procedure, Phase = phase };
            active = call;
            try { Post(() => Dispatch(call)); }
            catch (Exception error) { Finish(call, error); }
            return call.Completion.Task;
        }

        /// <summary>Prepares and revalidates the editor, arms result correlation, then dispatches exactly one native call.</summary>
        /// <param name="call">Outstanding attempt reserved by <see cref="InvokeAsync"/>.</param>
        private void Dispatch(Call call)
        {
            if (!ReferenceEquals(active, call) || call.Completion.Task.IsCompleted) return;
            try
            {
                // Native selection can pump editor callbacks; reserve it through dispatch and immediate observation.
                using (new VbeDebugInspection())
                {
                    Validate(call.Catalog);
                    call.Project = resolveProject(call.Catalog);
                    var support = call.Catalog.Project.Modules.Single(module => string.Equals(module.Name, VbaTestRuntimeSource.ModuleName, StringComparison.OrdinalIgnoreCase));
                    object prepared = Probe.Prepare(vbe, call.Project, support.Source);
                    // Showing a pane can pump messages. Revalidate all authority after selection.
                    Validate(call.Catalog);
                    if (Probe is NativeProbe nativeProbe)
                    {
                        if (!nativeProbe.SameIdentity(call.Project, resolveProject(call.Catalog)))
                            throw new InvalidOperationException("The selected project identity changed during native preparation.");
                        nativeProbe.RevalidateBeforeArming(vbe, call.Project, prepared);
                        // Focus activation can pump callbacks: verify all authority again before arming.
                        Validate(call.Catalog);
                    }
                    Probe.Revalidate(vbe, call.Project, prepared);
                    sink.Arm(call.Project, string.IsNullOrEmpty(call.Catalog.Project.HostPath) ? call.Catalog.Project.Id : call.Catalog.Project.HostPath,
                        VbaTestRuntimeSource.Version, call.Catalog.Project.Revision, runId(), call.Test, call.Phase, signature(call.Catalog));
                    call.Armed = true;
                    Probe.Revalidate(vbe, call.Project, prepared);
                    executionGuard();
                    sink.BeginNative();
                    call.Invoked = true;
                    call.Clock.Restart();
                    call.Exposure = VbaTestRuntime.Expose(sink);
                    // Never infer macro success from Execute returning: callback and Design mode are required.
                    Probe.Execute(prepared);
                    Observe(call);
                    if (!call.Completion.Task.IsCompleted) call.Polling = StartPolling(() => Observe(call));
                }
            }
            catch (Exception error) { Finish(call, error); }
        }

        /// <summary>Polls host mode and correlated callback state until both prove completion or a limit is reached.</summary>
        /// <param name="call">Attempt whose native state and result receipt are being observed.</param>
        private void Observe(Call call)
        {
            if (!ReferenceEquals(active, call) || call.Completion.Task.IsCompleted) return;
            try
            {
                RequireOwner();
                if (sink.HasFault) throw new InvalidOperationException("The runtime callback rejected this attempt.");
                Probe.RequireOwner(vbe);
                int mode = Probe.ReadMode(call.Project);
                if (mode == 1) throw new InvalidOperationException("The test entered break mode; its completion is uncertain.");
                if (mode != 0 && mode != 2) throw new InvalidOperationException("The test project mode is unavailable.");
                if (mode == 2 && sink.HasVerdict)
                {
                    object verdict = sink.CompleteNative(true);
                    call.Armed = false;
                    var result = VbaTestRuntimeSource.Decode(call.Test, verdict);
                    result.Phase = call.Phase;
                    result.Duration = call.Clock.Elapsed;
                    Revoke(call);
                    active = null;
                    call.Completion.TrySetResult(result);
                }
                else if (call.Clock.Elapsed >= VerificationTimeout)
                    throw new InvalidOperationException("Native completion and callback verdict were not both verified before the observation deadline.");
            }
            catch (Exception error) { Finish(call, error); }
        }

        /// <summary>Settles an attempt once, releases temporary exposure, and marks dispatched failures uncertain.</summary>
        /// <param name="call">Attempt to settle.</param>
        /// <param name="error">Preflight or completion failure.</param>
        private void Finish(Call call, Exception error)
        {
            if (call.Completion.Task.IsCompleted) return;
            if (call.Armed)
            {
                try { if (call.Invoked) sink.CompleteNative(false, error.Message); else sink.CancelUndispatched(); }
                catch (Exception callbackError) { if (call.Invoked) error = callbackError; }
                call.Armed = false;
            }
            if (call.Invoked) uncertain = true;
            Revoke(call);
            if (ReferenceEquals(active, call)) active = null;
            call.Completion.TrySetException(new VbaTestInvocationException(error.Message, call.Invoked, error));
        }

        /// <summary>Stops polling and removes the temporary result exposure associated with an attempt.</summary>
        /// <param name="call">Attempt whose temporary resources are released.</param>
        private static void Revoke(Call call)
        { call.Polling?.Dispose(); call.Polling = null; call.Exposure?.Dispose(); call.Exposure = null; }

        /// <summary>Throws unless the caller is on the thread that captured this VBE automation object.</summary>
        private void RequireOwner()
        { if (owner != Thread.CurrentThread.ManagedThreadId) throw new InvalidOperationException("Native test transport requires its owning thread."); }

        /// <summary>Disconnects the transport and settles an outstanding attempt as uncertain if it was dispatched.</summary>
        public void Dispose()
        {
            RequireOwner(); disposed = true;
            if (active != null) Finish(active, new InvalidOperationException("The live native test transport disconnected."));
        }

        /// <summary>Validates native run control for vba native test execution host.</summary>
        /// <param name="controlObject">Native command control inspected before dispatch.</param>
        /// <param name="expectedCaption">Optional caption captured earlier; a change causes refusal.</param>
        /// <exception cref="InvalidOperationException">The control is not the enabled built-in Run Sub command or its identity changed.</exception>
        internal static void ValidateNativeRunControl(object controlObject, string expectedCaption = null)
        {
            dynamic control = controlObject;
            if (control == null || (int)control.Id != 186 || (int)control.Type != 1 || !(bool)control.BuiltIn
                || !(bool)control.Enabled || !string.IsNullOrEmpty((string)control.OnAction)
                || !VbeDebug.IsAllowed("run", (string)control.Caption, 2)
                || (expectedCaption != null && (string)control.Caption != expectedCaption))
                throw new InvalidOperationException("The verified built-in native Run Sub command (186) is unavailable or changed.");
        }

        /// <summary>Selects and displays the requested module line, then captures pane and window identity for revalidation.</summary>
        /// <param name="editorObject">VBE editor window manager.</param>
        /// <param name="moduleObject">Module containing the intended procedure.</param>
        /// <param name="line">One-based declaration line to select.</param>
        /// <param name="observeAfterShow">Optional callback to verify the editor after showing its pane.</param>
        /// <param name="identity">COM identity comparer; defaults to the shared VBE identity check.</param>
        /// <returns>Opaque pane preparation record used to detect a changed selection before dispatch.</returns>
        internal static object PrepareNativePane(object editorObject, object moduleObject, int line, Action<object> observeAfterShow = null,
            Func<object, object, bool> identity = null)
        {
            if (identity == null) identity = VbeDebug.NativeProcedureValuesHost.SameComIdentity;
            dynamic editor = editorObject;
            // Show() selects a pane within its window; it does not make the VBE's
            // parent window visible. Materialize that parent before acquiring a pane.
            editor.MainWindow.Visible = true;
            if (!(bool)editor.MainWindow.Visible)
                throw new InvalidOperationException("The native VBE main window could not be made visible before pane selection.");
            editor.MainWindow.SetFocus();
            dynamic pane = ((dynamic)moduleObject).CodePane;
            pane.Show();
            observeAfterShow?.Invoke((object)pane);
            editor.ActiveCodePane = pane;
            pane.SetSelection(line, 1, line, 1);
            dynamic window = pane.Window;
            // VBIDE can report HWnd=0 for a live code window. Its COM identity,
            // module, type and visibility determine whether it is eligible for focus.
            if (!identity(moduleObject, (object)pane.CodeModule) || (int)window.Type != 0 || !(bool)window.Visible)
                throw new InvalidOperationException("The exact support code window is not visible before native focus.");
            window.SetFocus();
            return (object)pane;
        }

        /// <summary>Validates native pane context for vba native test execution host.</summary>
        /// <param name="editorObject">object that supplies the editor object for this operation.</param>
        /// <param name="projectObject">object that supplies the project object for this operation.</param>
        /// <param name="moduleObject">object that supplies the module object for this operation.</param>
        /// <param name="expectedPane">object that supplies the expected pane for this operation.</param>
        /// <param name="expectedLine">int that supplies the expected line for this operation.</param>
        /// <param name="identity">func&lt;object, object, bool&gt; that supplies the identity for this operation.</param>
        internal static void ValidateNativePaneContext(object editorObject, object projectObject, object moduleObject,
            object expectedPane, int expectedLine, Func<object, object, bool> identity = null)
        {
            if (identity == null) identity = VbeDebug.NativeProcedureValuesHost.SameComIdentity;
            dynamic editor = editorObject;
            object activeProject = (object)editor.ActiveVBProject;
            if (activeProject == null || !identity(projectObject, activeProject))
                throw new InvalidOperationException("The exact selected project is not active before native dispatch.");
            dynamic window = ((dynamic)expectedPane).Window;
            object activeWindow = (object)editor.ActiveWindow;
            bool available = activeWindow != null, matches = available && identity((object)window, activeWindow);
            int type = (int)window.Type;
            bool visible = (bool)window.Visible;
            if (!available || !matches || type != 0 || !visible)
                throw new InvalidOperationException("The exact support code window does not own native focus before dispatch. Observed focus: activeWindowAvailable="
                    + available + "; activeWindowMatches=" + matches + "; supportWindowType=" + type + "; supportWindowVisible=" + visible + ".");
            ValidateNativePaneSelection(editorObject, moduleObject, expectedPane, expectedLine, identity);
        }

        /// <summary>Validates native pane selection for vba native test execution host.</summary>
        /// <param name="editorObject">object that supplies the editor object for this operation.</param>
        /// <param name="moduleObject">object that supplies the module object for this operation.</param>
        /// <param name="expectedPane">object that supplies the expected pane for this operation.</param>
        /// <param name="expectedLine">int that supplies the expected line for this operation.</param>
        /// <param name="identity">func&lt;object, object, bool&gt; that supplies the identity for this operation.</param>
        private static void ValidateNativePaneSelection(object editorObject, object moduleObject,
            object expectedPane, int expectedLine, Func<object, object, bool> identity)
        {
            dynamic editor = editorObject;
            dynamic pane = editor.ActiveCodePane;
            if (pane == null || !identity(expectedPane, (object)pane) || !identity(moduleObject, (object)pane.CodeModule))
                throw new InvalidOperationException("The support code pane no longer owns the native run selection.");
            int line = 0, column = 0, endLine = 0, endColumn = 0;
            pane.GetSelection(ref line, ref column, ref endLine, ref endColumn);
            if (line != expectedLine || endLine != expectedLine || column != 1 || endColumn != 1
                || (int)((dynamic)moduleObject).ProcBodyLine[VbaTestRuntimeSource.PendingProcedure, 0] != expectedLine)
                throw new InvalidOperationException("The exact pending wrapper body no longer owns the native selection.");
        }

        /// <summary>Ensures native pane focus for vba native test execution host.</summary>
        /// <param name="editorObject">object that supplies the editor object for this operation.</param>
        /// <param name="projectObject">object that supplies the project object for this operation.</param>
        /// <param name="moduleObject">object that supplies the module object for this operation.</param>
        /// <param name="paneObject">object that supplies the pane object for this operation.</param>
        /// <param name="line">int that supplies the line for this operation.</param>
        /// <param name="windows">i windows that supplies the windows for this operation.</param>
        /// <param name="identity">func&lt;object, object, bool&gt; that supplies the identity for this operation.</param>
        /// <returns>target produced by the operation for ensure native pane focus on vba native test execution host.</returns>
        internal static VbaNativeTestWindowFocus.Target EnsureNativePaneFocus(object editorObject, object projectObject,
            object moduleObject, object paneObject, int line, VbaNativeTestWindowFocus.IWindows windows,
            Func<object, object, bool> identity = null)
        {
            if (identity == null) identity = VbeDebug.NativeProcedureValuesHost.SameComIdentity;
            dynamic editor = editorObject, window = ((dynamic)paneObject).Window;
            ValidateNativeFocusTarget(editorObject, projectObject, moduleObject, paneObject, line, identity);
            try
            {
                ValidateNativePaneContext(editorObject, projectObject, moduleObject, paneObject, line, identity);
                return null;
            }
            catch (InvalidOperationException)
            {
                ValidateNativeFocusTarget(editorObject, projectObject, moduleObject, paneObject, line, identity);
                // No native command was dispatched. The verified native child is
                // activated once, then COM must independently confirm the mapping.
                var target = VbaNativeTestWindowFocus.Focus(windows,
                    new IntPtr(Convert.ToInt64(editor.MainWindow.HWnd)), (string)window.Caption);
                ValidateNativePaneContext(editorObject, projectObject, moduleObject, paneObject, line, identity);
                ValidateRecoveredNativeFocus(editorObject, paneObject, windows, target);
                return target;
            }
        }

        /// <summary>Validates native focus target for vba native test execution host.</summary>
        /// <param name="editorObject">object that supplies the editor object for this operation.</param>
        /// <param name="projectObject">object that supplies the project object for this operation.</param>
        /// <param name="moduleObject">object that supplies the module object for this operation.</param>
        /// <param name="paneObject">object that supplies the pane object for this operation.</param>
        /// <param name="line">int that supplies the line for this operation.</param>
        /// <param name="identity">func&lt;object, object, bool&gt; that supplies the identity for this operation.</param>
        private static void ValidateNativeFocusTarget(object editorObject, object projectObject, object moduleObject,
            object paneObject, int line, Func<object, object, bool> identity)
        {
            dynamic editor = editorObject, pane = paneObject, window = pane.Window;
            // This fallback follows the completed COM focus attempt. It must not
            // repair a changed project, module, code window or wrapper selection.
            if (!identity(projectObject, (object)editor.ActiveVBProject) || !identity(moduleObject, (object)pane.CodeModule)
                || (int)window.Type != 0 || !(bool)window.Visible)
                throw new InvalidOperationException("The exact support context changed before native focus recovery.");
            int start = 0, column = 0, end = 0, endColumn = 0;
            pane.GetSelection(ref start, ref column, ref end, ref endColumn);
            if (start != line || end != line || column != 1 || endColumn != 1
                || (int)((dynamic)moduleObject).ProcBodyLine[VbaTestRuntimeSource.PendingProcedure, 0] != line)
                throw new InvalidOperationException("The pending wrapper selection changed before native focus recovery.");
        }

        /// <summary>Validates recovered native focus for vba native test execution host.</summary>
        /// <param name="editorObject">object that supplies the editor object for this operation.</param>
        /// <param name="paneObject">object that supplies the pane object for this operation.</param>
        /// <param name="windows">i windows that supplies the windows for this operation.</param>
        /// <param name="target">target that supplies the target for this operation.</param>
        internal static void ValidateRecoveredNativeFocus(object editorObject, object paneObject,
            VbaNativeTestWindowFocus.IWindows windows, VbaNativeTestWindowFocus.Target target)
        {
            if (target == null) return;
            if (new IntPtr(Convert.ToInt64(((dynamic)editorObject).MainWindow.HWnd)) != target.Main.Handle
                || !string.Equals((string)((dynamic)paneObject).Window.Caption, target.Caption, StringComparison.Ordinal))
                throw new InvalidOperationException("The COM-to-native code-window mapping changed after focus recovery.");
            VbaNativeTestWindowFocus.VerifyFocus(windows, target);
        }

        /// <summary>Handles describe native windows for vba native test execution host.</summary>
        /// <param name="editorObject">object that supplies the editor object for this operation.</param>
        /// <param name="expectedPane">object that supplies the expected pane for this operation.</param>
        /// <param name="expectedModule">object that supplies the expected module for this operation.</param>
        /// <param name="identity">func&lt;object, object, bool&gt; that supplies the identity for this operation.</param>
        /// <param name="nativeWindows">func&lt;int ptr, string&gt; that supplies the native windows for this operation.</param>
        /// <returns>Text produced by the operation for describe native windows on vba native test execution host.</returns>
        internal static string DescribeNativeWindows(object editorObject, object expectedPane, object expectedModule,
            Func<object, object, bool> identity = null, Func<IntPtr, string> nativeWindows = null)
        {
            if (identity == null) identity = VbeDebug.NativeProcedureValuesHost.SameComIdentity;
            object main = null, expectedWindow = null, activeWindow = null, activePane = null, activeModule = null;
            DiagnosticRead(() => { main = (object)((dynamic)editorObject).MainWindow; return "available"; });
            DiagnosticRead(() => { expectedWindow = (object)((dynamic)expectedPane).Window; return "available"; });
            DiagnosticRead(() => { activeWindow = (object)((dynamic)editorObject).ActiveWindow; return "available"; });
            DiagnosticRead(() => { activePane = (object)((dynamic)editorObject).ActiveCodePane; return "available"; });
            DiagnosticRead(() => { activeModule = (object)((dynamic)activePane).CodeModule; return "available"; });
            var details = new StringBuilder("main={" + DiagnosticWindow(main) + "}; paneWindow={" + DiagnosticWindow(expectedWindow)
                + "}; activeWindow={" + DiagnosticWindow(activeWindow) + "}; activeWindowMatches=" + DiagnosticMatches(identity, expectedWindow, activeWindow)
                + "; activePaneAvailable=" + (activePane != null) + "; activePaneMatches=" + DiagnosticMatches(identity, expectedPane, activePane)
                + "; activeModuleMatches=" + DiagnosticMatches(identity, expectedModule, activeModule) + "; selection=");
            details.Append(DiagnosticRead(() =>
            {
                int start = 0, column = 0, end = 0, endColumn = 0;
                ((dynamic)activePane).GetSelection(ref start, ref column, ref end, ref endColumn);
                return start + ":" + column + "-" + end + ":" + endColumn;
            }));
            details.Append("; COMWindows=[");
            try
            {
                int count = 0;
                foreach (object window in ((dynamic)editorObject).Windows)
                {
                    if (count++ >= 12) { details.Append("bounded"); break; }
                    details.Append("{" + DiagnosticWindow(window) + "}");
                }
            }
            catch (Exception error) { details.Append("unavailable(" + error.GetType().Name + ")"); }
            details.Append("]; ownedNative=[");
            try
            {
                var handle = new IntPtr(Convert.ToInt64(((dynamic)main).HWnd));
                if (nativeWindows == null) nativeWindows = NativeWindowObservation.Read;
                string native = nativeWindows(handle);
                details.Append(native == null ? "null" : native.Substring(0, Math.Min(native.Length, 4000)));
            }
            catch (Exception error) { details.Append("unavailable(" + error.GetType().Name + ")"); }
            return details.Append("]").ToString();
        }

        /// <summary>Handles diagnostic window for vba native test execution host.</summary>
        /// <param name="window">object that supplies the window for this operation.</param>
        /// <returns>Text produced by the operation for diagnostic window on vba native test execution host.</returns>
        private static string DiagnosticWindow(object window)
        {
            return "Type=" + DiagnosticRead(() => ((dynamic)window).Type)
                + ",Visible=" + DiagnosticRead(() => ((dynamic)window).Visible)
                + ",HWnd=" + DiagnosticRead(() => ((dynamic)window).HWnd)
                + ",State=" + DiagnosticRead(() => ((dynamic)window).WindowState);
        }

        /// <summary>Owns the native window observation state and operations.</summary>
        internal static class NativeWindowObservation
        {

            /// <summary>Defines the enum window callback.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <param name="unused">Native handle that supplies the unused for this operation.</param>
            /// <returns>Boolean indicating the result of the check for operation on native window observation.</returns>
            private delegate bool EnumWindow(IntPtr window, IntPtr unused);

            /// <summary>Handles enum child windows for native window observation.</summary>
            /// <param name="parent">Native handle that supplies the parent for this operation.</param>
            /// <param name="callback">enum window that supplies the callback for this operation.</param>
            /// <param name="unused">Native handle that supplies the unused for this operation.</param>
            /// <returns>Boolean indicating the result of the check for enum child windows on native window observation.</returns>
            [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, EnumWindow callback, IntPtr unused);

            /// <summary>Reads  for native window observation.</summary>
            /// <param name="main">Native handle that supplies the main for this operation.</param>
            /// <returns>Text produced by the operation for read on native window observation.</returns>
            internal static string Read(IntPtr main)
            { return Read(main, new VbaNativeTestWindowFocus.NativeWindows(), Enumerate); }

            /// <summary>Handles enumerate for native window observation.</summary>
            /// <param name="main">Native handle that supplies the main for this operation.</param>
            /// <param name="visit">func&lt;int ptr, bool&gt; that supplies the visit for this operation.</param>
            private static void Enumerate(IntPtr main, Func<IntPtr, bool> visit)
            { EnumChildWindows(main, (window, unused) => visit(window), IntPtr.Zero); }

            /// <summary>Reads  for native window observation.</summary>
            /// <param name="main">Native handle that supplies the main for this operation.</param>
            /// <param name="windows">i windows that supplies the windows for this operation.</param>
            /// <param name="enumerate">action&lt;int ptr, func&lt;int ptr, bool&gt;&gt; that supplies the enumerate for this operation.</param>
            /// <returns>Text produced by the operation for read on native window observation.</returns>
            internal static string Read(IntPtr main, VbaNativeTestWindowFocus.IWindows windows, Action<IntPtr, Func<IntPtr, bool>> enumerate)
            {
                var root = windows.Read(main);
                uint owner = root.Process;
                if (main == IntPtr.Zero || !root.Exists || owner != windows.CurrentProcess || root.Thread != windows.CurrentThread)
                    return "unavailable(owner mismatch)";
                var result = new StringBuilder();
                Append(result, root, owner);
                int count = 0;
                enumerate(main, window =>
                {
                    if (count++ >= 64) { result.Append("bounded"); return false; }
                    var child = windows.Read(window);
                    if (child.Process == owner) Append(result, child, child.Process);
                    return result.Length < 3900;
                });
                return result.ToString();
            }

            /// <summary>Handles append for native window observation.</summary>
            /// <param name="result">string builder that supplies the result for this operation.</param>
            /// <param name="window">window that supplies the window for this operation.</param>
            /// <param name="owner">uint that supplies the owner for this operation.</param>
            private static void Append(StringBuilder result, VbaNativeTestWindowFocus.Window window, uint owner)
            {
                string kind = window.Class ?? "";
                if (kind.Length > 95) kind = kind.Substring(0, 95);
                result.Append("{HWnd=" + window.Handle.ToInt64() + ",PID=" + owner + ",Class=" + kind
                    + ",Parent=" + window.Parent.ToInt64() + ",Visible=" + window.Visible + "}");
            }
        }

        // Read only after a refusal. Failed inspection must preserve the original refusal,
        // and must never reselect a pane, focus a window or dispatch a native command.
        /// <summary>Handles describe native selection for vba native test execution host.</summary>
        /// <param name="editorObject">object that supplies the editor object for this operation.</param>
        /// <param name="expectedProject">object that supplies the expected project for this operation.</param>
        /// <param name="expectedModule">object that supplies the expected module for this operation.</param>
        /// <param name="expectedPane">object that supplies the expected pane for this operation.</param>
        /// <param name="expectedLine">int that supplies the expected line for this operation.</param>
        /// <param name="identity">func&lt;object, object, bool&gt; that supplies the identity for this operation.</param>
        /// <returns>Text produced by the operation for describe native selection on vba native test execution host.</returns>
        internal static string DescribeNativeSelection(object editorObject, object expectedProject, object expectedModule,
            object expectedPane, int expectedLine, Func<object, object, bool> identity = null)
        {
            if (identity == null) identity = VbeDebug.NativeProcedureValuesHost.SameComIdentity;
            object activeProject = null, pane = null, module = null, paneProject = null;
            string activeProjectRead = DiagnosticRead(() => { activeProject = (object)((dynamic)editorObject).ActiveVBProject; return activeProject == null ? "null" : "available"; });
            string paneRead = DiagnosticRead(() => { pane = (object)((dynamic)editorObject).ActiveCodePane; return pane == null ? "null" : "available"; });
            string moduleRead = DiagnosticRead(() => { module = (object)((dynamic)pane).CodeModule; return module == null ? "null" : "available"; });
            string projectRead = DiagnosticRead(() => { paneProject = (object)((dynamic)module).Parent.Collection.Parent; return paneProject == null ? "null" : "available"; });
            bool activeProjectMatches = DiagnosticMatches(identity, expectedProject, activeProject);
            bool moduleMatches = DiagnosticMatches(identity, expectedModule, module);
            bool paneProjectMatches = DiagnosticMatches(identity, expectedProject, paneProject);
            bool inspectedModuleIsAuthorized = moduleMatches || paneProjectMatches;
            const string redacted = "redacted(outside selected project)";
            int startLine = 0, startColumn = 0, endLine = 0, endColumn = 0;
            string selection = DiagnosticRead(() =>
            {
                ((dynamic)pane).GetSelection(ref startLine, ref startColumn, ref endLine, ref endColumn);
                return startLine + ":" + startColumn + "-" + endLine + ":" + endColumn;
            });
            return " Expected: project=" + DiagnosticRead(() => ((dynamic)expectedProject).Name)
                + "; module=" + DiagnosticRead(() => ((dynamic)expectedModule).Parent.Name)
                + "; procedure=" + VbaTestRuntimeSource.PendingProcedure + "; runtimeVersion=" + VbaTestRuntimeSource.Version
                + "; mainWindowVisible=" + DiagnosticRead(() => ((dynamic)editorObject).MainWindow.Visible)
                + "; selection=" + expectedLine + ":1-" + expectedLine + ":1"
                + "; currentBodyLine=" + DiagnosticRead(() => ((dynamic)expectedModule).ProcBodyLine[VbaTestRuntimeSource.PendingProcedure, 0])
                + "; paneWindow=" + DiagnosticRead(() => ((dynamic)expectedPane).Window.HWnd)
                + ". Actual: activeProject=" + (activeProjectMatches ? DiagnosticRead(() => ((dynamic)activeProject).Name) : redacted)
                + "; activeProjectMatches=" + activeProjectMatches
                + "; paneMatches=" + DiagnosticRead(() => identity(expectedPane, pane))
                + "; paneWindow=" + DiagnosticRead(() => ((dynamic)pane).Window.HWnd)
                + "; module=" + (inspectedModuleIsAuthorized ? DiagnosticRead(() => ((dynamic)module).Parent.Name) : redacted)
                + "; moduleMatches=" + moduleMatches
                + "; paneProject=" + (paneProjectMatches ? DiagnosticRead(() => ((dynamic)paneProject).Name) : redacted)
                + "; paneProjectMatches=" + paneProjectMatches
                + "; selection=" + selection
                + "; startProcedure=" + (inspectedModuleIsAuthorized ? DiagnosticRead(() => DiagnosticProcedure(module, startLine)) : redacted)
                + "; endProcedure=" + (inspectedModuleIsAuthorized ? DiagnosticRead(() => DiagnosticProcedure(module, endLine)) : redacted)
                + "; reads=" + activeProjectRead + "/" + paneRead + "/" + moduleRead + "/" + projectRead + ".";
        }

        /// <summary>Handles diagnostic matches for vba native test execution host.</summary>
        /// <param name="identity">func&lt;object, object, bool&gt; that supplies the identity for this operation.</param>
        /// <param name="expected">object that supplies the expected for this operation.</param>
        /// <param name="actual">object that supplies the actual for this operation.</param>
        /// <returns>Boolean indicating the result of the check for diagnostic matches on vba native test execution host.</returns>
        private static bool DiagnosticMatches(Func<object, object, bool> identity, object expected, object actual)
        {
            if (expected == null || actual == null) return false;
            try { return identity(expected, actual); }
            catch { return false; }
        }

        /// <summary>Handles diagnostic procedure for vba native test execution host.</summary>
        /// <param name="module">object that supplies the module for this operation.</param>
        /// <param name="line">int that supplies the line for this operation.</param>
        /// <returns>Text produced by the operation for diagnostic procedure on vba native test execution host.</returns>
        private static string DiagnosticProcedure(object module, int line)
        {
            int kind = 0;
            string name = (string)((dynamic)module).ProcOfLine[line, ref kind];
            return name + " (kind=" + kind + ")";
        }

        /// <summary>Handles diagnostic read for vba native test execution host.</summary>
        /// <param name="read">func&lt;object&gt; that supplies the read for this operation.</param>
        /// <returns>Text produced by the operation for diagnostic read on vba native test execution host.</returns>
        private static string DiagnosticRead(Func<object> read)
        {
            try
            {
                object value = read();
                if (value == null) return "null";
                if (Marshal.IsComObject(value)) return "available";
                string text = Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture);
                text = text.Replace('\r', ' ').Replace('\n', ' ');
                return text.Length <= 256 ? text : text.Substring(0, 256);
            }
            catch (Exception error) { return "unavailable(" + error.GetType().Name + ")"; }
        }

        /// <summary>Owns the native probe state and operations.</summary>
        internal sealed class NativeProbe : IProbe
        {

            /// <summary>Owns the prepared state and operations.</summary>
            private sealed class Prepared
            {

                /// <summary>Maintains the editor and project and module and pane and control state for prepared.</summary>
                internal object Editor, Project, Module, Pane, Control;

                /// <summary>Maintains the source and caption and after show and after focus state for prepared.</summary>
                internal string Source, Caption, AfterShow, AfterFocus;

                /// <summary>Maintains the line state for prepared.</summary>
                internal int Line;

                /// <summary>Maintains the focus recovery attempted state for prepared.</summary>
                internal bool FocusRecoveryAttempted;

                /// <summary>Maintains the native focus state for prepared.</summary>
                internal VbaNativeTestWindowFocus.Target NativeFocus;
            }

            /// <summary>Maintains the windows state for native probe.</summary>
            private readonly VbaNativeTestWindowFocus.IWindows windows;

            /// <summary>Maintains the identity state for native probe.</summary>
            private readonly Func<object, object, bool> identity;

            /// <summary>Initializes a NativeProbe instance with the supplied state.</summary>
            /// <param name="windows">i windows that supplies the windows for this operation.</param>
            /// <param name="identity">func&lt;object, object, bool&gt; that supplies the identity for this operation.</param>
            internal NativeProbe(VbaNativeTestWindowFocus.IWindows windows = null, Func<object, object, bool> identity = null)
            {
                this.windows = windows ?? new VbaNativeTestWindowFocus.NativeWindows();
                this.identity = identity ?? VbeDebug.NativeProcedureValuesHost.SameComIdentity;
            }

            /// <summary>Compares identity for native probe.</summary>
            /// <param name="first">object that supplies the first for this operation.</param>
            /// <param name="second">object that supplies the second for this operation.</param>
            /// <returns>Boolean indicating the result of the check for same identity on native probe.</returns>
            internal bool SameIdentity(object first, object second) => identity(first, second);

            /// <summary>Requires owner for native probe.</summary>
            /// <param name="editor">object that supplies the editor for this operation.</param>
            public void RequireOwner(object editor)
            {
                if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA) throw new InvalidOperationException("Native VBE execution requires an STA.");
                var window = new IntPtr(Convert.ToInt64(((dynamic)editor).MainWindow.HWnd));
                var observed = windows.Read(window);
                if (window == IntPtr.Zero || observed.Process != windows.CurrentProcess || observed.Thread != windows.CurrentThread)
                    throw new InvalidOperationException("The VBE window does not belong to this process and owning UI thread.");
            }

            /// <summary>Handles prepare for native probe.</summary>
            /// <param name="editorObject">object that supplies the editor object for this operation.</param>
            /// <param name="projectObject">object that supplies the project object for this operation.</param>
            /// <param name="source">Text that supplies the source value. Use the format required by the calling operation.</param>
            /// <returns>object produced by the operation for prepare on native probe.</returns>
            public object Prepare(object editorObject, object projectObject, string source)
            {
                RequireOwner(editorObject);
                dynamic editor = editorObject, project = projectObject;
                object component = null;
                foreach (dynamic candidate in project.VBComponents)
                    if (string.Equals((string)candidate.Name, VbaTestRuntimeSource.ModuleName, StringComparison.OrdinalIgnoreCase))
                    { if (component != null) throw new InvalidOperationException("The support module is ambiguous."); component = candidate; }
                if (component == null || (int)((dynamic)component).Type != 1) throw new InvalidOperationException("The standard support module is unavailable.");
                dynamic module = ((dynamic)component).CodeModule;
                if (ReadSource(module) != source) throw new InvalidOperationException("The support source changed before selection.");
                int line = (int)module.ProcBodyLine[VbaTestRuntimeSource.PendingProcedure, 0];
                if (!Regex.IsMatch((string)module.Lines[line, 1], @"^\s*Public\s+Sub\s+" + Regex.Escape(VbaTestRuntimeSource.PendingProcedure)
                    + @"\s*\(\s*\)\s*(?:'.*)?$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                    throw new InvalidOperationException("An explicit public parameterless " + VbaTestRuntimeSource.PendingProcedure + " wrapper is required.");
                string afterShow = null;
                object pane;
                try { pane = PrepareNativePane(editorObject, (object)module, line, shown => afterShow = DescribeNativeWindows(editorObject, shown, (object)module), identity); }
                catch (Exception error)
                {
                    if (afterShow != null) throw new InvalidOperationException(error.Message + " Immediately after Show: " + afterShow, error);
                    throw;
                }
                RequireOwner(editorObject);
                string afterFocus = DescribeNativeWindows(editorObject, pane, (object)module);
                VbaNativeTestWindowFocus.Target nativeFocus;
                try { nativeFocus = EnsureNativePaneFocus(editorObject, projectObject, (object)module, pane, line, windows, identity); }
                catch (Exception error)
                {
                    throw new InvalidOperationException(error.Message
                        + DescribeNativeSelection(editorObject, projectObject, (object)module, pane, line)
                        + " Immediately after COM Focus: " + afterFocus
                        + " After native focus recovery: " + DescribeNativeWindows(editorObject, pane, (object)module), error);
                }
                RequireOwner(editorObject);
                dynamic control = editor.CommandBars.FindControl(1, 186);
                ValidateNativeRunControl((object)control);
                return new Prepared { Editor = editorObject, Project = projectObject, Module = module, Pane = pane, Control = control, Source = source, Caption = (string)control.Caption, Line = line, AfterShow = afterShow, AfterFocus = afterFocus, NativeFocus = nativeFocus, FocusRecoveryAttempted = nativeFocus != null };
            }

            /// <summary>Handles revalidate for native probe.</summary>
            /// <param name="editorObject">object that supplies the editor object for this operation.</param>
            /// <param name="projectObject">object that supplies the project object for this operation.</param>
            /// <param name="prepared">object that supplies the prepared for this operation.</param>
            public void Revalidate(object editorObject, object projectObject, object prepared)
            { Revalidate(editorObject, projectObject, prepared, false); }

            /// <summary>Handles revalidate before arming for native probe.</summary>
            /// <param name="editorObject">object that supplies the editor object for this operation.</param>
            /// <param name="projectObject">object that supplies the project object for this operation.</param>
            /// <param name="prepared">object that supplies the prepared for this operation.</param>
            internal void RevalidateBeforeArming(object editorObject, object projectObject, object prepared)
            { Revalidate(editorObject, projectObject, prepared, true); }

            /// <summary>Handles revalidate for native probe.</summary>
            /// <param name="editorObject">object that supplies the editor object for this operation.</param>
            /// <param name="projectObject">object that supplies the project object for this operation.</param>
            /// <param name="prepared">object that supplies the prepared for this operation.</param>
            /// <param name="allowFocusRecovery">Indicates whether allow focus recovery is enabled.</param>
            private void Revalidate(object editorObject, object projectObject, object prepared, bool allowFocusRecovery)
            {
                RequireOwner(editorObject);
                var plan = (Prepared)prepared;
                dynamic editor = editorObject, project = projectObject, control = plan.Control;
                bool found = false;
                foreach (object candidate in editor.VBProjects)
                    if (identity(candidate, projectObject)) { found = true; break; }
                if (!found || (int)project.Mode != 2 || ReadSource(plan.Module) != plan.Source)
                    throw new InvalidOperationException("The selected project, mode or support source changed before native dispatch.");
                string initialFocusRefusal = null;
                try
                {
                    try { ValidateNativePaneContext(editorObject, projectObject, plan.Module, plan.Pane, plan.Line, identity); }
                    catch (InvalidOperationException focusError)
                    {
                        initialFocusRefusal = focusError.Message;
                        if (!allowFocusRecovery || plan.FocusRecoveryAttempted) throw;
                        object handle = ((dynamic)plan.Pane).Window.HWnd;
                        if (handle == null || Convert.ToInt64(handle) != 0) throw;
                        // Only lost window focus may be repaired before arming. Changed selection or authority must refuse.
                        ValidateNativeFocusTarget(editorObject, projectObject, plan.Module, plan.Pane, plan.Line, identity);
                        ValidateNativePaneSelection(editorObject, plan.Module, plan.Pane, plan.Line, identity);
                        if ((int)project.Mode != 2 || ReadSource(plan.Module) != plan.Source)
                            throw new InvalidOperationException("The selected mode or support source changed before native focus recovery.");
                        plan.FocusRecoveryAttempted = true;
                        plan.NativeFocus = EnsureNativePaneFocus(editorObject, projectObject, plan.Module, plan.Pane, plan.Line, windows, identity);
                        if ((int)project.Mode != 2 || ReadSource(plan.Module) != plan.Source)
                            throw new InvalidOperationException("The selected mode or support source changed during native focus recovery.");
                    }
                    ValidateRecoveredNativeFocus(editorObject, plan.Pane, windows, plan.NativeFocus);
                }
                catch (Exception error)
                {
                    throw new InvalidOperationException(error.Message
                        + (initialFocusRefusal == null ? "" : " Initial focus refusal: " + initialFocusRefusal)
                        + DescribeNativeSelection(editorObject, projectObject, plan.Module, plan.Pane, plan.Line)
                        + " Immediately after Show: " + plan.AfterShow + " Immediately after Focus: " + plan.AfterFocus, error);
                }
                ValidateNativeRunControl((object)control, plan.Caption);
            }

            /// <summary>Executes  for native probe.</summary>
            /// <param name="prepared">object that supplies the prepared for this operation.</param>
            public void Execute(object prepared)
            {
                var plan = (Prepared)prepared;
                Revalidate(plan.Editor, plan.Project, plan);
                ((dynamic)plan.Control).Execute();
            }

            /// <summary>Reads mode for native probe.</summary>
            /// <param name="project">object that supplies the project for this operation.</param>
            /// <returns>int produced by the operation for read mode on native probe.</returns>
            public int ReadMode(object project) => (int)((dynamic)project).Mode;

            /// <summary>Reads source for native probe.</summary>
            /// <param name="module">dynamic that supplies the module for this operation.</param>
            /// <returns>Text produced by the operation for read source on native probe.</returns>
            private static string ReadSource(dynamic module)
            {
                int count = (int)module.CountOfLines;
                if (count < 0 || count > 200000) throw new InvalidOperationException("The support source exceeds inspection limits.");
                return count == 0 ? string.Empty : (string)module.Lines[1, count];
            }
        }
    }
}
