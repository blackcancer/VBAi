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

        /// <summary>Creates a session-bound native dispatcher whose COM work and callbacks stay on the constructing thread.</summary>
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

        /// <summary>Validates the active project, code window, and pending-wrapper selection before native dispatch.</summary>
        /// <param name="editorObject">Owning VBE automation object.</param>
        /// <param name="projectObject">Project authorized for the current run.</param>
        /// <param name="moduleObject">Generated support module containing the pending wrapper.</param>
        /// <param name="expectedPane">Pane captured during preparation.</param>
        /// <param name="expectedLine">One-based pending-wrapper body line that must remain selected.</param>
        /// <param name="identity">COM identity comparer; defaults to the shared native identity helper.</param>
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
        /// <param name="editorObject">Owning VBE automation object.</param>
        /// <param name="moduleObject">Generated support module containing the pending wrapper.</param>
        /// <param name="expectedPane">Pane captured during preparation.</param>
        /// <param name="expectedLine">One-based pending-wrapper body line that must remain selected.</param>
        /// <param name="identity">COM identity comparer; defaults to the shared native identity helper.</param>
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
        /// <param name="editorObject">Owning VBE automation object.</param>
        /// <param name="projectObject">Project authorized for the current run.</param>
        /// <param name="moduleObject">Generated support module containing the pending wrapper.</param>
        /// <param name="paneObject">Prepared code pane whose window is the required target.</param>
        /// <param name="line">Expected pending-wrapper body line.</param>
        /// <param name="windows">Native HWND reader/activator restricted to the current process and UI thread.</param>
        /// <param name="identity">COM identity comparer; defaults to the shared native identity helper.</param>
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
        /// <param name="editorObject">Owning VBE automation object.</param>
        /// <param name="projectObject">Project authorized for the current run.</param>
        /// <param name="moduleObject">Generated support module containing the pending wrapper.</param>
        /// <param name="paneObject">Prepared code pane whose window is the required target.</param>
        /// <param name="line">Expected pending-wrapper body line.</param>
        /// <param name="identity">COM identity comparer; defaults to the shared native identity helper.</param>
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
        /// <param name="editorObject">Owning VBE automation object.</param>
        /// <param name="paneObject">Prepared code pane whose window is the required target.</param>
        /// <param name="windows">Native HWND reader/activator restricted to the current process and UI thread.</param>
        /// <param name="target">Previously verified native VBE main, MDI, and code-window handle chain.</param>
        internal static void ValidateRecoveredNativeFocus(object editorObject, object paneObject,
            VbaNativeTestWindowFocus.IWindows windows, VbaNativeTestWindowFocus.Target target)
        {
            if (target == null) return;
            if (new IntPtr(Convert.ToInt64(((dynamic)editorObject).MainWindow.HWnd)) != target.Main.Handle
                || !string.Equals((string)((dynamic)paneObject).Window.Caption, target.Caption, StringComparison.Ordinal))
                throw new InvalidOperationException("The COM-to-native code-window mapping changed after focus recovery.");
            VbaNativeTestWindowFocus.VerifyFocus(windows, target);
        }

        /// <summary>Captures bounded COM and native-window diagnostics without focusing or dispatching a command.</summary>
        /// <param name="editorObject">Owning VBE automation object.</param>
        /// <param name="expectedPane">Pane captured during preparation.</param>
        /// <param name="expectedModule">Module expected to own the prepared pane.</param>
        /// <param name="identity">COM identity comparer; defaults to the shared native identity helper.</param>
        /// <param name="nativeWindows">Optional native-window formatter; defaults to a bounded owner-thread scan.</param>
        /// <returns>One bounded diagnostic string; individual failed COM reads are represented as unavailable.</returns>
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

        /// <summary>Reads bounded scalar properties from one observed VBIDE window.</summary>
        /// <param name="window">Window object, possibly unavailable.</param>
        /// <returns>Type, visibility, HWND, and state values with failed reads marked unavailable.</returns>
        private static string DiagnosticWindow(object window)
        {
            return "Type=" + DiagnosticRead(() => ((dynamic)window).Type)
                + ",Visible=" + DiagnosticRead(() => ((dynamic)window).Visible)
                + ",HWnd=" + DiagnosticRead(() => ((dynamic)window).HWnd)
                + ",State=" + DiagnosticRead(() => ((dynamic)window).WindowState);
        }

        /// <summary>Produces a bounded, read-only snapshot of native descendants belonging to the VBE process.</summary>
        internal static class NativeWindowObservation
        {

            /// <summary>Callback signature used by user32 to visit one descendant HWND.</summary>
            /// <param name="window">Current descendant handle.</param>
            /// <param name="unused">Opaque LPARAM supplied to the enumeration call.</param>
            /// <returns>True to continue enumeration; false to stop.</returns>
            private delegate bool EnumWindow(IntPtr window, IntPtr unused);

            /// <summary>Imports the user32 descendant-window enumeration function.</summary>
            /// <param name="parent">Root HWND whose descendants are enumerated.</param>
            /// <param name="callback">Managed callback invoked for each descendant.</param>
            /// <param name="unused">Opaque LPARAM supplied to the enumeration call.</param>
            /// <returns>Win32 enumeration success.</returns>
            [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, EnumWindow callback, IntPtr unused);

            /// <summary>Reads the owned VBE main window and a bounded set of same-process descendants.</summary>
            /// <param name="main">VBE main-window HWND.</param>
            /// <returns>Scalar window descriptions, or an owner-mismatch marker when the root is not current-thread owned.</returns>
            internal static string Read(IntPtr main)
            { return Read(main, new VbaNativeTestWindowFocus.NativeWindows(), Enumerate); }

            /// <summary>Adapts EnumChildWindows to a callback that can stop after a bound.</summary>
            /// <param name="main">VBE main-window HWND.</param>
            /// <param name="visit">Receives each HWND and returns false to stop early.</param>
            private static void Enumerate(IntPtr main, Func<IntPtr, bool> visit)
            { EnumChildWindows(main, (window, unused) => visit(window), IntPtr.Zero); }

            /// <summary>Reads the owned VBE main window and a bounded set of same-process descendants.</summary>
            /// <param name="main">VBE main-window HWND.</param>
            /// <param name="windows">Native HWND reader/activator restricted to the current process and UI thread.</param>
            /// <param name="enumerate">Descendant enumerator that stops when the callback returns false.</param>
            /// <returns>Scalar window descriptions, or an owner-mismatch marker when the root is not current-thread owned.</returns>
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

            /// <summary>Appends one scalar HWND record while limiting the class name to 95 characters.</summary>
            /// <param name="result">Diagnostic buffer receiving the record.</param>
            /// <param name="window">Observed native window snapshot.</param>
            /// <param name="owner">Process ID recorded for this observation.</param>
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
        /// <summary>Captures bounded selection and project diagnostics after native preparation is refused.</summary>
        /// <param name="editorObject">Owning VBE automation object.</param>
        /// <param name="expectedProject">Project expected to own the active editor selection.</param>
        /// <param name="expectedModule">Module expected to own the prepared pane.</param>
        /// <param name="expectedPane">Pane captured during preparation.</param>
        /// <param name="expectedLine">One-based pending-wrapper body line that must remain selected.</param>
        /// <param name="identity">COM identity comparer; defaults to the shared native identity helper.</param>
        /// <returns>Scalar expected/observed selection evidence; reads outside the selected project are redacted.</returns>
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

        /// <summary>Compares optional diagnostic objects without allowing identity failures to escape.</summary>
        /// <param name="identity">COM identity comparer; defaults to the shared native identity helper.</param>
        /// <param name="expected">Expected COM object; null never matches.</param>
        /// <param name="actual">Observed COM object; null never matches.</param>
        /// <returns>True only when the configured identity comparer succeeds.</returns>
        private static bool DiagnosticMatches(Func<object, object, bool> identity, object expected, object actual)
        {
            if (expected == null || actual == null) return false;
            try { return identity(expected, actual); }
            catch { return false; }
        }

        /// <summary>Reads the VBA procedure name and kind containing a source line.</summary>
        /// <param name="module">CodeModule whose procedure table is queried.</param>
        /// <param name="line">Expected pending-wrapper body line.</param>
        /// <returns>Procedure name and VBIDE procedure-kind value.</returns>
        private static string DiagnosticProcedure(object module, int line)
        {
            int kind = 0;
            string name = (string)((dynamic)module).ProcOfLine[line, ref kind];
            return name + " (kind=" + kind + ")";
        }

        /// <summary>Runs one optional diagnostic getter and converts failures to bounded markers.</summary>
        /// <param name="read">COM or scalar property getter to invoke once.</param>
        /// <returns>Null, available, invariant scalar text capped at 256 characters, or an exception-type marker.</returns>
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

        /// <summary>Implements STA-only native pane selection, focus revalidation, and one built-in Run dispatch.</summary>
        internal sealed class NativeProbe : IProbe
        {

            /// <summary>Captured editor/project/module/pane/control state required to revalidate one native attempt.</summary>
            private sealed class Prepared
            {

                /// <summary>COM identities captured before native dispatch.</summary>
                internal object Editor, Project, Module, Pane, Control;

                /// <summary>Expected support source, Run control caption, and bounded observations after Show and Focus.</summary>
                internal string Source, Caption, AfterShow, AfterFocus;

                /// <summary>One-based body line of the generated pending wrapper.</summary>
                internal int Line;

                /// <summary>Prevents more than one native focus recovery attempt.</summary>
                internal bool FocusRecoveryAttempted;

                /// <summary>Verified HWND ownership/focus snapshot, or null when COM focus already matched.</summary>
                internal VbaNativeTestWindowFocus.Target NativeFocus;
            }

            /// <summary>Native HWND inspection and guarded activation implementation.</summary>
            private readonly VbaNativeTestWindowFocus.IWindows windows;

            /// <summary>COM identity comparer used to revalidate editor objects.</summary>
            private readonly Func<object, object, bool> identity;

            /// <summary>Creates a native probe with optional HWND and COM-identity adapters.</summary>
            /// <param name="windows">Native HWND reader/activator restricted to the current process and UI thread.</param>
            /// <param name="identity">COM identity comparer; defaults to the shared native identity helper.</param>
            internal NativeProbe(VbaNativeTestWindowFocus.IWindows windows = null, Func<object, object, bool> identity = null)
            {
                this.windows = windows ?? new VbaNativeTestWindowFocus.NativeWindows();
                this.identity = identity ?? VbeDebug.NativeProcedureValuesHost.SameComIdentity;
            }

            /// <summary>Compares two objects using the configured COM identity rule.</summary>
            /// <param name="first">First COM object.</param>
            /// <param name="second">Second COM object.</param>
            /// <returns>True when the configured comparer reports the same COM identity.</returns>
            internal bool SameIdentity(object first, object second) => identity(first, second);

            /// <summary>Requires an STA caller and a VBE main window owned by this process and thread.</summary>
            /// <param name="editor">VBE automation object whose main-window ownership is checked.</param>
            public void RequireOwner(object editor)
            {
                if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA) throw new InvalidOperationException("Native VBE execution requires an STA.");
                var window = new IntPtr(Convert.ToInt64(((dynamic)editor).MainWindow.HWnd));
                var observed = windows.Read(window);
                if (window == IntPtr.Zero || observed.Process != windows.CurrentProcess || observed.Thread != windows.CurrentThread)
                    throw new InvalidOperationException("The VBE window does not belong to this process and owning UI thread.");
            }

            /// <summary>Validates the generated wrapper, selects its exact code pane, and captures dispatch identities.</summary>
            /// <param name="editorObject">Owning VBE automation object.</param>
            /// <param name="projectObject">Project authorized for the current run.</param>
            /// <param name="source">Expected complete generated support-module source.</param>
            /// <returns>Prepared identity snapshot used by subsequent revalidation and execution.</returns>
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

            /// <summary>Rechecks project, design mode, support source, native pane selection, focus, and Run control identity.</summary>
            /// <param name="editorObject">Owning VBE automation object.</param>
            /// <param name="projectObject">Project authorized for the current run.</param>
            /// <param name="prepared">Prepared native dispatch snapshot.</param>
            public void Revalidate(object editorObject, object projectObject, object prepared)
            { Revalidate(editorObject, projectObject, prepared, false); }

            /// <summary>Performs the last revalidation that permits one bounded focus recovery before callback correlation is armed.</summary>
            /// <param name="editorObject">Owning VBE automation object.</param>
            /// <param name="projectObject">Project authorized for the current run.</param>
            /// <param name="prepared">Prepared native dispatch snapshot.</param>
            internal void RevalidateBeforeArming(object editorObject, object projectObject, object prepared)
            { Revalidate(editorObject, projectObject, prepared, true); }

            /// <summary>Rechecks project, design mode, support source, native pane selection, focus, and Run control identity.</summary>
            /// <param name="editorObject">Owning VBE automation object.</param>
            /// <param name="projectObject">Project authorized for the current run.</param>
            /// <param name="prepared">Prepared native dispatch snapshot.</param>
            /// <param name="allowFocusRecovery">Whether this pass may use its one permitted recovery when only HWND focus was lost.</param>
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

            /// <summary>Executes the verified built-in Run Sub command exactly once after final revalidation.</summary>
            /// <param name="prepared">Prepared native dispatch snapshot.</param>
            public void Execute(object prepared)
            {
                var plan = (Prepared)prepared;
                Revalidate(plan.Editor, plan.Project, plan);
                ((dynamic)plan.Control).Execute();
            }

            /// <summary>Reads the live VBE project mode used to verify native execution completion.</summary>
            /// <param name="project">Live project whose Mode property is observed.</param>
            /// <returns>VBE mode value returned by the project.</returns>
            public int ReadMode(object project) => (int)((dynamic)project).Mode;

            /// <summary>Reads the complete support module source within the inspection line limit.</summary>
            /// <param name="module">Support CodeModule whose text is revalidated before execution.</param>
            /// <returns>Complete source text, or an exception when its line count exceeds 200,000.</returns>
            private static string ReadSource(dynamic module)
            {
                int count = (int)module.CountOfLines;
                if (count < 0 || count > 200000) throw new InvalidOperationException("The support source exceeds inspection limits.");
                return count == 0 ? string.Empty : (string)module.Lines[1, count];
            }
        }
    }
}
