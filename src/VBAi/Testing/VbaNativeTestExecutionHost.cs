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
        internal interface IProbe
        {
            void RequireOwner(object vbe);
            object Prepare(object vbe, object project, string source);
            void Revalidate(object vbe, object project, object prepared);
            void Execute(object prepared);
            int ReadMode(object project);
        }

        private readonly object vbe;
        private readonly VbaTestResultSink sink;
        private readonly Func<VbaTestCatalog, object> resolveProject;
        private readonly Action<VbaTestCatalog> validateCatalog;
        private readonly Action executionGuard;
        private readonly Func<VbaTestCatalog, string> signature;
        private readonly Func<string> runId;
        private readonly int owner = Thread.CurrentThread.ManagedThreadId;
        private Call active;
        private bool disposed, uncertain;
        internal IProbe Probe = new NativeProbe();
        internal Action<Action> Post;
        internal Func<Action, IDisposable> StartPolling;
        internal TimeSpan VerificationTimeout = TimeSpan.FromSeconds(30);

        private sealed class Call
        {
            internal VbaTestCatalog Catalog;
            internal VbaTestDescriptor Test;
            internal string Phase;
            internal object Project;
            internal bool Armed, Invoked;
            internal Stopwatch Clock = Stopwatch.StartNew();
            internal IDisposable Exposure, Polling;
            internal TaskCompletionSource<VbaTestResult> Completion = new TaskCompletionSource<VbaTestResult>(TaskCreationOptions.RunContinuationsAsynchronously);
        }

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

        public void Validate(VbaTestCatalog catalog)
        {
            RequireOwner();
            if (disposed || uncertain) throw new VbaTestInvocationException("The live native transport is disconnected or its preceding outcome is uncertain.", false);
            executionGuard();
            Probe.RequireOwner(vbe);
            validateCatalog(catalog);
        }

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

        private void Dispatch(Call call)
        {
            if (!ReferenceEquals(active, call) || call.Completion.Task.IsCompleted) return;
            try
            {
                Validate(call.Catalog);
                call.Project = resolveProject(call.Catalog);
                var support = call.Catalog.Project.Modules.Single(module => string.Equals(module.Name, VbaTestRuntimeSource.ModuleName, StringComparison.OrdinalIgnoreCase));
                object prepared = Probe.Prepare(vbe, call.Project, support.Source);
                // Showing a pane can pump messages. Revalidate all authority after selection.
                Validate(call.Catalog);
                if (!VbeDebug.NativeProcedureValuesHost.SameComIdentity(call.Project, resolveProject(call.Catalog)) && Probe is NativeProbe)
                    throw new InvalidOperationException("The selected project identity changed during native preparation.");
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
            catch (Exception error) { Finish(call, error); }
        }

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

        private static void Revoke(Call call)
        { call.Polling?.Dispose(); call.Polling = null; call.Exposure?.Dispose(); call.Exposure = null; }

        private void RequireOwner()
        { if (owner != Thread.CurrentThread.ManagedThreadId) throw new InvalidOperationException("Native test transport requires its owning thread."); }

        public void Dispose()
        {
            RequireOwner(); disposed = true;
            if (active != null) Finish(active, new InvalidOperationException("The live native test transport disconnected."));
        }

        internal static void ValidateNativeRunControl(object controlObject, string expectedCaption = null)
        {
            dynamic control = controlObject;
            if (control == null || (int)control.Id != 186 || (int)control.Type != 1 || !(bool)control.BuiltIn
                || !(bool)control.Enabled || !string.IsNullOrEmpty((string)control.OnAction)
                || !VbeDebug.IsAllowed("run", (string)control.Caption, 2)
                || (expectedCaption != null && (string)control.Caption != expectedCaption))
                throw new InvalidOperationException("The verified built-in native Run Sub command (186) is unavailable or changed.");
        }

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
            if (activeWindow == null || !identity((object)window, activeWindow) || (int)window.Type != 0 || !(bool)window.Visible)
                throw new InvalidOperationException("The exact support code window does not own native focus before dispatch.");
            dynamic pane = editor.ActiveCodePane;
            if (pane == null || !identity(expectedPane, (object)pane) || !identity(moduleObject, (object)pane.CodeModule))
                throw new InvalidOperationException("The support code pane no longer owns the native run selection.");
            int line = 0, column = 0, endLine = 0, endColumn = 0;
            pane.GetSelection(ref line, ref column, ref endLine, ref endColumn);
            if (line != expectedLine || endLine != expectedLine || column != 1 || endColumn != 1
                || (int)((dynamic)moduleObject).ProcBodyLine[VbaTestRuntimeSource.PendingProcedure, 0] != expectedLine)
                throw new InvalidOperationException("The exact pending wrapper body no longer owns the native selection.");
        }

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

        internal static void ValidateRecoveredNativeFocus(object editorObject, object paneObject,
            VbaNativeTestWindowFocus.IWindows windows, VbaNativeTestWindowFocus.Target target)
        {
            if (target == null) return;
            if (new IntPtr(Convert.ToInt64(((dynamic)editorObject).MainWindow.HWnd)) != target.Main.Handle
                || !string.Equals((string)((dynamic)paneObject).Window.Caption, target.Caption, StringComparison.Ordinal))
                throw new InvalidOperationException("The COM-to-native code-window mapping changed after focus recovery.");
            VbaNativeTestWindowFocus.VerifyFocus(windows, target);
        }

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

        private static string DiagnosticWindow(object window)
        {
            return "Type=" + DiagnosticRead(() => ((dynamic)window).Type)
                + ",Visible=" + DiagnosticRead(() => ((dynamic)window).Visible)
                + ",HWnd=" + DiagnosticRead(() => ((dynamic)window).HWnd)
                + ",State=" + DiagnosticRead(() => ((dynamic)window).WindowState);
        }

        private static class NativeWindowObservation
        {
            private delegate bool EnumWindow(IntPtr window, IntPtr unused);
            [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, EnumWindow callback, IntPtr unused);
            [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr window);
            [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
            [DllImport("user32.dll")] private static extern IntPtr GetParent(IntPtr window);
            [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder text, int count);
            [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();

            internal static string Read(IntPtr main)
            {
                uint owner;
                uint thread = VbeDebugWindows.GetWindowThreadProcessId(main, out owner);
                using (var process = Process.GetCurrentProcess())
                    if (main == IntPtr.Zero || !IsWindow(main) || owner != (uint)process.Id || thread != GetCurrentThreadId())
                        return "unavailable(owner mismatch)";
                var result = new StringBuilder();
                Append(result, main, owner);
                int count = 0;
                EnumChildWindows(main, (window, unused) =>
                {
                    if (count++ >= 64) { result.Append("bounded"); return false; }
                    uint childOwner;
                    VbeDebugWindows.GetWindowThreadProcessId(window, out childOwner);
                    if (childOwner == owner) Append(result, window, childOwner);
                    return result.Length < 3900;
                }, IntPtr.Zero);
                return result.ToString();
            }

            private static void Append(StringBuilder result, IntPtr window, uint owner)
            {
                var kind = new StringBuilder(96);
                GetClassName(window, kind, kind.Capacity);
                result.Append("{HWnd=" + window.ToInt64() + ",PID=" + owner + ",Class=" + kind
                    + ",Parent=" + GetParent(window).ToInt64() + ",Visible=" + IsWindowVisible(window) + "}");
            }
        }

        // Read only after a refusal. Failed inspection must preserve the original refusal,
        // and must never reselect a pane, focus a window or dispatch a native command.
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

        private static bool DiagnosticMatches(Func<object, object, bool> identity, object expected, object actual)
        {
            if (expected == null || actual == null) return false;
            try { return identity(expected, actual); }
            catch { return false; }
        }

        private static string DiagnosticProcedure(object module, int line)
        {
            int kind = 0;
            string name = (string)((dynamic)module).ProcOfLine[line, ref kind];
            return name + " (kind=" + kind + ")";
        }

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

        private sealed class NativeProbe : IProbe
        {
            [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
            private sealed class Prepared
            {
                internal object Editor, Project, Module, Pane, Control;
                internal string Source, Caption, AfterShow, AfterFocus;
                internal int Line;
                internal VbaNativeTestWindowFocus.Target NativeFocus;
            }

            private readonly VbaNativeTestWindowFocus.IWindows windows = new VbaNativeTestWindowFocus.NativeWindows();

            public void RequireOwner(object editor)
            {
                if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA) throw new InvalidOperationException("Native VBE execution requires an STA.");
                var window = new IntPtr(Convert.ToInt64(((dynamic)editor).MainWindow.HWnd));
                uint process;
                uint thread = VbeDebugWindows.GetWindowThreadProcessId(window, out process);
                using (var current = Process.GetCurrentProcess())
                    if (window == IntPtr.Zero || process != (uint)current.Id || thread != GetCurrentThreadId())
                        throw new InvalidOperationException("The VBE window does not belong to this process and owning UI thread.");
            }

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
                try { pane = PrepareNativePane(editorObject, (object)module, line, shown => afterShow = DescribeNativeWindows(editorObject, shown, (object)module)); }
                catch (Exception error)
                {
                    if (afterShow != null) throw new InvalidOperationException(error.Message + " Immediately after Show: " + afterShow, error);
                    throw;
                }
                RequireOwner(editorObject);
                string afterFocus = DescribeNativeWindows(editorObject, pane, (object)module);
                VbaNativeTestWindowFocus.Target nativeFocus;
                try { nativeFocus = EnsureNativePaneFocus(editorObject, projectObject, (object)module, pane, line, windows); }
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
                return new Prepared { Editor = editorObject, Project = projectObject, Module = module, Pane = pane, Control = control, Source = source, Caption = (string)control.Caption, Line = line, AfterShow = afterShow, AfterFocus = afterFocus, NativeFocus = nativeFocus };
            }

            public void Revalidate(object editorObject, object projectObject, object prepared)
            {
                RequireOwner(editorObject);
                var plan = (Prepared)prepared;
                dynamic editor = editorObject, project = projectObject, control = plan.Control;
                bool found = false;
                foreach (object candidate in editor.VBProjects)
                    if (VbeDebug.NativeProcedureValuesHost.SameComIdentity(candidate, projectObject)) { found = true; break; }
                if (!found || (int)project.Mode != 2 || ReadSource(plan.Module) != plan.Source)
                    throw new InvalidOperationException("The selected project, mode or support source changed before native dispatch.");
                try
                {
                    ValidateNativePaneContext(editorObject, projectObject, plan.Module, plan.Pane, plan.Line);
                    ValidateRecoveredNativeFocus(editorObject, plan.Pane, windows, plan.NativeFocus);
                }
                catch (Exception error)
                {
                    throw new InvalidOperationException(error.Message
                        + DescribeNativeSelection(editorObject, projectObject, plan.Module, plan.Pane, plan.Line)
                        + " Immediately after Show: " + plan.AfterShow + " Immediately after Focus: " + plan.AfterFocus, error);
                }
                ValidateNativeRunControl((object)control, plan.Caption);
            }

            public void Execute(object prepared)
            {
                var plan = (Prepared)prepared;
                Revalidate(plan.Editor, plan.Project, plan);
                ((dynamic)plan.Control).Execute();
            }
            public int ReadMode(object project) => (int)((dynamic)project).Mode;
            private static string ReadSource(dynamic module)
            {
                int count = (int)module.CountOfLines;
                if (count < 0 || count > 200000) throw new InvalidOperationException("The support source exceeds inspection limits.");
                return count == 0 ? string.Empty : (string)module.Lines[1, count];
            }
        }
    }
}
