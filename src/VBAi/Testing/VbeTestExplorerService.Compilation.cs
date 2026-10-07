using System;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VBAi
{

    /// <summary>Implements guarded compilation observation for disposable coverage projects.</summary>
    internal sealed partial class VbeTestExplorerService
    {

        /// <summary>Original coverage compiler callback retained so temporary test overrides can be restored.</summary>
        private readonly Action<object> defaultCompileCoverageDelegate;

        /// <summary>Revalidates coverage authorization immediately before invoking the native compile command.</summary>
        private Action currentCoverageCompileGuard;

        /// <summary>Creates the owner-thread timer used to observe whether compilation has finished.</summary>
        internal Func<Action, IDisposable> StartCoverageCompilationTimer = StartCompilationTimer;

        /// <summary>Monotonic millisecond clock used for the compile observation deadline.</summary>
        internal Func<long> CoverageCompilationClock = () => (long)(1000d * Stopwatch.GetTimestamp() / Stopwatch.Frequency);

        /// <summary>Maximum time to wait for the compile command to become disabled before refusing test dispatch.</summary>
        private const int CompilationDeadlineMilliseconds = 3000;

        /// <summary>Invokes the built-in compile command only while the exact owned coverage copy is active.</summary>
        /// <param name="copiedProject">Disposable project clone that must remain active and in design mode.</param>
        private void CompileCoverageClone(object copiedProject)
        {
            RequireOwner();
            dynamic runtime = null;
            foreach (dynamic component in ((dynamic)copiedProject).VBComponents)
                if (string.Equals((string)component.Name, VbaCoverageInstrumentation.ModuleName, StringComparison.OrdinalIgnoreCase)) runtime = component;
            if (runtime == null) throw new InvalidOperationException("The coverage runtime module is missing before compilation.");
            dynamic pane = runtime.CodeModule.CodePane;
            pane.Show(); vbe.MainWindow.Visible = true; vbe.ActiveCodePane = pane; pane.Window.SetFocus();
            RequireCompilationProject(copiedProject);
            dynamic compile = ReadCompilationControl();
            if (!(bool)compile.Enabled) return;
            currentCoverageCompileGuard?.Invoke();
            RequireCompilationProject(copiedProject);
            compile = ReadCompilationControl();
            if (!(bool)compile.Enabled) return;
            try
            {
                compile.Execute();
                RequireCompilationProject(copiedProject);
            }
            catch (Exception error)
            { throw new VbaTestInvocationException("Coverage compilation command completion is uncertain; no retry was attempted. " + error.Message, true, error); }
        }

        /// <summary>Waits for a later owner-thread observation that the native compile command is no longer enabled.</summary>
        /// <param name="copiedProject">Exact disposable coverage project expected to remain active during observation.</param>
        /// <param name="guard">Authorization callback rerun before and after each COM observation.</param>
        /// <returns>Task that completes when compilation appears idle; faults on identity loss, timeout, or disposal.</returns>
        /// <remarks>The command is never repeated. A timeout prevents tests from being dispatched.</remarks>
        internal Task<bool> VerifyCoverageCompilationAsync(object copiedProject, Action guard)
        {
            RequireOwner();
            if (copiedProject == null) throw new ArgumentNullException(nameof(copiedProject));
            if (guard == null) throw new ArgumentNullException(nameof(guard));
            var completion = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            long started = CoverageCompilationClock();
            IDisposable timer = null;
            EventHandler dispatcherDisposed = null;
            void finish(Exception error)
            {
                timer?.Dispose(); timer = null;
                dispatcher.Disposed -= dispatcherDisposed;
                if (error == null) completion.TrySetResult(true);
                else completion.TrySetException(error);
            }
            void tick()
            {
                if (completion.Task.IsCompleted) return;
                try
                {
                    RequireOwner(); guard();
                    RequireCompilationProject(copiedProject);
                    dynamic compile = ReadCompilationControl();
                    if (!(bool)compile.Enabled)
                    {
                        // COM getters can pump messages; confirm authority and context after the final read.
                        guard(); RequireCompilationProject(copiedProject);
                        finish(null);
                    }
                    else if (CoverageCompilationClock() - started >= CompilationDeadlineMilliseconds)
                        throw new InvalidOperationException("The coverage compiler is still enabled when observed after the three-second deadline; no tests were dispatched.");
                }
                catch (Exception error) { finish(error); }
            }
            dispatcherDisposed = (sender, args) => finish(new ObjectDisposedException(nameof(VbeTestExplorerService)));
            dispatcher.Disposed += dispatcherDisposed;
            try
            {
                // Even a control already disabled is re-read only after an owner UI turn.
                dispatcher.BeginInvoke(new Action(() =>
                {
                    if (completion.Task.IsCompleted) return;
                    try
                    {
                        RequireOwner();
                        timer = StartCoverageCompilationTimer(tick) ?? throw new InvalidOperationException("The owner compilation observer could not be started.");
                        if (completion.Task.IsCompleted) { timer.Dispose(); timer = null; }
                    }
                    catch (Exception error) { finish(error); }
                }));
            }
            catch (Exception error) { finish(error); }
            return completion.Task;
        }

        /// <summary>Requires the exact coverage clone to be active and in design mode on the owner thread.</summary>
        /// <param name="copiedProject">Owned disposable project clone targeted by the native compiler.</param>
        private void RequireCompilationProject(object copiedProject)
        {
            RequireOwner();
            object activeProject = (object)vbe.ActiveVBProject;
            if (!ReferenceEquals(activeProject, copiedProject) && !VbeDebug.NativeProcedureValuesHost.SameComIdentity(activeProject, copiedProject))
                throw new InvalidOperationException("The native compiler does not target the exact owned coverage copy.");
            if ((int)((dynamic)copiedProject).Mode != 2)
                throw new InvalidOperationException("The coverage compiler requires the owned project in design mode.");
        }

        /// <summary>Finds and validates the built-in VBE Compile command before it can be invoked.</summary>
        /// <returns>Native command-bar control with built-in ID 578 and no assigned macro action.</returns>
        private object ReadCompilationControl()
        {
            dynamic compile = vbe.CommandBars.FindControl(1, 578);
            if (compile == null || (int)compile.Id != 578 || (int)compile.Type != 1 || !(bool)compile.BuiltIn
                || !string.IsNullOrEmpty((string)compile.OnAction))
                throw new InvalidOperationException("The native VBA compile command is unavailable.");
            return compile;
        }

        /// <summary>Starts a WinForms timer that invokes the polling action on the UI thread.</summary>
        /// <param name="tick">Owner-thread observation to run on each timer tick.</param>
        /// <returns>Timer lifetime; disposing stops the timer and detaches its callback.</returns>
        private static IDisposable StartCompilationTimer(Action tick)
        {
            return StartCompilationTimer(tick, timer => timer.Start());
        }

        /// <summary>Creates a 50 ms WinForms timer and arranges for the supplied owner-thread polling callback.</summary>
        /// <param name="tick">Compilation observation performed on every timer tick.</param>
        /// <param name="start">Starts the timer; an exception disposes it before being rethrown.</param>
        /// <returns>Timer lifetime; disposing detaches the event handler and stops the timer.</returns>
        internal static IDisposable StartCompilationTimer(Action tick, Action<Timer> start)
        {
            var timer = new Timer { Interval = 50 };
            timer.Tick += (sender, args) => tick();
            try { start(timer); return timer; }
            catch { timer.Dispose(); throw; }
        }
    }
}
