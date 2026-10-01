using System;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VBAi
{
    internal sealed partial class VbeTestExplorerService
    {
        private readonly Action<object> defaultCompileCoverageDelegate;
        private Action currentCoverageCompileGuard;
        internal Func<Action, IDisposable> StartCoverageCompilationTimer = StartCompilationTimer;
        internal Func<long> CoverageCompilationClock = () => (long)(1000d * Stopwatch.GetTimestamp() / Stopwatch.Frequency);
        private const int CompilationDeadlineMilliseconds = 3000;

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

        /// <summary>Observes compilation only after returning to the owner UI; never executes the command again.</summary>
        internal Task<bool> VerifyCoverageCompilationAsync(object copiedProject, Action guard)
        {
            RequireOwner();
            if (copiedProject == null) throw new ArgumentNullException(nameof(copiedProject));
            if (guard == null) throw new ArgumentNullException(nameof(guard));
            var completion = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            long started = CoverageCompilationClock();
            IDisposable timer = null;
            EventHandler dispatcherDisposed = null;
            Action<Exception> finish = error => {
                timer?.Dispose(); timer = null;
                dispatcher.Disposed -= dispatcherDisposed;
                if (error == null) completion.TrySetResult(true);
                else completion.TrySetException(error);
            };
            Action tick = () => {
                if (completion.Task.IsCompleted) return;
                try
                {
                    RequireOwner(); guard();
                    RequireCompilationProject(copiedProject);
                    if (CoverageCompilationClock() - started >= CompilationDeadlineMilliseconds)
                        throw new InvalidOperationException("The coverage copy did not reach the verified compiled state within three seconds; no tests were dispatched.");
                    dynamic compile = ReadCompilationControl();
                    if (!(bool)compile.Enabled)
                    {
                        // COM getters can pump messages; confirm authority and context after the final read.
                        guard(); RequireCompilationProject(copiedProject);
                        finish(null);
                    }
                }
                catch (Exception error) { finish(error); }
            };
            dispatcherDisposed = (sender, args) => finish(new ObjectDisposedException(nameof(VbeTestExplorerService)));
            dispatcher.Disposed += dispatcherDisposed;
            try
            {
                // Even a control already disabled is re-read only after an owner UI turn.
                dispatcher.BeginInvoke(new Action(() => {
                    if (completion.Task.IsCompleted) return;
                    try
                    {
                        RequireOwner();
                        timer = StartCoverageCompilationTimer(tick);
                        if (timer == null) throw new InvalidOperationException("The owner compilation observer could not be started.");
                        if (completion.Task.IsCompleted) { timer.Dispose(); timer = null; }
                    }
                    catch (Exception error) { finish(error); }
                }));
            }
            catch (Exception error) { finish(error); }
            return completion.Task;
        }

        private void RequireCompilationProject(object copiedProject)
        {
            RequireOwner();
            object activeProject = (object)vbe.ActiveVBProject;
            if (!ReferenceEquals(activeProject, copiedProject) && !VbeDebug.NativeProcedureValuesHost.SameComIdentity(activeProject, copiedProject))
                throw new InvalidOperationException("The native compiler does not target the exact owned coverage copy.");
            if ((int)((dynamic)copiedProject).Mode != 2)
                throw new InvalidOperationException("The coverage compiler requires the owned project in design mode.");
        }

        private object ReadCompilationControl()
        {
            dynamic compile = vbe.CommandBars.FindControl(1, 578);
            if (compile == null || (int)compile.Id != 578 || (int)compile.Type != 1 || !(bool)compile.BuiltIn
                || !string.IsNullOrEmpty((string)compile.OnAction))
                throw new InvalidOperationException("The native VBA compile command is unavailable.");
            return compile;
        }

        private static IDisposable StartCompilationTimer(Action tick)
        {
            return StartCompilationTimer(tick, timer => timer.Start());
        }

        internal static IDisposable StartCompilationTimer(Action tick, Action<Timer> start)
        {
            var timer = new Timer { Interval = 50 };
            timer.Tick += (sender, args) => tick();
            try { start(timer); return timer; }
            catch { timer.Dispose(); throw; }
        }
    }
}
