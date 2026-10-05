using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi;

namespace VBAi.Tests.Unit
{
    /// <summary>Isolates the synthetic Access scenario and any unfinished dispatcher on a fresh STA.</summary>
    internal sealed class AccessStaTestMethodAttribute : TestMethodAttribute
    {
        [ThreadStatic] private static Scope current;
        private static readonly List<Scope> retained = new List<Scope>();

        internal sealed class Tracked
        {
            internal Task Task;
            internal readonly ConcurrentQueue<string> Events = new ConcurrentQueue<string>();
            internal string Diagnostic => "TaskStatus=" + Task.Status + "; TraceBound=" + VbeInspectionTrace.MaximumEvents +
                "; Trace=" + string.Join(Environment.NewLine, Events.ToArray());
        }

        private sealed class Scope
        {
            internal readonly List<Tracked> Tasks = new List<Tracked>();
            internal readonly ManualResetEventSlim Park = new ManualResetEventSlim();
            internal Control Dispatcher;
            internal Thread Thread;
            internal int Owner;
        }

        internal sealed class RunResult<T>
        {
            internal T Value;
            internal Exception Error;
            internal bool Retained, DispatcherDisposed, ThreadExitObserved;
            internal int OwnerThread;
            internal string Diagnostic;
        }

        public override TestResult[] Execute(ITestMethod testMethod)
        {
            var run = RunOnFreshSta(() => base.Execute(testMethod));
            if (run.Error != null) ExceptionDispatchInfo.Capture(run.Error).Throw();
            return PreserveResults(run.Value, run.Retained, run.Diagnostic);
        }

        internal static TestResult[] PreserveResults(TestResult[] results, bool isRetained, string diagnostic)
        {
            if (results == null || results.Length == 0) throw new InvalidOperationException("The original MSTest result is unavailable.");
            if (!isRetained) return results;
            foreach (var result in results)
            {
                result.LogOutput = (result.LogOutput ?? "") + Environment.NewLine + diagnostic;
                if (result.Outcome == UnitTestOutcome.Passed)
                {
                    result.Outcome = UnitTestOutcome.Failed;
                    result.TestFailureException = new AssertFailedException("The synthetic Access scenario returned with an unfinished original task. " + diagnostic);
                }
            }
            return results; // Existing failure objects/DataRow results are never replaced or rehabilitated.
        }

        /// <summary>Starts a real test message loop; failure retention parks it without pumping or retrying after the test returns.</summary>
        internal static RunResult<T> RunOnFreshSta<T>(Func<T> invoke)
        {
            if (invoke == null) throw new ArgumentNullException(nameof(invoke));
            var run = new RunResult<T>();
            var finished = new ManualResetEventSlim();
            var scope = new Scope();
            var adapterClock = Stopwatch.StartNew();
            scope.Thread = new Thread(() => {
                current = scope; scope.Owner = Thread.CurrentThread.ManagedThreadId; run.OwnerThread = scope.Owner;
                try
                {
                    using (var dispatcher = new Control())
                    {
                        scope.Dispatcher = dispatcher;
                        _ = dispatcher.Handle;
                        dispatcher.BeginInvoke(new Action(() => {
                            try { run.Value = invoke(); }
                            catch (Exception error) { run.Error = error; }
                            finally
                            {
                                run.Retained = scope.Tasks.Any(item => !item.Task.IsCompleted);
                                if (run.Retained)
                                {
                                    run.Diagnostic = "SYNTHETIC_STA_RETAINED_UNTIL_TESTHOST_EXIT; OwnerThread=" + scope.Owner +
                                        "; OriginalFailureUnchanged=true; SaveReplay=0; LateCompletionAccepted=false; " +
                                        string.Join(Environment.NewLine, scope.Tasks.Select(item => item.Diagnostic));
                                    lock (retained) retained.Add(scope);
                                    finished.Set();
                                    // Keep HWND/STA alive for a late BeginInvoke, but run no more callbacks after the failed bound.
                                    // This is fake-only test isolation, never host recovery or proof of the first timeout's cause.
                                    scope.Park.Wait();
                                }
                                else Application.ExitThread();
                            }
                        }));
                        Application.Run();
                    }
                    run.DispatcherDisposed = true;
                }
                catch (Exception error) { run.Error = run.Error == null ? error : new AggregateException(run.Error, error); }
                finally { current = null; finished.Set(); }
            }) { IsBackground = true };
            scope.Thread.SetApartmentState(ApartmentState.STA);
            scope.Thread.Start();
            // This bounds the test adapter itself, as the existing WinForms attribute does. Access remains bounded to five seconds.
            TimeSpan initialRemaining = TimeSpan.FromMinutes(2) - adapterClock.Elapsed;
            if (initialRemaining <= TimeSpan.Zero || !finished.Wait(initialRemaining))
                throw new TimeoutException("The synthetic Access test thread did not return its original result; no thread termination or replay is attempted.");
            if (!run.Retained)
            {
                TimeSpan remaining = TimeSpan.FromMinutes(2) - adapterClock.Elapsed;
                if (remaining <= TimeSpan.Zero || !scope.Thread.Join(remaining))
                    throw new TimeoutException("The original synthetic Access STA did not exit within the original adapter bound.");
                run.ThreadExitObserved = true;
                finished.Dispose(); scope.Park.Dispose();
            }
            return run;
        }

        /// <summary>Capture before VbeUiTask.Run: its owning context snapshots the trace during admission.</summary>
        internal static Task<object> StartSave(Func<Task<object>> start)
        {
            if (start == null) throw new ArgumentNullException(nameof(start));
            if (current == null || current.Owner != Thread.CurrentThread.ManagedThreadId)
                throw new InvalidOperationException("Synthetic Access Save requires its dedicated test STA scope.");
            var tracked = new Tracked();
            var trace = new VbeInspectionTrace(tracked.Events.Enqueue);
            using (trace.Enter()) tracked.Task = start();
            if (tracked.Task == null) throw new InvalidOperationException("The original Access task is absent.");
            current.Tasks.Add(tracked);
            return (Task<object>)tracked.Task;
        }

        internal static string Describe(Task task)
        {
            var tracked = current?.Tasks.SingleOrDefault(item => ReferenceEquals(item.Task, task));
            return "OwnerThread=" + Thread.CurrentThread.ManagedThreadId + "; Apartment=" + Thread.CurrentThread.GetApartmentState() +
                "; MessageLoop=" + Application.MessageLoop + "; " + (tracked?.Diagnostic ?? "OriginalTaskTraceUnavailable");
        }
    }
}
