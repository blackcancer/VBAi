namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;
    using System.Reflection;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Windows.Forms;
    using VBAi;

    public sealed partial class GitWindowCoverageTests
    {
        private sealed class Fixture : IDisposable
        {
            internal readonly MacroGitOperationsTests.Fixture Git = new MacroGitOperationsTests.Fixture();
            internal readonly GitWindow Window;
            internal readonly string UiCache;
            private readonly Func<string, string> previousCache = GitWindow.CacheDirectory;
            private readonly SynchronizationContext previousContext = SynchronizationContext.Current;
            private Exception uiError;
            private Exception operationFailure;
            private readonly HashSet<Task> operations = new HashSet<Task>();
            private bool disposed;
            internal Exception CapturedUiFailure => uiError;
            internal Fixture(bool bound = true)
            {
                UiCache = Path.Combine(Git.Root, "ui");
                GitWindow.CacheDirectory = _ => UiCache;
                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
                Application.ThreadException += OnUiError;
                Window = new GitWindow(Git.Project, "fixture", "Disposable coverage window");
                if (bound) Set("repository", Git.Repository);
                Window.Show();
                var connect = Window.ConnectRepository;
                Window.ConnectRepository = (repository, url) => connect(repository, Git.Remote);
            }
            internal T Get<T>(string name) => (T)typeof(GitWindow).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(Window);
            internal void Set(string name, object value) => typeof(GitWindow).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).SetValue(Window, value);
            internal object Call(string name, params object[] args)
            {
                try
                {
                    object result = typeof(GitWindow).GetMethod(name, BindingFlags.Instance | BindingFlags.NonPublic).Invoke(Window, args);
                    if (result is Task task) operations.Add(task);
                    return result;
                }
                catch (TargetInvocationException ex) { System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(ex.InnerException).Throw(); throw; }
            }
            internal void Pump(Task task = null)
            {
                if (task != null) operations.Add(task);
                try
                {
                    var clock = Stopwatch.StartNew();
                    while ((task != null && !task.IsCompleted) || Get<bool>("running"))
                    {
                        Application.DoEvents(); Thread.Sleep(1);
                        if (uiError != null) throw new InvalidOperationException("Disposable GitWindow UI failed", uiError);
                        if (clock.ElapsedMilliseconds > 40000) throw new TimeoutException("Disposable GitWindow operation timed out");
                    }
                    task?.GetAwaiter().GetResult(); Application.DoEvents();
                }
                catch (Exception error) { operationFailure = error; throw; }
            }
            internal void Event(string name, object sender = null)
            {
                Call(name, sender, EventArgs.Empty);
                Pump();
            }
            internal void Action(string action, string name = null, string text = null, string choice = null, string path = null)
            { Pump((Task)Call("RunGitAction", action, name, text, choice, path)); }
            internal void Compare() => Event("Compare_Click");
            internal void AssertLiveMatchesHead()
            {
                var expected = Git.Repository.Read(Git.Repository.Resolve(Git.Repository.Head));
                var observed = Git.Project.Capture();
                string details = string.Join("\n", observed.Changes(expected));
                foreach (var entry in observed.Serialize())
                    if (expected.Serialize().TryGetValue(entry.Key, out var baseline) && !System.Linq.Enumerable.SequenceEqual(entry.Value, baseline))
                        details += "\n" + entry.Key + " expected: " + System.Text.Encoding.UTF8.GetString(baseline) + " observed: " + System.Text.Encoding.UTF8.GetString(entry.Value);
                Assert.IsTrue(observed.SameAs(expected), details);
            }
            internal void Select(ListBox list, int index) { Set("running", true); list.SelectedIndex = index; Set("running", false); }
            internal string Status => Get<Label>("status").Text;
            private void OnUiError(object sender, ThreadExceptionEventArgs e) { uiError = e.Exception; }
            public void Dispose()
            {
                if (disposed) return;
                disposed = true;
                var failures = new List<Exception>();
                Action<Action> cleanup = action => { try { action(); } catch (Exception error) { failures.Add(error); } };
                // Async-void handlers expose their lifetime through running; reflective
                // Task operations are tracked separately. Never delete a live cache.
                bool terminal = false;
                cleanup(() => terminal = !Get<bool>("running") && System.Linq.Enumerable.All(operations, task => task.IsCompleted));
                if (terminal) cleanup(() => Window.Dispose());
                else failures.Add(new InvalidOperationException("Disposable GitWindow work is still active. The window and owned Git scratch directory were retained: " + Git.Root));
                cleanup(() => Application.ThreadException -= OnUiError);
                cleanup(() => GitWindow.CacheDirectory = previousCache);
                cleanup(() => SynchronizationContext.SetSynchronizationContext(previousContext));
                if (terminal) cleanup(() => Git.Dispose());
                if (failures.Count == 0) return;
                // A using statement must not replace the original operation failure
                // with a later disposal or locked-file error.
                if (operationFailure != null && !failures.Contains(operationFailure)) failures.Insert(0, operationFailure);
                if (failures.Count == 1) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failures[0]).Throw();
                throw new AggregateException("Disposable GitWindow operation and cleanup failures; owned artifacts may be retained at " + Git.Root, failures);
            }
        }
    }
}
