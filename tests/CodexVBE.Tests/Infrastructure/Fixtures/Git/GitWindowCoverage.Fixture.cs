namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Diagnostics;
    using System.IO;
    using System.Reflection;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Windows.Forms;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

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
                try { return typeof(GitWindow).GetMethod(name, BindingFlags.Instance | BindingFlags.NonPublic).Invoke(Window, args); }
                catch (TargetInvocationException ex) { System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(ex.InnerException).Throw(); throw; }
            }
            internal void Pump(Task task = null)
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
            internal void Event(string name, object sender = null)
            {
                Call(name, sender, EventArgs.Empty);
                Pump();
            }
            internal void Action(string action, string name = null, string text = null, string choice = null, string path = null)
            { Pump((Task)Call("RunGitAction", action, name, text, choice, path)); }
            internal void Compare() => Event("Compare_Click");
            internal void Select(ListBox list, int index) { Set("running", true); list.SelectedIndex = index; Set("running", false); }
            internal string Status => Get<Label>("status").Text;
            private void OnUiError(object sender, ThreadExceptionEventArgs e) { uiError = e.Exception; }
            public void Dispose()
            {
                Window.Dispose(); Application.ThreadException -= OnUiError; GitWindow.CacheDirectory = previousCache;
                SynchronizationContext.SetSynchronizationContext(previousContext); Git.Dispose();
            }
        }
    }
}
