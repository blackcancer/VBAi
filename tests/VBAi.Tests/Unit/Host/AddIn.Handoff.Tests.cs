using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Unit
{
    public sealed partial class AddInCoverageTests
    {
        [STATestMethod, DataRow(false), DataRow(true)]
        public void GitCallerWithoutAmbientContextDisposesAndReportsOnItsOriginalSta(bool fail)
        {
            using (var scope = new HostUiScope())
            {
                System.IO.File.WriteAllText(scope.Host.Project.FileName, "disposable synthetic document identity");
                var instance = new AddIn(); LlmBoundaryScope.Set(instance, "vbe", scope.Host);
                var probe = new GitHandoffCallerProbe(fail);
                AddIn.ShowModal = probe.Show;
                AddIn.ShowNotice = (text, title, buttons, icon) => { probe.Report(); return DialogResult.OK; };
                try
                {
                    probe.Start(() => Call(instance, "ShowGitHub"));
                    probe.Finish(); Assert.IsFalse(GitModalSession.IsActive(scope.Host.Owner));
                }
                finally { scope.Close(instance); }
            }
        }
    }

    internal sealed class GitHandoffCallerProbe
    {
        private readonly int thread = Thread.CurrentThread.ManagedThreadId;
        private readonly bool fail;
        private readonly TaskCompletionSource<bool> gate = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly TaskCompletionSource<bool> done = new TaskCompletionSource<bool>();
        private Task operation;
        private int shows, disposed, reports;
        internal GitHandoffCallerProbe(bool fail) { this.fail = fail; }
        internal DialogResult Show(Form form, IWin32Window owner)
        {
            AssertThread();
            if (++shows == 1)
            {
                form.Disposed += (sender, args) => { AssertThread(); disposed++; if (!fail) done.SetResult(true); };
                var session = LlmBoundaryScope.Get<GitModalSession>(form, "modalSession");
                var request = new GitModalSession.Request("pull", null, null, null, null, null, false, "revision");
                operation = Complete(session.Queue(request, () => { }), request);
            }
            else if (fail) throw new InvalidOperationException("post-handoff presentation failed");
            return DialogResult.Cancel;
        }
        private async Task Complete(Task admission, GitModalSession.Request request)
        { await admission; await gate.Task; AssertThread(); request.Complete(null); }
        internal void Report() { AssertThread(); reports++; done.TrySetResult(true); }
        internal void Start(Action entry)
        {
            var previous = SynchronizationContext.Current;
            try { SynchronizationContext.SetSynchronizationContext(null); entry(); Assert.IsNull(SynchronizationContext.Current); }
            finally { SynchronizationContext.SetSynchronizationContext(previous); }
            Assert.AreEqual(1, shows); Assert.AreEqual(0, disposed); Assert.IsFalse(done.Task.IsCompleted);
        }
        internal void Finish()
        {
            gate.SetResult(true); LlmBoundaryScope.Pump(done.Task); operation.GetAwaiter().GetResult();
            Assert.AreEqual(2, shows); Assert.AreEqual(1, disposed); Assert.AreEqual(fail ? 1 : 0, reports);
        }
        private void AssertThread()
        { Assert.AreEqual(thread, Thread.CurrentThread.ManagedThreadId); Assert.AreEqual(ApartmentState.STA, Thread.CurrentThread.GetApartmentState()); }
    }
}
