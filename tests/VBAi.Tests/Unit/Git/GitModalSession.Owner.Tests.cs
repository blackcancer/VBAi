using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    public sealed partial class GitWindowCoverageTests
    {
        [WinFormsTestMethod]
        [DataRow(false)][DataRow(true)]
        public void ExactOwnerLeaseRefusesDuplicateBeforeAnyShowOrDiagnosticEvenWhenDisabled(bool disabled)
        {
            using (var owner = new Form())
            using (var window = new GitWindow())
            using (var lease = GitModalSession.TryAcquire(owner))
            {
                owner.Enabled = !disabled;
                Assert.IsTrue(GitModalSession.IsActive(owner));
                Assert.IsNull(GitModalSession.TryAcquire(owner));
                int shows = 0;
                var task = GitModalSession.ShowAsync(window, owner, (form, parent) => { shows++; return DialogResult.Cancel; });
                Assert.ThrowsException<InvalidOperationException>(() => task.GetAwaiter().GetResult());
                Assert.AreEqual(0, shows);
            }
        }

        [WinFormsTestMethod]
        public void DifferentOwnedWindowsHaveIndependentLeasesAndDisposedLeaseCanBeReacquired()
        {
            using (var first = new Form())
            using (var second = new Form())
            {
                var one = GitModalSession.TryAcquire(first);
                using (var two = GitModalSession.TryAcquire(second))
                {
                    Assert.IsNotNull(two); one.Dispose(); one.Dispose();
                    Assert.IsFalse(GitModalSession.IsActive(first)); Assert.IsTrue(GitModalSession.IsActive(second));
                    using (var again = GitModalSession.TryAcquire(first)) Assert.IsNotNull(again);
                }
                Assert.IsFalse(GitModalSession.IsActive(second));
            }
        }

        [WinFormsTestMethod]
        [DataRow(false)][DataRow(true)]
        public void OriginalLeaseSurvivesShowReentrancyAndCanRunOnlyOnce(bool failShow)
        {
            using (var owner = new Form())
            using (var window = new GitWindow())
            using (var lease = GitModalSession.TryAcquire(owner))
            {
                var original = new InvalidOperationException("show failed"); int shows = 0;
                var task = lease.ShowAsync(window, (form, parent) => {
                    shows++; Assert.IsNull(GitModalSession.TryAcquire(owner));
                    Assert.ThrowsException<InvalidOperationException>(() => lease.Dispose());
                    if (failShow) throw original; return DialogResult.Cancel;
                });
                if (failShow) Assert.AreSame(original, Assert.ThrowsException<InvalidOperationException>(() => task.GetAwaiter().GetResult()));
                else task.GetAwaiter().GetResult();
                Assert.ThrowsException<InvalidOperationException>(() => lease.ShowAsync(window, (form, parent) => DialogResult.Cancel));
                Assert.AreEqual(1, shows);
            }
        }

        [WinFormsTestMethod]
        public void PendingRequestRetainsOwnerUntilItsWholeSessionTerminal()
        {
            using (var f = new Fixture())
            using (var owner = new Form())
            using (var lease = GitModalSession.TryAcquire(owner))
            {
                f.Window.Hide(); int shows = 0;
                var finish = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
                var request = new GitModalSession.Request("pull", null, null, null, null, null, false, "revision");
                Task work = null;
                var sessionTask = lease.ShowAsync(f.Window, (form, parent) => {
                    if (++shows == 1)
                    {
                        var session = f.Get<GitModalSession>("modalSession");
                        work = CompleteOwnedRequest(session.Queue(request, () => { }), request, finish.Task);
                    }
                    return DialogResult.Cancel;
                });
                Assert.IsFalse(sessionTask.IsCompleted); Assert.IsNull(GitModalSession.TryAcquire(owner));
                Assert.ThrowsException<InvalidOperationException>(() => lease.Dispose());
                finish.SetResult(true); f.Pump(sessionTask); work.GetAwaiter().GetResult();
                Assert.AreEqual(2, shows); Assert.IsTrue(GitModalSession.IsActive(owner));
            }
        }

        private static async Task CompleteOwnedRequest(Task admission, GitModalSession.Request request, Task finish)
        { await admission; await finish; request.Complete(null); }

        [TestMethod]
        public void MtaEntryRefusesBeforeAttachingAWindowOrAcquiringAnOwnerLease()
        {
            Exception failure = null;
            var thread = new Thread(() => {
                try
                {
                    using (var owner = new Form())
                    using (var window = new GitWindow())
                    {
                        int shows = 0;
                        var task = GitModalSession.ShowAsync(window, owner, (form, parent) => { shows++; return DialogResult.Cancel; });
                        Assert.ThrowsException<InvalidOperationException>(() => task.GetAwaiter().GetResult());
                        Assert.AreEqual(0, shows); Assert.IsFalse(GitModalSession.IsActive(owner));
                    }
                }
                catch (Exception error) { failure = error; }
            });
            thread.SetApartmentState(ApartmentState.MTA); thread.IsBackground = true; thread.Start();
            Assert.IsTrue(thread.Join(10000)); if (failure != null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
        }
    }
}
