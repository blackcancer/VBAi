using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Exercises synchronous STA dispatch and resource ownership without Office activation.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class OfficeCoverageEventDispatcherTests
    {
        [TestMethod]
        public void MtaCannotCreateTheOwnedStaMessageTarget()
        {
            Exception failure = null;
            var thread = new Thread(() =>
            {
                try { using (var target = new OfficeCoverageEventDispatcher()) { } }
                catch (Exception error) { failure = error; }
            });
            thread.SetApartmentState(ApartmentState.MTA); thread.Start();
            Assert.IsTrue(thread.Join(10000));
            Assert.IsInstanceOfType(failure, typeof(InvalidOperationException));
        }

        [TestMethod]
        public void OwnerCallbacksAreImmediateAndDisposedTargetsRejectFurtherWork()
        {
            OnSta(() =>
            {
                var target = new OfficeCoverageEventDispatcher();
                int calls = 0;
                target.Invoke(() => calls++);
                Assert.AreEqual(1, calls);
                Assert.ThrowsException<ArgumentNullException>(() => target.Invoke(null));
                var failure = new InvalidOperationException("callback failure");
                Assert.AreSame(failure, Assert.ThrowsException<InvalidOperationException>(() => target.Invoke(() => { throw failure; })));
                target.Dispose(); target.Dispose();
                Assert.ThrowsException<ObjectDisposedException>(() => target.Invoke(() => calls++));
                Assert.AreEqual(1, calls);
            });
        }

        [TestMethod]
        public void WorkerWaitsForTheOwnerAndCannotConsumeDisposal()
        {
            OnSta(() =>
            {
                using (var target = new OfficeCoverageEventDispatcher())
                {
                    int owner = Thread.CurrentThread.ManagedThreadId, calls = 0;
                    var task = Task.Run(() =>
                    {
                        Assert.ThrowsException<InvalidOperationException>(target.Dispose);
                        target.Invoke(() => { Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId); calls++; });
                        Assert.AreEqual(1, calls, "Synchronous dispatch must complete before returning the borrowed argument.");
                    });
                    Pump(task);
                    target.Invoke(() => calls++);
                    Assert.AreEqual(2, calls);
                }
            });
        }

        [TestMethod]
        public void WorkerReceivesTheOriginalCallbackFailureWithoutRepeatingIt()
        {
            OnSta(() =>
            {
                using (var target = new OfficeCoverageEventDispatcher())
                {
                    int calls = 0;
                    var expected = new InvalidOperationException("borrowed getter failure");
                    var task = Task.Run(() =>
                    {
                        var failure = Assert.ThrowsException<InvalidOperationException>(() =>
                            target.Invoke(() => { calls++; throw expected; }));
                        Assert.AreSame(expected, failure);
                    });
                    Pump(task);
                    Assert.AreEqual(1, calls);
                }
            });
        }

        private static void Pump(Task task)
        {
            DateTime deadline = DateTime.UtcNow.AddSeconds(10);
            while (!task.IsCompleted && DateTime.UtcNow < deadline) { Application.DoEvents(); Thread.Sleep(1); }
            Assert.IsTrue(task.IsCompleted, "The owner did not process its queued callback within the test bound.");
            task.GetAwaiter().GetResult();
        }

        internal static void OnSta(Action action)
        {
            Exception failure = null;
            var thread = new Thread(() => { try { action(); } catch (Exception error) { failure = error; } }) { IsBackground = true };
            thread.SetApartmentState(ApartmentState.STA); thread.Start();
            Assert.IsTrue(thread.Join(30000), "The isolated STA scenario did not terminate.");
            if (failure != null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
        }
    }
}
