using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbeUiTaskTests
    {
        [STATestMethod]
        public void RealAwaitResumesOnOriginalStaWithNullAmbientContext()
        {
            VerifyStaContinuation(null);
        }

        [STATestMethod]
        public void RealAwaitResumesOnOriginalStaWithBaseAmbientContext()
        {
            VerifyStaContinuation(new SynchronizationContext());
        }

        [STATestMethod]
        public void AsynchronousFailurePropagatesAndRestoresAmbientContext()
        {
            var original = SynchronizationContext.Current;
            var ambient = new SynchronizationContext();
            try
            {
                SynchronizationContext.SetSynchronizationContext(ambient);
                int ownerThread = Thread.CurrentThread.ManagedThreadId;
                var expected = new InvalidOperationException("async failure");
                Task<int> pending = VbeUiTask.Run(async () =>
                {
                    await Task.Delay(30);
                    Assert.AreEqual(ownerThread, Thread.CurrentThread.ManagedThreadId);
                    Assert.AreEqual(ApartmentState.STA, Thread.CurrentThread.GetApartmentState());
                    throw expected;
#pragma warning disable CS0162
                    return 0;
#pragma warning restore CS0162
                });
                Assert.AreSame(ambient, SynchronizationContext.Current);
                PumpUntilComplete(pending);
                Assert.AreSame(expected, Assert.ThrowsException<InvalidOperationException>(() => pending.GetAwaiter().GetResult()));
            }
            finally { SynchronizationContext.SetSynchronizationContext(original); }
        }

        private static void VerifyStaContinuation(SynchronizationContext ambient)
        {
            var original = SynchronizationContext.Current;
            try
            {
                SynchronizationContext.SetSynchronizationContext(ambient);
                int ownerThread = Thread.CurrentThread.ManagedThreadId;
                Assert.AreEqual(ApartmentState.STA, Thread.CurrentThread.GetApartmentState());
                Task<int> pending = VbeUiTask.Run(async () =>
                {
                    Assert.AreEqual(ownerThread, Thread.CurrentThread.ManagedThreadId);
                    Assert.AreEqual(ApartmentState.STA, Thread.CurrentThread.GetApartmentState());
                    Assert.IsNotNull(SynchronizationContext.Current);
                    await Task.Delay(30);
                    Assert.AreEqual(ownerThread, Thread.CurrentThread.ManagedThreadId);
                    Assert.AreEqual(ApartmentState.STA, Thread.CurrentThread.GetApartmentState());
                    Assert.IsNotNull(SynchronizationContext.Current);
                    return 42;
                });
                Assert.AreSame(ambient, SynchronizationContext.Current);
                PumpUntilComplete(pending);
                Assert.AreEqual(42, pending.GetAwaiter().GetResult());
            }
            finally { SynchronizationContext.SetSynchronizationContext(original); }
        }

        private static void PumpUntilComplete(Task pending)
        {
            var elapsed = Stopwatch.StartNew();
            while (!pending.IsCompleted && elapsed.ElapsedMilliseconds < 3000)
            {
                Application.DoEvents();
                Thread.Sleep(1);
            }
            Assert.IsTrue(pending.IsCompleted, "The pinned STA operation did not finish while pumping messages.");
        }
    }
}
