using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbaTestOwnerAwaitableTests
    {
        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void NativeCompletionAndFailureResumeCatchAndFinallyOnOwnerWithNoSynchronizationContext(bool fail)
        {
            var previous = SynchronizationContext.Current;
            SynchronizationContext.SetSynchronizationContext(null);
            try
            {
                int owner = Thread.CurrentThread.ManagedThreadId;
                using (var queue = new BlockingCollection<Action>())
                {
                    var source = new TaskCompletionSource<int>(TaskCreationOptions.RunContinuationsAsynchronously);
                    var threads = new List<int>();
                    var phases = new List<string>();
                    var awaiting = Observe(source.Task, queue.Add, () =>
                    {
                        if (owner != Thread.CurrentThread.ManagedThreadId) throw new InvalidOperationException("Wrong owner thread.");
                    }, threads, phases);
                    Assert.IsFalse(awaiting.IsCompleted);
                    int worker = Task.Run(() =>
                    {
                        if (fail) source.SetException(new InvalidOperationException("Native completion failed.")); else source.SetResult(42);
                        return Thread.CurrentThread.ManagedThreadId;
                    }).GetAwaiter().GetResult();
                    Assert.AreNotEqual(owner, worker);
                    Pump(awaiting, queue);
                    Assert.AreEqual(fail ? -1 : 42, awaiting.GetAwaiter().GetResult());
                    CollectionAssert.AreEqual(fail ? new[] { "catch", "finally" } : new[] { "result", "finally" }, phases);
                    CollectionAssert.AreEqual(new[] { owner, owner }, threads);
                    Assert.IsNull(SynchronizationContext.Current);
                }
            }
            finally { SynchronizationContext.SetSynchronizationContext(previous); }
        }

        private static async Task<int> Observe(Task<int> task, Action<Action> post, Action owner, List<int> threads, List<string> phases)
        {
            try
            {
                int value = await new VbaTestOwnerAwaitable<int>(task, post, owner);
                threads.Add(Thread.CurrentThread.ManagedThreadId); phases.Add("result");
                return value;
            }
            catch (InvalidOperationException)
            { threads.Add(Thread.CurrentThread.ManagedThreadId); phases.Add("catch"); return -1; }
            finally { owner(); threads.Add(Thread.CurrentThread.ManagedThreadId); phases.Add("finally"); }
        }

        [TestMethod]
        public void SynchronouslyCompletedTasksStillRequireTheirOwnerBeforeReturningValues()
        {
            int owner = Thread.CurrentThread.ManagedThreadId;
            var awaitable = new VbaTestOwnerAwaitable<int>(Task.FromResult(42), _ => Assert.Fail("No continuation should be queued."),
                () => { if (Thread.CurrentThread.ManagedThreadId != owner) throw new InvalidOperationException("Wrong owner thread."); });
            Assert.AreEqual(42, awaitable.GetAwaiter().GetResult());
            var error = Task.Run(() =>
            {
                try { awaitable.GetAwaiter().GetResult(); return null; }
                catch (Exception caught) { return caught; }
            }).GetAwaiter().GetResult();
            Assert.IsInstanceOfType(error, typeof(InvalidOperationException));
        }

        internal static void Pump(Task task, BlockingCollection<Action> queue)
        {
            var deadline = DateTime.UtcNow.AddSeconds(10);
            while (!task.IsCompleted && DateTime.UtcNow < deadline)
            {
                Action continuation;
                if (queue.TryTake(out continuation, 50)) continuation();
            }
            Assert.IsTrue(task.IsCompleted, "The explicit owner queue did not complete the test continuation.");
        }
    }
}
