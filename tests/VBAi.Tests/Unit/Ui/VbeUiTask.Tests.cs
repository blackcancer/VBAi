using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

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

        [TestMethod]
        public void FaultedSiblingDispatchersCannotConsumeTheOriginalAcceptedOwnerContinuation()
        {
            var events = new ConcurrentQueue<string>();
            var trace = new VbeInspectionTrace(events.Enqueue);
            var gate = new TaskCompletionSource<int>(TaskCreationOptions.RunContinuationsAsynchronously);
            var run = AccessStaTestMethodAttribute.RunOnFreshSta(() =>
            {
                int ownerThread = Thread.CurrentThread.ManagedThreadId;
                var ambient = SynchronizationContext.Current;
                using (trace.Enter())
                {
                    Task<object> first = AccessStaTestMethodAttribute.StartSave(() =>
                    {
                        // StartSave enters its own diagnostic scope; select this test's writer
                        // at admission, when VbeUiTask captures the continuation trace.
                        using (trace.Enter()) return VbeUiTask.Run(async () =>
                        {
                            int value = await gate.Task;
                            Assert.AreEqual(ownerThread, Thread.CurrentThread.ManagedThreadId);
                            Assert.AreEqual(ApartmentState.STA, Thread.CurrentThread.GetApartmentState());
                            return (object)value;
                        });
                    });
                    Assert.IsFalse(first.IsCompleted);
                    for (int index = 0; index < 2; index++)
                    {
                        var failure = new InvalidOperationException("sibling " + index);
                        Task<int> sibling = VbeUiTask.Run(() => Task.FromException<int>(failure));
                        Assert.AreSame(failure, Assert.ThrowsException<InvalidOperationException>(() => sibling.GetAwaiter().GetResult()));
                        Assert.IsFalse(first.IsCompleted, "A faulted sibling must not complete or dispose the original dispatcher.");
                    }
                    Assert.AreSame(ambient, SynchronizationContext.Current);
                    Task released = Task.Run(() => gate.SetResult(42));
                    PumpUntilComplete(first, () => events.Any(line =>
                        line.Contains("\"Phase\":\"ContinuationPostReturned\"") ||
                        line.Contains("\"Phase\":\"ContinuationPostFailed\"")));
                    released.GetAwaiter().GetResult();
                    Assert.AreEqual(42, first.GetAwaiter().GetResult());
                }
                return ownerThread;
            });
            Assert.IsNull(run.Error, run.Error + Environment.NewLine + string.Join(Environment.NewLine, events));
            Assert.IsFalse(run.Retained, run.Diagnostic);
            Assert.IsTrue(run.ThreadExitObserved);
            int owner = run.Value;
            var rows = events.Select(ParseTrace).ToArray();
            Assert.AreEqual(1, rows.Count(row => row.Phase == "ContinuationEnqueued"));
            Assert.AreEqual(1, rows.Count(row => row.Phase == "ContinuationPostReturned"));
            Assert.AreEqual(0, rows.Count(row => row.Phase == "ContinuationPostFailed"));
            Assert.AreEqual(1, rows.Count(row => row.Phase == "ContinuationEntered" && row.ThreadId == owner && row.Apartment == "STA"));
            Assert.AreEqual(1, rows.Count(row => row.Phase == "ContinuationReturned" && row.ThreadId == owner));
            Assert.AreEqual(1, rows.Count(row => row.Phase == "ContinuationPostReturned" && row.ThreadId != owner && row.Apartment == "MTA"));
            Assert.AreEqual(17, (int)VbeInspectionTrace.Phase.OptionsComboInspection, "Existing phase numbers must remain stable.");
        }

        [STATestMethod]
        public void DisposedDispatcherPostReportsFailureAndPreservesItsOriginalException()
        {
            var events = new ConcurrentQueue<string>();
            SynchronizationContext captured = null;
            var trace = new VbeInspectionTrace(events.Enqueue);
            using (trace.Enter())
            {
                Task<int> completed = VbeUiTask.Run(() =>
                {
                    captured = SynchronizationContext.Current;
                    return Task.FromResult(1);
                });
                Assert.AreEqual(1, completed.GetAwaiter().GetResult());
            }
            Assert.IsNotNull(captured);
            bool invoked = false;
            Exception failure = null;
            try { captured.Post(_ => invoked = true, null); }
            catch (Exception error) { failure = error; }
            Assert.IsNotNull(failure, "A disposed owner dispatcher must reject a post.");
            Assert.IsFalse(invoked);
            var rows = events.Select(ParseTrace).ToArray();
            Assert.AreEqual(1, rows.Count(row => row.Phase == "ContinuationEnqueued"));
            Assert.AreEqual(0, rows.Count(row => row.Phase == "ContinuationPostReturned"));
            Assert.AreEqual(1, rows.Count(row => row.Phase == "ContinuationPostFailed" && row.ErrorType == failure.GetType().Name));
        }

        [STATestMethod]
        public void ThrowingTraceWriterCannotChangeSuccessfulOrFailedPostOutcomes()
        {
            var writerFailure = new InvalidOperationException("diagnostic writer unavailable");
            var trace = new VbeInspectionTrace(_ => { throw writerFailure; });
            var gate = new TaskCompletionSource<int>(TaskCreationOptions.RunContinuationsAsynchronously);
            SynchronizationContext captured = null;
            using (trace.Enter())
            {
                Task<int> first = VbeUiTask.Run(async () =>
                {
                    captured = SynchronizationContext.Current;
                    return await gate.Task;
                });
                Task released = Task.Run(() => gate.SetResult(7));
                PumpUntilComplete(first);
                released.GetAwaiter().GetResult();
                Assert.AreEqual(7, first.GetAwaiter().GetResult());
            }
            Exception postFailure = null;
            try { captured.Post(_ => Assert.Fail("Disposed dispatcher callback must not run."), null); }
            catch (Exception error) { postFailure = error; }
            Assert.IsNotNull(postFailure);
            Assert.AreNotSame(writerFailure, postFailure, "Trace failure must not replace the original post failure.");
        }

        private sealed class TraceRow
        {
            public string Phase { get; set; }
            public int ThreadId { get; set; }
            public string Apartment { get; set; }
            public string ErrorType { get; set; }
        }

        private static TraceRow ParseTrace(string line)
        {
            var row = new JavaScriptSerializer().DeserializeObject(line) as IDictionary<string, object>;
            Assert.IsNotNull(row);
            return new TraceRow
            {
                Phase = Convert.ToString(row["Phase"]),
                ThreadId = Convert.ToInt32(row["ThreadId"]),
                Apartment = Convert.ToString(row["Apartment"]),
                ErrorType = Convert.ToString(row["ErrorType"])
            };
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

        private static void PumpUntilComplete(Task pending, Func<bool> postObserved = null)
        {
            var elapsed = Stopwatch.StartNew();
            while ((!pending.IsCompleted || postObserved != null && !postObserved()) && elapsed.ElapsedMilliseconds < 3000)
            {
                Application.DoEvents();
                Thread.Sleep(1);
            }
            Assert.IsTrue(pending.IsCompleted, "The pinned STA operation did not finish while pumping messages.");
            if (postObserved != null) Assert.IsTrue(postObserved(), "The original post had no accepted or failed acknowledgment within the same owner pump deadline.");
        }
    }
}
