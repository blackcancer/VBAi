using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks that asynchronous event handlers cannot outlive their owned STA dispatcher.</summary>
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class VbeUiTaskEventLifetimeTests
    {
        [STATestMethod]
        [DataRow(false, "success"), DataRow(true, "success")]
        [DataRow(false, "failure"), DataRow(true, "failure")]
        [DataRow(false, "cancel"), DataRow(true, "cancel")]
        public void OutstandingEventDefersCleanupAndPreservesTheInitiatingOutcome(bool completeOnWorker, string outcome)
        {
            var original = SynchronizationContext.Current;
            var ambient = new SynchronizationContext();
            var failure = new ApplicationException("initiating work failed");
            SynchronizationContext context = null;
            int publications = 0, owner = Thread.CurrentThread.ManagedThreadId;
            Exception published = null;
            using (var cancellation = new CancellationTokenSource())
            {
                cancellation.Cancel();
                try
                {
                    SynchronizationContext.SetSynchronizationContext(ambient);
                    var pending = VbeUiTask.RunWithCompletion(() =>
                    {
                        context = SynchronizationContext.Current;
                        context.OperationStarted();
                        return outcome == "failure" ? Task.FromException<int>(failure) :
                            outcome == "cancel" ? Task.FromCanceled<int>(cancellation.Token) : Task.FromResult(7);
                    }, error => { Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId); publications++; published = error; });
                    Assert.AreSame(ambient, SynchronizationContext.Current);
                    Assert.IsFalse(pending.IsCompleted, "Cleanup must wait for the admitted event, even when the initiating task settles synchronously.");
                    Assert.AreEqual(0, publications);
                    if (completeOnWorker) Task.Run(() => context.OperationCompleted()).GetAwaiter().GetResult();
                    else context.OperationCompleted();
                    Pump(pending);
                    Assert.AreEqual(1, publications);
                    if (outcome == "failure")
                    { Assert.AreSame(failure, Assert.ThrowsException<ApplicationException>(() => pending.GetAwaiter().GetResult())); Assert.AreSame(failure, published); }
                    else if (outcome == "cancel")
                    { Assert.ThrowsException<TaskCanceledException>(() => pending.GetAwaiter().GetResult()); Assert.IsTrue(pending.IsCanceled); Assert.IsInstanceOfType(published, typeof(OperationCanceledException)); }
                    else { Assert.AreEqual(7, pending.GetAwaiter().GetResult()); Assert.IsNull(published); }
                    Assert.ThrowsException<InvalidOperationException>(() => context.Post(_ => Assert.Fail("Dispatcher must close after its last event."), null));
                }
                finally { SynchronizationContext.SetSynchronizationContext(original); }
            }
        }

        [STATestMethod]
        [DataRow(1, false, "success"), DataRow(2, false, "success")]
        [DataRow(1, true, "success"), DataRow(2, true, "success")]
        [DataRow(1, false, "failure"), DataRow(2, false, "failure")]
        [DataRow(1, true, "failure"), DataRow(2, true, "failure")]
        [DataRow(1, false, "cancel"), DataRow(2, false, "cancel")]
        [DataRow(1, true, "cancel"), DataRow(2, true, "cancel")]
        public void RealAsyncEventsAndTheirNestedEventsFinishOnOwnerBeforeTerminalPublication(int siblings, bool nested, string outcome)
        {
            int owner = Thread.CurrentThread.ManagedThreadId, finished = 0, started = 0, publications = 0;
            var first = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            var second = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            var nestedGate = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            var failure = new ApplicationException("event failure is handled by its caller");
            var token = new CancellationToken(true);

            async void NestedEvent()
            {
                started++;
                await nestedGate.Task;
                Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                Assert.AreEqual(ApartmentState.STA, Thread.CurrentThread.GetApartmentState());
                finished++;
            }
            async void Event(Task release)
            {
                started++;
                try
                {
                    await release;
                    Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                    if (nested) NestedEvent();
                    if (outcome == "failure") throw failure;
                    if (outcome == "cancel") token.ThrowIfCancellationRequested();
                }
                catch (ApplicationException error) { Assert.AreSame(failure, error); }
                catch (OperationCanceledException error) { Assert.AreEqual(token, error.CancellationToken); }
                finally { Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId); finished++; }
            }
            var pending = VbeUiTask.RunWithCompletion(() =>
            {
                Event(first.Task);
                if (siblings == 2) Event(second.Task);
                return Task.FromResult(42);
            }, error => { Assert.IsNull(error); Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId); Assert.AreEqual(started, finished); publications++; });
            Assert.IsFalse(pending.IsCompleted);
            first.SetResult(true);
            PumpUntil(() => finished >= 1);
            if (siblings == 2)
            { Assert.IsFalse(pending.IsCompleted); second.SetResult(true); PumpUntil(() => finished >= siblings); }
            if (nested)
            { Assert.IsFalse(pending.IsCompleted); nestedGate.SetResult(true); }
            Pump(pending);
            Assert.AreEqual(42, pending.GetAwaiter().GetResult());
            Assert.AreEqual(siblings * (nested ? 2 : 1), finished);
            Assert.AreEqual(1, publications);
        }

        [STATestMethod]
        public void ANewEventAdmittedBeforeTheDrainContinuationResumesRetainsTheDispatcher()
        {
            SynchronizationContext context = null;
            var pending = VbeUiTask.Run(() =>
            {
                context = SynchronizationContext.Current;
                context.OperationStarted();
                return Task.FromResult(7);
            });
            context.OperationCompleted();
            context.OperationStarted();
            bool callback = false;
            context.Post(_ => callback = true, null);
            PumpUntil(() => callback);
            Assert.IsFalse(pending.IsCompleted, "The drain must observe the new admission before disposing the handle.");
            context.OperationCompleted();
            Pump(pending);
            Assert.AreEqual(7, pending.GetAwaiter().GetResult());
        }

        [STATestMethod]
        public void UnmatchedEventCompletionIsRejectedWithoutChangingTheNextOwnerOperation()
        {
            var pending = VbeUiTask.Run(() =>
            {
                var context = SynchronizationContext.Current;
                Assert.ThrowsException<InvalidOperationException>(() => context.OperationCompleted());
                return Task.FromResult(7);
            });
            Assert.AreEqual(7, pending.GetAwaiter().GetResult());
        }

        private static void Pump(Task task) => PumpUntil(() => task.IsCompleted);

        private static void PumpUntil(Func<bool> terminal)
        {
            var clock = Stopwatch.StartNew();
            while (!terminal())
            {
                if (clock.ElapsedMilliseconds > 5000) Assert.Fail("Owned event did not settle on its STA.");
                Application.DoEvents(); Thread.Sleep(1);
            }
        }
    }
}
