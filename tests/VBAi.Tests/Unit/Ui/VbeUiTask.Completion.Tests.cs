using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbeUiTaskCompletionTests
    {
        private sealed class NonReturningContext : SynchronizationContext
        {
            private int posts;
            internal int Posts => Volatile.Read(ref posts);
            public override void Post(SendOrPostCallback callback, object state) { Interlocked.Increment(ref posts); }
        }

        [STATestMethod]
        [DataRow(false, "null", "success")]
        [DataRow(false, "null", "failure")]
        [DataRow(false, "null", "cancel")]
        [DataRow(false, "base", "success")]
        [DataRow(false, "base", "failure")]
        [DataRow(false, "base", "cancel")]
        [DataRow(false, "lost", "success")]
        [DataRow(false, "lost", "failure")]
        [DataRow(false, "lost", "cancel")]
        [DataRow(true, "null", "success")]
        [DataRow(true, "null", "failure")]
        [DataRow(true, "null", "cancel")]
        [DataRow(true, "base", "success")]
        [DataRow(true, "base", "failure")]
        [DataRow(true, "base", "cancel")]
        [DataRow(true, "lost", "success")]
        [DataRow(true, "lost", "failure")]
        [DataRow(true, "lost", "cancel")]
        public void TerminalPublicationAndReturnedTaskRemainOnOwnerIndependentOfAmbientContext(bool asynchronous, string ambientKind, string outcome)
        {
            var previous = SynchronizationContext.Current;
            var discarded = new NonReturningContext();
            SynchronizationContext ambient = ambientKind == "null" ? null : ambientKind == "lost" ? discarded : new SynchronizationContext();
            var error = new InvalidOperationException("original work failure");
            var gate = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            SynchronizationContext owned = null; int callbacks = 0, owner = Thread.CurrentThread.ManagedThreadId;
            Exception published = null;
            using (var cancelled = new CancellationTokenSource())
            {
                cancelled.Cancel();
                try
                {
                    SynchronizationContext.SetSynchronizationContext(ambient);
                    Task<int> pending = VbeUiTask.RunWithCompletion(async () =>
                    {
                        owned = SynchronizationContext.Current;
                        Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                        if (asynchronous) await gate.Task;
                        Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                        if (outcome == "failure") throw error;
                        if (outcome == "cancel") cancelled.Token.ThrowIfCancellationRequested();
                        return 42;
                    }, failure =>
                    {
                        callbacks++; published = failure;
                        Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                        Assert.AreEqual(ApartmentState.STA, Thread.CurrentThread.GetApartmentState());
                        bool rejected = false;
                        try { owned.Post(_ => Assert.Fail("Dispatcher is already disposed."), null); }
                        catch (InvalidOperationException) { rejected = true; }
                        Assert.IsTrue(rejected, "Terminal publication must follow dispatcher cleanup.");
                    });
                    Assert.AreSame(ambient, SynchronizationContext.Current);
                    if (asynchronous)
                    {
                        Assert.IsFalse(pending.IsCompleted); Assert.AreEqual(0, callbacks);
                        Task.Run(() => gate.SetResult(true)).GetAwaiter().GetResult();
                    }
                    Pump(pending);
                    if (outcome == "success") { Assert.AreEqual(42, pending.GetAwaiter().GetResult()); Assert.IsNull(published); }
                    else if (outcome == "failure") { Assert.AreSame(error, Assert.ThrowsException<InvalidOperationException>(() => pending.GetAwaiter().GetResult())); Assert.AreSame(error, published); }
                    else
                    {
                        Assert.IsTrue(pending.IsCanceled);
                        try { pending.GetAwaiter().GetResult(); Assert.Fail("Cancellation lost."); }
                        catch (OperationCanceledException failure) { Assert.AreEqual(cancelled.Token, failure.CancellationToken); }
                        Assert.IsInstanceOfType(published, typeof(OperationCanceledException));
                    }
                    Assert.AreEqual(1, callbacks); Assert.AreEqual(0, discarded.Posts);
                }
                finally { SynchronizationContext.SetSynchronizationContext(previous); }
            }
        }

        [STATestMethod]
        [DataRow("success")]
        [DataRow("failure")]
        [DataRow("cancel")]
        public void PublicationFailureFaultsTheWholeTaskAndRetainsTheOriginalWorkOutcome(string outcome)
        {
            var work = new InvalidOperationException("work"); var publish = new ApplicationException("publication");
            int callbacks = 0; Exception supplied = null;
            using (var cancellation = new CancellationTokenSource())
            {
                cancellation.Cancel();
                var task = VbeUiTask.RunWithCompletion(() => outcome == "success" ? Task.FromResult(7) :
                    outcome == "failure" ? Task.FromException<int>(work) : Task.FromCanceled<int>(cancellation.Token),
                    error => { callbacks++; supplied = error; throw publish; });
                Assert.IsTrue(task.IsCompleted); Assert.AreEqual(1, callbacks);
                if (outcome == "success") { Assert.IsNull(supplied); Assert.AreSame(publish, Assert.ThrowsException<ApplicationException>(() => task.GetAwaiter().GetResult())); }
                else
                {
                    var combined = Assert.ThrowsException<AggregateException>(() => task.GetAwaiter().GetResult());
                    CollectionAssert.AreEqual(new[] { supplied, publish }, combined.InnerExceptions.ToArray());
                    if (outcome == "failure") Assert.AreSame(work, supplied);
                    else Assert.IsInstanceOfType(supplied, typeof(OperationCanceledException));
                }
            }
        }

        [DataTestMethod]
        [DataRow(false, false, false)]
        [DataRow(false, false, true)]
        [DataRow(false, true, false)]
        [DataRow(false, true, true)]
        [DataRow(true, false, false)]
        [DataRow(true, false, true)]
        [DataRow(true, true, false)]
        [DataRow(true, true, true)]
        public void CleanupAndPublicationAreAttemptedOnceAndEveryDistinctFailureIsRetained(bool workFails, bool cleanupFails, bool publicationFails)
        {
            var work = new InvalidOperationException("work"); var cleanup = new ApplicationException("cleanup"); var publication = new Exception("publication");
            var log = new List<string>(); Exception seen = null;
            var failure = VbeUiTask.FinishOwnerOperation(workFails ? work : null,
                () => { log.Add("cleanup"); if (cleanupFails) throw cleanup; },
                error => { log.Add("publication"); seen = error; if (publicationFails) throw publication; });
            CollectionAssert.AreEqual(new[] { "cleanup", "publication" }, log);
            var expected = new[] { workFails ? work : null, cleanupFails ? cleanup : null, publicationFails ? publication : null }.Where(error => error != null).ToArray();
            if (expected.Length == 0) Assert.IsNull(failure);
            else if (expected.Length == 1) Assert.AreSame(expected[0], failure);
            else CollectionAssert.AreEquivalent(expected, ((AggregateException)failure).Flatten().InnerExceptions.ToArray());
            var prior = expected.Where(error => !ReferenceEquals(error, publication)).ToArray();
            if (prior.Length == 0) Assert.IsNull(seen);
            else if (prior.Length == 1) Assert.AreSame(prior[0], seen);
            else CollectionAssert.AreEqual(prior, ((AggregateException)seen).InnerExceptions.ToArray());
        }

        [DataTestMethod]
        [DataRow(false, true)]
        [DataRow(true, false)]
        [DataRow(true, true)]
        public void TheSameExceptionIdentityIsNotDuplicatedByCleanupOrPublication(bool cleanupFails, bool publicationFails)
        {
            var original = new Exception("same failure");
            var result = VbeUiTask.FinishOwnerOperation(original,
                () => { if (cleanupFails) throw original; }, error => { Assert.AreSame(original, error); if (publicationFails) throw original; });
            Assert.AreSame(original, result);
        }

        [TestMethod]
        public void InvalidCompletionDependenciesOrForeignApartmentCannotStartWork()
        {
            Assert.ThrowsException<ArgumentNullException>(() => VbeUiTask.RunWithCompletion<int>(null, _ => Assert.Fail()));
            Assert.ThrowsException<ArgumentNullException>(() => VbeUiTask.RunWithCompletion(() => Task.FromResult(1), null));
            Assert.ThrowsException<ArgumentNullException>(() => VbeUiTask.FinishOwnerOperation(null, null, _ => Assert.Fail()));
            Exception observed = null; int calls = 0;
            var thread = new Thread(() => { try { VbeUiTask.RunWithCompletion(() => { calls++; return Task.FromResult(1); }, _ => calls++); } catch (Exception error) { observed = error; } });
            thread.SetApartmentState(ApartmentState.MTA); thread.Start(); Assert.IsTrue(thread.Join(5000));
            Assert.IsInstanceOfType(observed, typeof(InvalidOperationException)); Assert.AreEqual(0, calls);
        }

        private static void Pump(Task task)
        {
            var time = Stopwatch.StartNew();
            while (!task.IsCompleted)
            {
                if (time.ElapsedMilliseconds > 5000) Assert.Fail("Owned terminal task must settle without an ambient callback.");
                Application.DoEvents(); Thread.Sleep(1);
            }
        }
    }
}
