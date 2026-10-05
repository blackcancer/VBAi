using System;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class AccessStaTestMethodAttributeTests
    {
        [ThreadStatic] private static int testMarker;

        [DataTestMethod]
        [DataRow(false)] [DataRow(true)]
        public void EachInvocationOwnsAFreshStaLoopAndPreservesItsOriginalOutcome(bool throws)
        {
            int caller = Thread.CurrentThread.ManagedThreadId; var original = new InvalidOperationException("original synthetic failure");
            var first = AccessStaTestMethodAttribute.RunOnFreshSta(() => {
                Assert.AreEqual(ApartmentState.STA, Thread.CurrentThread.GetApartmentState());
                Assert.IsTrue(Application.MessageLoop); Assert.AreEqual(0, testMarker); testMarker = 1;
                if (throws) throw original; return 17;
            });
            Assert.AreNotEqual(caller, first.OwnerThread);
            Assert.IsTrue(first.DispatcherDisposed); Assert.IsFalse(first.Retained); Assert.IsTrue(first.ThreadExitObserved);
            if (throws) Assert.AreSame(original, first.Error); else { Assert.IsNull(first.Error); Assert.AreEqual(17, first.Value); }
            var second = AccessStaTestMethodAttribute.RunOnFreshSta(() => testMarker);
            Assert.AreEqual(0, second.Value); Assert.AreNotEqual(first.OwnerThread, second.OwnerThread);
            Assert.IsTrue(second.DispatcherDisposed); Assert.IsFalse(second.Retained); Assert.IsTrue(second.ThreadExitObserved);
        }

        [DataTestMethod]
        [DataRow(false)] [DataRow(true)]
        public void UnfinishedTaskParksOnlyItsOriginalDispatcherAndCannotContaminateTheNextScenario(bool throws)
        {
            var pending = new TaskCompletionSource<object>(); var original = new InvalidOperationException("first failure");
            var retained = AccessStaTestMethodAttribute.RunOnFreshSta(() => {
                testMarker = 99;
                Assert.AreSame(pending.Task, AccessStaTestMethodAttribute.StartSave(() => pending.Task));
                if (throws) throw original; return 23;
            });
            Assert.IsTrue(retained.Retained); Assert.IsFalse(retained.DispatcherDisposed); Assert.IsFalse(retained.ThreadExitObserved);
            if (throws) Assert.AreSame(original, retained.Error); else Assert.AreEqual(23, retained.Value);
            StringAssert.Contains(retained.Diagnostic, "LateCompletionAccepted=false");
            StringAssert.Contains(retained.Diagnostic, "SaveReplay=0");
            var clean = AccessStaTestMethodAttribute.RunOnFreshSta(() => {
                Assert.AreEqual(0, testMarker);
                return AccessStaTestMethodAttribute.StartSave(() => Task.FromResult<object>("next")).GetAwaiter().GetResult();
            });
            Assert.IsNull(clean.Error); Assert.AreEqual("next", clean.Value); Assert.IsFalse(clean.Retained);
            Assert.AreNotEqual(retained.OwnerThread, clean.OwnerThread);
            // Synthetic completion changes no already-returned verdict and does not release the retained test dispatcher.
            pending.SetResult("late");
            Assert.IsTrue(retained.Retained); Assert.IsFalse(retained.DispatcherDisposed);
            if (throws) Assert.AreSame(original, retained.Error);
        }

        [DataTestMethod]
        [DataRow(false)] [DataRow(true)]
        public void RealVbeUiTaskTraceIsCapturedBeforeAdmissionAndKeepsOwnerStaContinuations(bool throws)
        {
            var original = new InvalidOperationException("async original");
            var run = AccessStaTestMethodAttribute.RunOnFreshSta(() => {
                int owner = Thread.CurrentThread.ManagedThreadId;
                var gate = new TaskCompletionSource<object>();
                var pending = AccessStaTestMethodAttribute.StartSave(() => VbeUiTask.Run(async () => {
                    await gate.Task;
                    Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                    Assert.AreEqual(ApartmentState.STA, Thread.CurrentThread.GetApartmentState());
                    if (throws) throw original; return (object)42;
                }));
                Assert.IsFalse(pending.IsCompleted);
                SynchronizationContext.Current.Post(_ => gate.SetResult(null), null);
                var clock = Stopwatch.StartNew();
                while (!pending.IsCompleted && clock.Elapsed < TimeSpan.FromSeconds(5)) Application.DoEvents();
                Assert.IsTrue(pending.IsCompleted, AccessStaTestMethodAttribute.Describe(pending));
                if (throws) Assert.AreSame(original, Assert.ThrowsException<InvalidOperationException>(() => pending.GetAwaiter().GetResult()));
                else Assert.AreEqual(42, pending.GetAwaiter().GetResult());
                string trace = AccessStaTestMethodAttribute.Describe(pending);
                foreach (string phase in new[] { "OwnerSta", "ContinuationEnqueued", "ContinuationEntered", "ContinuationReturned" })
                    StringAssert.Contains(trace, phase);
                return trace;
            });
            Assert.IsNull(run.Error, run.Error?.ToString()); Assert.IsFalse(run.Retained); Assert.IsTrue(run.DispatcherDisposed);
        }

        [DataTestMethod]
        [DataRow(false)] [DataRow(true)]
        public void OriginalDataRowArrayAndFailureObjectsSurviveRetention(bool retained)
        {
            var original = new AssertFailedException("original bound failed");
            var pass = new TestResult { Outcome = UnitTestOutcome.Passed, DisplayName = "row (1)", LogOutput = "before" };
            var fail = new TestResult { Outcome = UnitTestOutcome.Failed, DisplayName = "row (2)", TestFailureException = original };
            var rows = new[] { pass, fail };
            Assert.AreSame(rows, AccessStaTestMethodAttribute.PreserveResults(rows, retained, "retained details"));
            Assert.AreEqual(2, rows.Length); Assert.AreSame(pass, rows[0]); Assert.AreSame(fail, rows[1]);
            Assert.AreSame(original, fail.TestFailureException); Assert.AreEqual(UnitTestOutcome.Failed, fail.Outcome);
            Assert.AreEqual(retained ? UnitTestOutcome.Failed : UnitTestOutcome.Passed, pass.Outcome);
            Assert.AreEqual("row (1)", pass.DisplayName); Assert.AreEqual("row (2)", fail.DisplayName);
            if (retained) { StringAssert.Contains(fail.LogOutput, "retained details"); StringAssert.Contains(pass.LogOutput, "before"); }
            else { Assert.IsNull(fail.LogOutput); Assert.AreEqual("before", pass.LogOutput); }
        }

        [TestMethod]
        public void SaveTrackingCannotRunOutsideItsDedicatedScope()
        {
            int attempts = 0;
            Assert.ThrowsException<InvalidOperationException>(() => AccessStaTestMethodAttribute.StartSave(() => {
                attempts++; return Task.FromResult<object>(null);
            }));
            Assert.AreEqual(0, attempts);
        }

        [DataTestMethod]
        [DataRow(false)] [DataRow(true)]
        public void NullOrSynchronouslyThrowingAdmissionDoesNotInventAnOutstandingTask(bool throws)
        {
            var original = new InvalidOperationException("admission refused"); int attempts = 0;
            var run = AccessStaTestMethodAttribute.RunOnFreshSta(() => AccessStaTestMethodAttribute.StartSave(() => {
                attempts++; if (throws) throw original; return null;
            }));
            Assert.AreEqual(1, attempts); Assert.IsFalse(run.Retained); Assert.IsTrue(run.DispatcherDisposed);
            if (throws) Assert.AreSame(original, run.Error); else Assert.IsInstanceOfType(run.Error, typeof(InvalidOperationException));
        }

        [TestMethod]
        public void RetainedDispatcherAcceptsLatePostingWithoutExecutingAnotherCallback()
        {
            SynchronizationContext context = null; int callbacks = 0;
            var pending = new TaskCompletionSource<object>();
            var run = AccessStaTestMethodAttribute.RunOnFreshSta(() => {
                context = SynchronizationContext.Current;
                AccessStaTestMethodAttribute.StartSave(() => pending.Task);
                // Queue before publication, behind this original callback. Retention must not pump it.
                context.Post(_ => Interlocked.Increment(ref callbacks), null);
                return false;
            });
            Assert.IsTrue(run.Retained); Assert.IsFalse(run.DispatcherDisposed); Assert.IsNotNull(context);
            context.Post(_ => Interlocked.Increment(ref callbacks), null); // Original HWND remains alive; this must not throw.
            Assert.AreEqual(0, Volatile.Read(ref callbacks));
            Assert.IsFalse(pending.Task.IsCompleted); Assert.IsFalse(run.Value);
        }

        [DataTestMethod]
        [DataRow(false)] [DataRow(true)]
        public void AlreadyFaultedOrCanceledTasksAreTerminalAndNeverRetained(bool canceled)
        {
            var expected = new InvalidOperationException("faulted terminal");
            var run = AccessStaTestMethodAttribute.RunOnFreshSta(() => {
                var task = new TaskCompletionSource<object>();
                if (canceled) task.SetCanceled(); else task.SetException(expected);
                AccessStaTestMethodAttribute.StartSave(() => task.Task);
                try { task.Task.GetAwaiter().GetResult(); } catch (Exception error) { if (!canceled) Assert.AreSame(expected, error); }
                return task.Task.Status;
            });
            Assert.IsNull(run.Error); Assert.IsFalse(run.Retained); Assert.IsTrue(run.DispatcherDisposed);
            Assert.AreEqual(canceled ? TaskStatus.Canceled : TaskStatus.Faulted, run.Value);
        }
    }
}
