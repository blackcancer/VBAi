using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;
using System.Threading.Tasks;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    public sealed partial class ChatGitModalDiagnosticTests
    {
        [STATestMethod, DataRow(false), DataRow(true)]
        public void OwnerDispatcherKeepsAsyncDiagnosticDisposalAndPostOnStaWithoutAmbientContext(bool fail)
        {
            int thread = System.Threading.Thread.CurrentThread.ManagedThreadId;
            Action requireOwner = () =>
            {
                Assert.AreEqual(thread, System.Threading.Thread.CurrentThread.ManagedThreadId);
                Assert.AreEqual(System.Threading.ApartmentState.STA, System.Threading.Thread.CurrentThread.GetApartmentState());
            };
            var o = new Observation(); var d = o.Begin(); var gate = new TaskCompletionSource<bool>();
            var original = new InvalidOperationException("session failed"); int disposed = 0;
            var previous = System.Threading.SynchronizationContext.Current; Task operation;
            try
            {
                System.Threading.SynchronizationContext.SetSynchronizationContext(null);
                operation = VbeUiTask.Run(async () =>
                {
                    try
                    {
                        await d.RunModalAsync(async () => { await gate.Task; requireOwner(); if (fail) throw original; },
                            () => { requireOwner(); disposed++; });
                    }
                    catch (Exception error) { requireOwner(); Assert.AreSame(original, error); d.Fail(error); }
                    finally { requireOwner(); d.SchedulePostHandler(callback => { requireOwner(); o.Post(callback); }); }
                    return true;
                });
                Assert.IsNull(System.Threading.SynchronizationContext.Current);
            }
            finally { System.Threading.SynchronizationContext.SetSynchronizationContext(previous); }
            Assert.AreEqual(0, disposed); Assert.IsFalse(operation.IsCompleted);
            gate.SetResult(true); VBAi.Tests.Infrastructure.LlmBoundaryScope.Pump(operation);
            Assert.AreEqual(1, disposed); Assert.AreEqual(fail ? 0 : 1, o.Posts);
            if (!fail) { o.Callback(); Assert.IsTrue(d.Completed); }
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public async Task PendingAsyncDiagnosticRejectsBothConcurrentEntryRoutesWithoutDisposingAgain(bool synchronousSecond)
        {
            var o = new Observation(); var d = o.Begin(); var gate = new TaskCompletionSource<bool>(); int disposals = 0;
            var first = d.RunModalAsync(() => gate.Task, () => disposals++);
            if (synchronousSecond) Assert.ThrowsException<InvalidOperationException>(() => d.RunModal(() => Assert.Fail(), () => disposals++));
            else await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => d.RunModalAsync(() => { Assert.Fail(); return Task.CompletedTask; }, () => disposals++));
            Assert.AreEqual(0, disposals); Assert.IsFalse(first.IsCompleted);
            gate.SetResult(true); await first; Assert.AreEqual(1, disposals);
        }

        [TestMethod]
        public async Task CompletedSynchronousDiagnosticCannotStartAnAsyncSession()
        {
            var o = new Observation(); var d = o.Begin(); int disposals = 0;
            d.RunModal(() => { }, () => disposals++);
            await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => d.RunModalAsync(() => { Assert.Fail(); return Task.CompletedTask; }, () => disposals++));
            Assert.AreEqual(1, disposals);
        }
        [TestMethod]
        public async Task IntermediateModalReturnCannotPublishFinalDisposalOrPostHandlerSuccess()
        {
            var o = new Observation(); var d = o.Begin(); int disposals = 0;
            var completedSession = new TaskCompletionSource<bool>();
            var run = d.RunModalAsync(() => completedSession.Task, () => disposals++);
            Assert.IsFalse(run.IsCompleted); Assert.AreEqual(0, disposals);
            CollectionAssert.AreEqual(new[] { "Started" }, o.Attempts.ToArray());
            completedSession.SetResult(true); await run;
            Assert.AreEqual(1, disposals); d.SchedulePostHandler(o.Post); Assert.IsFalse(d.Completed);
            o.Callback(); Assert.IsTrue(d.Completed); ChatGitDiagnosticReceipt.Validate(o.Chain.ToArray(), Nonce, Identity());
        }

        [DataTestMethod, DataRow("session"), DataRow("dispose"), DataRow("both")]
        public async Task AsyncSessionAndDisposalErrorsRemainOriginalAndNeverAuthorizePostHandler(string stage)
        {
            var o = new Observation(); var d = o.Begin(); int disposals = 0;
            var first = new InvalidOperationException("session"); var second = new InvalidOperationException("dispose");
            Exception observed = null;
            try
            {
                await d.RunModalAsync(() => stage == "dispose" ? Task.CompletedTask : Task.FromException(first),
                () => { disposals++; if (stage != "session") throw second; });
            }
            catch (Exception error) { observed = error; }
            Assert.AreEqual(1, disposals);
            if (stage == "both") CollectionAssert.AreEqual(new Exception[] { first, second }, ((AggregateException)observed).InnerExceptions.ToArray());
            else Assert.AreSame(stage == "session" ? first : second, observed);
            d.Fail(observed); d.SchedulePostHandler(o.Post); Assert.AreEqual(0, o.Posts); Assert.IsFalse(d.Completed);
        }

        [DataTestMethod, DataRow("ShowModalReturned"), DataRow("DisposeReturned")]
        public async Task AsyncTerminalPublicationFailureStillDisposesOnceAndCannotBecomeSuccess(string phase)
        {
            var o = new Observation { ThrowPhase = phase }; var d = o.Begin(); int disposals = 0;
            await Assert.ThrowsExceptionAsync<System.IO.IOException>(() => d.RunModalAsync(() => Task.CompletedTask, () => disposals++));
            d.SchedulePostHandler(o.Post); Assert.AreEqual(1, disposals); Assert.AreEqual(0, o.Posts); Assert.IsFalse(d.Completed);
        }
    }
}
