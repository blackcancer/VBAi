using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Threading;
using System.Threading.Tasks;
using VBAi.Tests.Unit.Editor;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks owner-context quiescence when async events are queued before their operation counter starts.</summary>
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class ModernEditorLifetimeFixtureTests
    {
        [WinFormsTestMethod]
        [DataRow("success"), DataRow("failure"), DataRow("cancel")]
        public void QueuedAsynchronousEventSettlesBeforeDrainReturns(string outcome)
        {
            using (var context = new OwnedEditorDispatcher())
            {
                bool entered = false, completed = false;
                int owner = Thread.CurrentThread.ManagedThreadId;
                context.Post(async state =>
                {
                    entered = true;
                    Assert.AreSame(context, SynchronizationContext.Current);
                    await Task.Delay(40);
                    Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                    Assert.AreSame(context, SynchronizationContext.Current);
                    completed = true;
                    if (outcome == "failure") throw new InvalidOperationException("owned event failure");
                    if (outcome == "cancel") throw new OperationCanceledException("owned event cancellation");
                }, null);
                Assert.IsFalse(entered, "Posting alone must not count as execution of the event.");
                context.Drain(outcome == "failure" ? typeof(InvalidOperationException) : outcome == "cancel" ? typeof(OperationCanceledException) : null);
                Assert.IsTrue(completed, "A zero count before message pumping is not owner quiescence.");
            }
        }

        [WinFormsTestMethod]
        public void NestedAdmissionKeepsTheSameOwnerContextAndCannotEscapeTheDrain()
        {
            using (var context = new OwnedEditorDispatcher())
            {
                bool nested = false;
                context.Post(async state =>
                {
                    await Task.Delay(20);
                    Assert.AreSame(context, SynchronizationContext.Current);
                    context.Post(async child =>
                    {
                        await Task.Delay(40);
                        Assert.AreSame(context, SynchronizationContext.Current);
                        nested = true;
                    }, null);
                }, null);
                context.Drain();
                Assert.IsTrue(nested);
            }
        }
    }
}
