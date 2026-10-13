using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Threading;
using System.Threading.Tasks;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks owner dispatch without a result, including asynchronous failure and cancellation.</summary>
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class VbeUiTaskAsyncOwnerTests
    {
        [STATestMethod, DataRow(false, "success"), DataRow(true, "success")]
        [DataRow(false, "failure"), DataRow(true, "failure")]
        [DataRow(false, "cancellation"), DataRow(true, "cancellation")]
        public void VoidOwnerOperationPreservesThreadContextAndTerminalOutcome(bool baseContext, string outcome)
        {
            var original = SynchronizationContext.Current;
            var ambient = baseContext ? new SynchronizationContext() : null;
            int owner = Thread.CurrentThread.ManagedThreadId, calls = 0;
            var failure = new InvalidOperationException("owned asynchronous failure");
            using (var cancellation = new CancellationTokenSource())
            {
                try
                {
                    SynchronizationContext.SetSynchronizationContext(ambient);
                    var pending = VbeUiTask.RunAsync(async () =>
                    {
                        calls++;
                        Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                        await Task.Delay(2);
                        Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                        Assert.AreEqual(ApartmentState.STA, Thread.CurrentThread.GetApartmentState());
                        if (outcome == "failure") throw failure;
                        if (outcome == "cancellation") { cancellation.Cancel(); cancellation.Token.ThrowIfCancellationRequested(); }
                    });
                    Assert.AreSame(ambient, SynchronizationContext.Current);
                    if (outcome == "failure") Assert.AreSame(failure, Assert.ThrowsException<InvalidOperationException>(() => ModernEditorDebugFixture.Wait(pending)));
                    else if (outcome == "cancellation")
                    {
                        try { ModernEditorDebugFixture.Wait(pending); Assert.Fail("Cancellation must be preserved."); }
                        catch (OperationCanceledException error) { Assert.AreEqual(cancellation.Token, error.CancellationToken); }
                        Assert.IsTrue(pending.IsCanceled);
                    }
                    else ModernEditorDebugFixture.Wait(pending);
                    Assert.AreEqual(1, calls);
                    // DoEvents may install a default context when the ambient context was null.
                    if (baseContext) Assert.AreSame(ambient, SynchronizationContext.Current);
                }
                finally { SynchronizationContext.SetSynchronizationContext(original); }
            }
        }

        [TestMethod]
        public void VoidOwnerOperationRejectsMissingOperationAndForeignApartmentBeforeCallingWork()
        {
            Assert.ThrowsException<ArgumentNullException>(() => VbeUiTask.RunAsync(null));
            Exception observed = null;
            int calls = 0;
            var thread = new Thread(() =>
            {
                try { VbeUiTask.RunAsync(() => { calls++; return Task.CompletedTask; }); }
                catch (Exception error) { observed = error; }
            });
            thread.SetApartmentState(ApartmentState.MTA);
            thread.Start(); Assert.IsTrue(thread.Join(5000));
            Assert.IsInstanceOfType(observed, typeof(InvalidOperationException));
            Assert.AreEqual(0, calls);
        }
    }
}
