using System;
using System.Threading;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class EditorSyncWorkerTests
    {
        [TestMethod]
        public void WorkerOwnsPreparationAndPersistsImmutableSnapshotWithoutWritingVba()
        {
            using (var host = new EditorFixture())
            using (var worker = new EditorSyncWorker())
            {
                var store = new EditorDraftStore(host.Root); var doc = new EditorDocument(host);
                doc.Edit(doc.Text + "\n' first draft"); string first = doc.Text;
                var pending = worker.Prepare(doc, store); doc.Edit(first + "\n' newer typing");
                var plan = pending.GetAwaiter().GetResult();
                Assert.AreNotEqual(Thread.CurrentThread.ManagedThreadId, plan.WorkerThreadId);
                Assert.AreEqual(first, plan.After); Assert.AreEqual(first, store.Recover(doc.RecoveryKey).Text);
                Assert.AreEqual(0, host.Writes);
                Assert.ThrowsException<InvalidOperationException>(() => doc.Synchronize(plan));
                var current = worker.Prepare(doc, store).GetAwaiter().GetResult();
                Assert.AreEqual(plan.WorkerThreadId, current.WorkerThreadId);
                string result = doc.Synchronize(current); doc.Acknowledge(result, doc.Text);
                Assert.IsFalse(doc.Dirty); Assert.AreEqual(1, host.Writes);
            }
        }
        [TestMethod]
        public void WorkerPropagatesOwnedEvaluationDiskAndCompletedQueueFailures()
        {
            using (var host = new EditorFixture())
            {
                var worker = new EditorSyncWorker();
                Assert.ThrowsException<InvalidOperationException>(() => worker.Evaluate<int>(() => throw new InvalidOperationException("owned action refusal")).GetAwaiter().GetResult());
                Assert.AreEqual(42, worker.Evaluate(() => 42).GetAwaiter().GetResult());
                System.IO.File.WriteAllText(host.Root, "owned path occupies store root");
                var document = new EditorDocument(host); document.Edit(document.Text + "\n' owned draft");
                Assert.ThrowsException<System.IO.IOException>(() => worker.Prepare(document, new EditorDraftStore(host.Root)).GetAwaiter().GetResult());
                System.IO.File.Delete(host.Root);
                worker.Dispose();
                var thread = (Thread)typeof(EditorSyncWorker).GetField("thread", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).GetValue(worker);
                Assert.IsTrue(thread.Join(5000), "Owned worker must finish its queue and release the collection.");
                Assert.ThrowsException<ObjectDisposedException>(() => worker.Evaluate(() => 1).GetAwaiter().GetResult());
                Assert.ThrowsException<ObjectDisposedException>(() => worker.Prepare(document, new EditorDraftStore(host.Root)).GetAwaiter().GetResult());
                worker.Dispose();
            }
        }
    }
}
