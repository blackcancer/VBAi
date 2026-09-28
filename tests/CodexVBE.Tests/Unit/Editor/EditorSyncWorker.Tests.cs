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
    }
}
