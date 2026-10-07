using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Linq;
using System.IO;
using System.Web.Script.Serialization;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed partial class ImportedFormReadbackTests
    {
        [DataTestMethod, DataRow(false), DataRow(true)]
        public void ExactInitialSnapshotNeverQueuesOrRecaptures(bool nativeForms)
        {
            var target = Snapshot(FormStreamPaddingTests.ContainerResourceBefore());
            WithQueue(queue =>
            {
                var task = ImportedFormReadback.VerifyAsync(target, target, nativeForms,
                    () => Assert.Fail("Unexpected guard"), () => throw new AssertFailedException("Unexpected readback"), QueuePulse);
                Assert.IsTrue(task.GetAwaiter().GetResult()); Assert.AreEqual(0, queue.Count);
            });
        }

        [TestMethod]
        public void UnsettledNativeImportIsNeverRecapturedByTheReadbackGate()
        {
            var target = Snapshot(FormStreamPaddingTests.ContainerResourceBefore());
            var actual = Snapshot(UserFormQualificationFontsTests.ResourceWithChangedDeclaredFont("Frame", 8.25m));
            WithQueue(queue =>
            {
                Assert.IsFalse(ImportedFormReadback.VerifyAsync(target, actual, false,
                    () => Assert.Fail("Unexpected guard"), () => throw new AssertFailedException("Unexpected readback"), QueuePulse).GetAwaiter().GetResult());
                Assert.AreEqual(0, queue.Count);
            });
        }

        [TestMethod]
        public void Frame827PersistenceMustBecomeExactlyEqualOnTheSingleOwnerContinuation()
        {
            var target = Snapshot(FormStreamPaddingTests.ContainerResourceBefore());
            var stale = Snapshot(UserFormQualificationFontsTests.ResourceWithChangedDeclaredFont("Frame", 8.25m));
            Assert.IsFalse(stale.SameAs(target), "The first rounded font must stay a strict discrepancy.");
            int thread = Thread.CurrentThread.ManagedThreadId, guards = 0, captures = 0;
            WithQueue(queue =>
            {
                var task = ImportedFormReadback.VerifyAsync(target, stale, true,
                    () => { Assert.AreEqual(thread, Thread.CurrentThread.ManagedThreadId); guards++; },
                    () => { Assert.AreEqual(1, guards); captures++; return target; }, QueuePulse);
                Assert.IsFalse(task.IsCompleted); Assert.AreEqual(1, queue.Count); Assert.AreEqual(0, captures);
                queue.RunOne(); Assert.IsTrue(task.GetAwaiter().GetResult());
                Assert.AreEqual(1, guards); Assert.AreEqual(1, captures); Assert.AreEqual(0, queue.Count);
            });
        }

        [DataTestMethod, DataRow("font"), DataRow("resource"), DataRow("source"), DataRow("references")]
        public void PersistentDiscrepanciesRemainRefusedAfterOneReadback(string discrepancy)
        {
            var target = Snapshot(FormStreamPaddingTests.ContainerResourceBefore());
            var stale = Snapshot(UserFormQualificationFontsTests.ResourceWithChangedDeclaredFont("Frame", 8.25m));
            var actual = stale;
            if (discrepancy == "resource")
            {
                var resource = new byte[target.Files["Form1.frx"].Length + 1];
                Array.Copy(target.Files["Form1.frx"], resource, resource.Length - 1); resource[resource.Length - 1] = 71;
                actual = Snapshot(resource);
            }
            if (discrepancy == "source") actual = Snapshot(target.Files["Form1.frx"], "' changed\n");
            if (discrepancy == "references") actual = Snapshot(target.Files["Form1.frx"], references: "changed");
            int captures = 0;
            WithQueue(queue =>
            {
                var task = ImportedFormReadback.VerifyAsync(target, stale, true, () => { }, () => { captures++; return actual; }, QueuePulse);
                queue.RunOne(); Assert.IsFalse(task.GetAwaiter().GetResult());
                Assert.AreEqual(1, captures); Assert.AreEqual(0, queue.Count);
            });
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void FailedOwnerGuardOrExportPreservesTheErrorAndNeverRequeues(bool export)
        {
            var target = Snapshot(FormStreamPaddingTests.ContainerResourceBefore());
            var stale = Snapshot(UserFormQualificationFontsTests.ResourceWithChangedDeclaredFont("Frame", 8.25m));
            var failure = new InvalidOperationException(export ? "Native export failed" : "Owner or recovery changed");
            int guards = 0, captures = 0;
            WithQueue(queue =>
            {
                var task = ImportedFormReadback.VerifyAsync(target, stale, true,
                    () => { guards++; if (!export) throw failure; },
                    () => { captures++; throw failure; }, QueuePulse);
                queue.RunOne(); var thrown = Assert.ThrowsException<InvalidOperationException>(() => task.GetAwaiter().GetResult());
                Assert.AreSame(failure, thrown); Assert.AreEqual(1, guards); Assert.AreEqual(export ? 1 : 0, captures);
                Assert.AreEqual(0, queue.Count);
            });
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void OptionalTraceCannotChangeGuardOrderOrReplaceItsOriginalFailure(bool brokenWriter)
        {
            var target = Snapshot(FormStreamPaddingTests.ContainerResourceBefore());
            var stale = Snapshot(UserFormQualificationFontsTests.ResourceWithChangedDeclaredFont("Frame", 8.25m));
            var rows = new List<string>();
            var trace = new VbeInspectionTrace(row =>
            {
                rows.Add(row);
                if (brokenWriter) throw new IOException("SECRET_TRACE_PATH");
            });
            var failure = new InvalidOperationException("SECRET_GUARD_FAILURE");
            int guards = 0, captures = 0;
            using (trace.Enter())
                WithQueue(queue =>
                {
                    var task = ImportedFormReadback.VerifyAsync(target, stale, true,
                        () => { guards++; throw failure; }, () => { captures++; return target; }, QueuePulse);
                    CollectionAssert.AreEqual(new[] { "ImportReadbackYieldBefore" }, Phases(rows));
                    Assert.IsFalse(task.IsCompleted); Assert.AreEqual(1, queue.Count);
                    queue.RunOne();
                    Assert.AreSame(failure, Assert.ThrowsException<InvalidOperationException>(() => task.GetAwaiter().GetResult()));
                    Assert.AreEqual(0, queue.Count);
                });
            Assert.AreEqual(1, guards); Assert.AreEqual(0, captures);
            CollectionAssert.AreEqual(new[] { "ImportReadbackYieldBefore", "ImportReadbackYieldReturned", "ImportReadbackGuardBefore" }, Phases(rows));
            Assert.IsFalse(string.Join("", rows).Contains("SECRET_"));
        }

        [TestMethod]
        public void ReadbackTraceSharesItsBoundedBudgetAndPreservesOneGuardAndCapturePerAttempt()
        {
            var target = Snapshot(FormStreamPaddingTests.ContainerResourceBefore());
            var stale = Snapshot(UserFormQualificationFontsTests.ResourceWithChangedDeclaredFont("Frame", 8.25m));
            var rows = new List<string>(); var trace = new VbeInspectionTrace(rows.Add);
            int guards = 0, captures = 0;
            using (trace.Enter())
                WithQueue(queue =>
                {
                    for (int attempt = 0; attempt < 50; attempt++)
                    {
                        bool guarded = false;
                        var task = ImportedFormReadback.VerifyAsync(target, stale, true,
                            () => { guarded = true; guards++; },
                            () => { Assert.IsTrue(guarded); captures++; return target; }, QueuePulse);
                        queue.RunOne(); Assert.IsTrue(task.GetAwaiter().GetResult()); Assert.AreEqual(0, queue.Count);
                    }
                });
            Assert.AreEqual(50, guards); Assert.AreEqual(50, captures);
            Assert.AreEqual(VbeInspectionTrace.MaximumEvents, rows.Count);
            CollectionAssert.AreEqual(new[] { "ImportReadbackYieldBefore", "ImportReadbackYieldReturned",
                "ImportReadbackGuardBefore", "ImportReadbackGuardReturned", "ImportReadbackCaptureBefore",
                "ImportReadbackCaptureReturned" }, Phases(rows.Take(6)));
        }

        private static string[] Phases(IEnumerable<string> rows)
        {
            return rows.Select(row => (string)new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(row)["Phase"]).ToArray();
        }

        private static async Task QueuePulse() { await Task.Yield(); }

        private static VbaGitSnapshot Snapshot(byte[] resource, string source = "", string references = "")
        {
            return new VbaGitSnapshot(new VbaGitManifest
            {
                References = references,
                Components = new[] { new VbaGitComponent { Name = "Form1", Type = 3, HasResources = true } }
            }, new Dictionary<string, byte[]>
            {
                ["Form1.frm"] = VbaGitSnapshot.Utf8.GetBytes("VERSION 5.00\nBegin SyntheticForm\n OleObjectBlob = \"Form1.frx\":0000\nEnd\nAttribute VB_Name = \"Form1\"\n" + source),
                ["Form1.frx"] = (byte[])resource.Clone()
            });
        }

        private static void WithQueue(Action<OwnerQueue> action)
        {
            var previous = SynchronizationContext.Current; var queue = new OwnerQueue();
            try { SynchronizationContext.SetSynchronizationContext(queue); action(queue); }
            finally { SynchronizationContext.SetSynchronizationContext(previous); }
        }

        private sealed class OwnerQueue : SynchronizationContext
        {
            private readonly Queue<Action> work = new Queue<Action>();
            internal int Count => work.Count;
            public override void Post(SendOrPostCallback callback, object state) { work.Enqueue(() => callback(state)); }
            internal void RunOne() { work.Dequeue()(); }
        }
    }
}
