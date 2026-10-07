using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class WordChatTopWindowInventoryTests
    {
        private const int ProcessId = 51256;
        private const uint ThreadId = 59236;

        [TestMethod]
        public void HiddenAndOtherThreadWordWindowsDoNotConsumeVisibleTargetBound()
        {
            var identities = new Dictionary<IntPtr, WordChatTopWindowInventory.Identity>();
            for (int index = 1; index <= 100; index++)
                identities.Add(new IntPtr(index), Id(ProcessId, ThreadId, false));
            for (int index = 101; index <= 200; index++)
                identities.Add(new IntPtr(index), Id(ProcessId, ThreadId + 1, true));
            for (int index = 201; index <= 300; index++)
                identities.Add(new IntPtr(index), Id(ProcessId + 1, ThreadId, true));
            IntPtr target = new IntPtr(301);
            identities.Add(target, Id(ProcessId, ThreadId, true));
            WordChatTopWindowInventory.Receipt refusal = null;
            var selected = WordChatTopWindowInventory.Read(Enumerate(identities.Keys),
                window => identities[window], ProcessId, ThreadId, row => refusal = row);
            CollectionAssert.AreEqual(new[] { target }, selected);
            Assert.IsNull(refusal);
        }

        [TestMethod]
        public void TooManyVisibleTargetsRecordsCountsBeforeRefusal()
        {
            var windows = new List<IntPtr>();
            for (int index = 1; index <= 65; index++) windows.Add(new IntPtr(index));
            WordChatTopWindowInventory.Receipt refusal = null;
            Assert.ThrowsException<InvalidOperationException>(() => WordChatTopWindowInventory.Read(
                Enumerate(windows), window => Id(ProcessId, ThreadId, true), ProcessId, ThreadId,
                row => refusal = row));
            Assert.IsNotNull(refusal);
            Assert.IsTrue(refusal.ApiReturned);
            Assert.AreEqual(65, refusal.VisitedTotal);
            Assert.AreEqual(65, refusal.OwnedProcessCount);
            Assert.AreEqual(65, refusal.ExactThreadVisibleCount);
            Assert.IsFalse(refusal.GlobalBoundHit);
            Assert.AreEqual("TargetBound", refusal.FailureStatus);
        }

        [TestMethod]
        public void GlobalCallbackBoundAndActualApiFailureHaveDistinctRecordedStatuses()
        {
            var windows = new List<IntPtr>();
            for (int index = 1; index <= 4097; index++) windows.Add(new IntPtr(index));
            WordChatTopWindowInventory.Receipt refusal = null;
            Assert.ThrowsException<InvalidOperationException>(() => WordChatTopWindowInventory.Read(
                Enumerate(windows), window => Id(ProcessId + 1, ThreadId, false), ProcessId, ThreadId,
                row => refusal = row));
            Assert.IsNotNull(refusal);
            Assert.IsFalse(refusal.ApiReturned);
            Assert.IsTrue(refusal.GlobalBoundHit);
            Assert.AreEqual(4097, refusal.VisitedTotal);
            Assert.AreEqual(0, refusal.OwnedProcessCount);
            Assert.AreEqual(0, refusal.ExactThreadVisibleCount);
            Assert.AreEqual("GlobalBound", refusal.FailureStatus);

            refusal = null;
            Assert.ThrowsException<InvalidOperationException>(() => WordChatTopWindowInventory.Read(
                (visitor, state) => false, window => Id(ProcessId, ThreadId, true), ProcessId, ThreadId,
                row => refusal = row));
            Assert.IsNotNull(refusal);
            Assert.IsFalse(refusal.ApiReturned);
            Assert.IsFalse(refusal.GlobalBoundHit);
            Assert.AreEqual(0, refusal.VisitedTotal);
            Assert.AreEqual("ApiFailure", refusal.FailureStatus);
        }

        private static WordChatTopWindowInventory.Identity Id(int processId, uint threadId, bool visible)
            => new WordChatTopWindowInventory.Identity
            {
                ProcessId = processId,
                ThreadId = threadId,
                Visible = visible
            };

        private static WordChatTopWindowInventory.Enumerator Enumerate(IEnumerable<IntPtr> windows)
            => (visitor, state) =>
            {
                foreach (IntPtr window in windows)
                    if (!visitor(window, state)) return false;
                return true;
            };
    }
}
