using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Windows.Automation;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class OfficeUiSnapshotTests
    {
        [TestMethod]
        public void PropertyRequestDoesNotCacheDescendantsOrDisableActionPatterns()
        {
            var request = OfficeUiSnapshot.CreateRequest();
            Assert.AreEqual(TreeScope.Element, request.TreeScope);
            Assert.AreEqual(AutomationElementMode.Full, request.AutomationElementMode);
            Assert.AreSame(Automation.RawViewCondition, request.TreeFilter);
        }

        [TestMethod]
        public void ReadRequiresTheOwnerBeforeAndAfterTheSingleFetch()
        {
            var order = new List<string>();
            int result = OfficeUiSnapshot.ReadOnce(() => order.Add("owner"), () => { order.Add("read"); return 42; });
            Assert.AreEqual(42, result);
            CollectionAssert.AreEqual(new[] { "owner", "read", "owner" }, order);
        }

        [DataTestMethod][DataRow(1)][DataRow(2)]
        public void OwnerRefusalBeforeOrAfterTheFetchCannotReturnAUsableSnapshot(int refusedGuard)
        {
            int guards = 0, reads = 0;
            var refusal = new InvalidOperationException("Original owner ended.");
            Assert.AreSame(refusal, Assert.ThrowsException<InvalidOperationException>(() =>
                OfficeUiSnapshot.ReadOnce(() => { if (++guards == refusedGuard) throw refusal; },
                    () => { reads++; return 42; })));
            Assert.AreEqual(refusedGuard == 1 ? 0 : 1, reads);
        }

        [TestMethod]
        public void FailedFetchIsPreservedAndNeverRepeated()
        {
            int reads = 0;
            var failure = new TimeoutException("UIA read did not finish.");
            Assert.AreSame(failure, Assert.ThrowsException<TimeoutException>(() =>
                OfficeUiSnapshot.ReadOnce<int>(null, () => { reads++; throw failure; })));
            Assert.AreEqual(1, reads);
        }

        [TestMethod]
        public void EachObservationFetchesFreshValues()
        {
            int provider = 0;
            Assert.AreEqual(1, OfficeUiSnapshot.ReadOnce(null, () => ++provider));
            Assert.AreEqual(2, OfficeUiSnapshot.ReadOnce(null, () => ++provider));
        }

        [TestMethod]
        public void InvalidInputsFailBeforeAnyOwnerOrProviderCall()
        {
            Assert.ThrowsException<ArgumentNullException>(() => OfficeUiSnapshot.ReadOnce<int>(() => Assert.Fail(), null));
            Assert.ThrowsException<ArgumentNullException>(() => OfficeUiSnapshot.Refresh(null, () => Assert.Fail()));
            Assert.ThrowsException<ArgumentNullException>(() => OfficeUiSnapshot.RefreshIdentity(null, () => Assert.Fail()));
            Assert.ThrowsException<ArgumentNullException>(() => OfficeUiSnapshot.FindAll(null, TreeScope.Element, Condition.TrueCondition, () => Assert.Fail()));
        }
    }
}
