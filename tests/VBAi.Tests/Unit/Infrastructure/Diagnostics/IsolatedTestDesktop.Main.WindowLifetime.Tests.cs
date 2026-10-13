using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class IsolatedTestDesktopMainWindowLifetimeTests
    {
        private static IsolatedTestDesktop.MainWindowIdentity Identity(uint pid = 42, uint thread = 7) =>
            new IsolatedTestDesktop.MainWindowIdentity { ProcessId = pid, ThreadId = thread };

        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void OwnedWindowRequiresStableIdentityAfterClassAndVisibilityReads(bool visible)
        {
            var log = new List<string>(); var handle = new IntPtr(11);
            var result = IsolatedTestDesktop.ReadMainWindow(42, handle,
                window => { Assert.AreEqual(handle, window); log.Add("identity"); return Identity(); },
                window => { Assert.AreEqual(handle, window); log.Add("class"); return "OpusApp"; },
                window => { Assert.AreEqual(handle, window); log.Add("visible"); return visible; });
            CollectionAssert.AreEqual(new[] { "identity", "class", "visible", "identity" }, log);
            Assert.AreEqual(11L, result.Handle); Assert.AreEqual(42u, result.ProcessId);
            Assert.AreEqual(7u, result.ThreadId); Assert.AreEqual("OpusApp", result.ClassName);
            Assert.AreEqual(visible, result.Visible);
        }

        [TestMethod]
        public void ForeignWindowIsSkippedWithoutReadingItsClassOrVisibility()
        {
            int identities = 0;
            var row = IsolatedTestDesktop.ReadMainWindow(42, new IntPtr(11),
                window => { identities++; return Identity(43); },
                window => { Assert.Fail("Foreign class must remain unread."); return null; },
                window => { Assert.Fail("Foreign visibility must remain unread."); return false; });
            Assert.IsNull(row); Assert.AreEqual(1, identities);
        }

        [DataTestMethod]
        [DataRow(0)]
        [DataRow(1)]
        [DataRow(2)]
        [DataRow(3)]
        [DataRow(4)]
        [DataRow(5)]
        [DataRow(6)]
        [DataRow(7)]
        [DataRow(8)]
        [DataRow(9)]
        [DataRow(10)]
        public void MissingOrReusedWindowIdentityNeverSuppliesAnOwnedRow(int fault)
        {
            int reads = 0, classes = 0;
            Assert.ThrowsException<InvalidOperationException>(() => IsolatedTestDesktop.ReadMainWindow(
                fault == 0 ? 0u : 42u, new IntPtr(fault == 1 ? 0 : 11), window =>
                {
                    reads++;
                    if (reads == 1) return fault == 2 ? null : Identity(fault == 3 ? 0u : 42u, fault == 4 ? 0u : 7u);
                    return fault == 5 ? null : Identity(fault == 6 ? 43u : fault == 7 ? 0u : 42u,
                        fault == 8 ? 8u : fault == 9 ? 0u : 7u);
                }, window => { classes++; return fault == 10 ? "" : "OpusApp"; }, window => true));
            Assert.AreEqual(fault < 2 ? 0 : fault < 5 ? 1 : 2, reads);
            Assert.AreEqual(fault < 5 ? 0 : 1, classes);
        }

        [DataTestMethod]
        [DataRow(0)]
        [DataRow(1)]
        [DataRow(2)]
        [DataRow(3)]
        public void DestroyedWindowAtAnyNativeReadRejectsTheWholeSnapshotBeforeOneFreshRead(int stage)
        {
            int enumeration = 0, identities = 0, placement = 0;
            IsolatedTestDesktop.MainInventory interrupted = null;
            var result = IsolatedTestDesktop.ReadCompleteMainInventory(42, () =>
            {
                enumeration++;
                if (enumeration == 2) return IsolatedTestDesktopMainTests.Inventory();
                interrupted = IsolatedTestDesktop.ReadMainWindowsOnce(42, visit => visit(new IntPtr(11)),
                    handle => IsolatedTestDesktop.ReadMainWindow(42, handle, window =>
                    {
                        identities++;
                        if (stage == 0 || (stage == 3 && identities == 2)) throw new Win32Exception(1400);
                        return Identity();
                    }, window => { if (stage == 1) throw new Win32Exception(1400); return "OpusApp"; },
                    window => { if (stage == 2) throw new Win32Exception(1400); return true; }),
                    () => 5, () => placement++);
                return interrupted;
            }, () => placement++);
            Assert.AreEqual(2, enumeration); Assert.AreEqual(6, placement);
            Assert.IsFalse(interrupted.Complete); Assert.AreEqual(1400, interrupted.EnumerationError);
            Assert.IsTrue(interrupted.WindowDestroyedDuringRead); Assert.AreEqual(1, interrupted.Visited);
            Assert.AreEqual(0, interrupted.Windows.Length); Assert.AreEqual(0, interrupted.ReadAttempts);
            Assert.AreEqual(2, result.ReadAttempts); Assert.AreEqual(1, result.RejectedObservations.Length);
            Assert.IsTrue(result.RejectedObservations[0].WindowDestroyedDuringRead);
            Assert.AreEqual(1400, result.RejectedObservations[0].EnumerationError);
        }

        [DataTestMethod]
        [DataRow(0)]
        [DataRow(5)]
        [DataRow(6)]
        [DataRow(995)]
        public void OtherCallbackErrorsPreserveTheOriginalExceptionWithoutAnotherEnumeration(int nativeError)
        {
            var error = new Win32Exception(nativeError); int reads = 0;
            var actual = Assert.ThrowsException<Win32Exception>(() =>
                IsolatedTestDesktop.ReadCompleteMainInventory(42, () =>
                {
                    reads++;
                    return IsolatedTestDesktop.ReadMainWindowsOnce(42, visit => visit(new IntPtr(11)),
                        window => { throw error; }, () => 1400, () => { });
                }, () => { }));
            Assert.AreSame(error, actual); Assert.AreEqual(1, reads);
        }

        [TestMethod]
        public void RetainedPartialRowsNeverAuthorizeARequiredRootAbsentFromTheFreshSnapshot()
        {
            int reads = 0;
            var result = IsolatedTestDesktop.ReadCompleteMainInventory(42, () =>
            {
                if (++reads == 2) return new IsolatedTestDesktop.MainInventory
                { Desktop = "Default", Complete = true, Windows = new IsolatedTestDesktop.MainWindow[0] };
                return IsolatedTestDesktop.ReadMainWindowsOnce(42, visit =>
                    visit(new IntPtr(11)) && visit(new IntPtr(12)), window =>
                    {
                        if (window.ToInt64() == 12) throw new Win32Exception(1400);
                        return IsolatedTestDesktopMainTests.Inventory().Windows[0];
                    }, () => 0, () => { });
            }, () => { });
            Assert.AreEqual(2, reads); Assert.AreEqual(1, result.RejectedObservations[0].OwnedWindows);
            Assert.AreEqual(2, result.RejectedObservations[0].Visited); Assert.AreEqual(0, result.Windows.Length);
            Assert.ThrowsException<InvalidOperationException>(() =>
                IsolatedTestDesktop.RequireMainInventory(result, 42, new IntPtr(11), false, false));
        }

        [TestMethod]
        public void RepeatedDestroyedWindowStopsAfterTwoWholeReadsAndNeverAdmitsPartialRows()
        {
            int reads = 0;
            var error = Assert.ThrowsException<InvalidOperationException>(() =>
                IsolatedTestDesktop.ReadCompleteMainInventory(42, () =>
                {
                    reads++;
                    return IsolatedTestDesktop.ReadMainWindowsOnce(42, visit => visit(new IntPtr(11)),
                        window => { throw new Win32Exception(1400); }, () => 0, () => { });
                }, () => { }));
            Assert.AreEqual(2, reads); StringAssert.Contains(error.Message, "Win32=1400");
        }

        [DataTestMethod]
        [DataRow(true, 1400)]
        [DataRow(false, 0)]
        [DataRow(false, 5)]
        public void DestroyedWindowMarkerMustDescribeAnIncompleteNativeInvalidHandleRead(bool complete, int error)
        {
            int reads = 0; var snapshot = IsolatedTestDesktopMainTests.Inventory();
            snapshot.WindowDestroyedDuringRead = true; snapshot.Complete = complete; snapshot.EnumerationError = error;
            Assert.ThrowsException<InvalidOperationException>(() =>
                IsolatedTestDesktop.ReadCompleteMainInventory(42, () => { reads++; return snapshot; }, () => { }));
            Assert.AreEqual(1, reads);
        }

        [TestMethod]
        public void NativeEnumerationInvalidHandleWithoutAWindowReadFailureDoesNotRequestAnotherRead()
        {
            int reads = 0;
            Assert.ThrowsException<InvalidOperationException>(() =>
                IsolatedTestDesktop.ReadCompleteMainInventory(42, () =>
                {
                    reads++;
                    return IsolatedTestDesktop.ReadMainWindowsOnce(42, visit => false,
                        window => { Assert.Fail(); return null; }, () => 1400, () => { });
                }, () => { }));
            Assert.AreEqual(1, reads);
        }

        [TestMethod]
        public void EnumerationBoundFailureIsPreservedWithoutFreshReadOrAnExtraWindowGetter()
        {
            int getters = 0, reads = 0;
            var error = Assert.ThrowsException<InvalidOperationException>(() =>
                IsolatedTestDesktop.ReadCompleteMainInventory(42, () =>
                {
                    reads++;
                    return IsolatedTestDesktop.ReadMainWindowsOnce(42, visit =>
                    { for (int i = 1; i <= 8193; i++) if (!visit(new IntPtr(i))) return false; return true; },
                    window => { getters++; return null; }, () => 1400, () => { });
                }, () => { }));
            Assert.AreEqual(1, reads); Assert.AreEqual(8192, getters); StringAssert.Contains(error.Message, "bound");
        }

        [TestMethod]
        public void PlacementLossAfterDestroyedWindowPreventsTheFreshRead()
        {
            int reads = 0, guards = 0; var lost = new InvalidOperationException("input desktop changed");
            var error = Assert.ThrowsException<InvalidOperationException>(() =>
                IsolatedTestDesktop.ReadCompleteMainInventory(42, () =>
                {
                    reads++;
                    return IsolatedTestDesktop.ReadMainWindowsOnce(42, visit => visit(new IntPtr(11)),
                        window => { throw new Win32Exception(1400); }, () => 0, () => { });
                }, () => { if (++guards == 2) throw lost; }));
            Assert.AreSame(lost, error); Assert.AreEqual(1, reads); Assert.AreEqual(2, guards);
        }
    }
}
