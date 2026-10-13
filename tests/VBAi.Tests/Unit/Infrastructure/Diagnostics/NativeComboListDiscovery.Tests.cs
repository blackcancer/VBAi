using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using VBAi.Desktop.Helper;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class NativeComboListDiscoveryTests
    {
        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void OriginalOwnedPopupMustBecomeVisibleBeforeAnItemCanBeSelected(bool visible)
        {
            Assert.AreEqual(visible ? NativeComboListDiscovery.NativeListReadiness.Ready : NativeComboListDiscovery.NativeListReadiness.Waiting,
                NativeComboListDiscovery.AssessNativeList(500, 60, 500, 60, 500, 60, "ComboLBox", visible, true, true));
            Assert.AreEqual(visible, NativeComboListDiscovery.MatchesNativeList(500, 60, 500, 60, 500, 60, "ComboLBox", visible, true, true));
        }

        [DataTestMethod]
        [DataRow("comboPid")]
        [DataRow("comboThread")]
        [DataRow("listPid")]
        [DataRow("listThread")]
        [DataRow("expectedPid")]
        [DataRow("expectedThread")]
        [DataRow("class")]
        [DataRow("comboDesktop")]
        [DataRow("listDesktop")]
        public void AHiddenForeignPopupIsRejectedRatherThanWaitedFor(string changed)
        {
            foreach (bool visible in new[] { false, true })
                Assert.AreEqual(NativeComboListDiscovery.NativeListReadiness.Rejected,
                    NativeComboListDiscovery.AssessNativeList(changed == "comboPid" ? 501 : 500,
                        changed == "comboThread" ? 61u : 60u, changed == "listPid" ? 501 : 500,
                        changed == "listThread" ? 61u : 60u, changed == "expectedPid" ? 0 : 500,
                        changed == "expectedThread" ? 0u : 60u, changed == "class" ? "Button" : "ComboLBox",
                        visible, changed != "comboDesktop", changed != "listDesktop"));
        }

        [TestMethod]
        public void OpeningObservationCannotRebindTheOriginalPopup()
        {
            NativeComboListDiscovery.RequireOriginalNativeList(new IntPtr(42), new IntPtr(42));
            Assert.ThrowsException<InvalidOperationException>(() => NativeComboListDiscovery.RequireOriginalNativeList(new IntPtr(42), new IntPtr(43)));
            Assert.ThrowsException<InvalidOperationException>(() => NativeComboListDiscovery.RequireOriginalNativeList(new IntPtr(42), IntPtr.Zero));
            Assert.ThrowsException<InvalidOperationException>(() => NativeComboListDiscovery.RequireOriginalNativeList(IntPtr.Zero, IntPtr.Zero));
        }

        [TestMethod]
        public void NativeDropdownRequiresExactProcessThreadDesktopVisibilityAndListClass()
        {
            const int pid = 500;
            const uint tid = 60;
            Assert.IsTrue(NativeComboListDiscovery.MatchesNativeList(pid, tid, pid, tid,
                pid, tid, "ComboLBox", true, true, true));
            Assert.IsTrue(NativeComboListDiscovery.MatchesNativeList(pid, tid, pid, tid,
                pid, tid, "WindowsForms10.LISTBOX.app.0", true, true, true));
            Assert.IsFalse(NativeComboListDiscovery.MatchesNativeList(pid + 1, tid, pid, tid,
                pid, tid, "ComboLBox", true, true, true));
            Assert.IsFalse(NativeComboListDiscovery.MatchesNativeList(pid, tid, pid + 1, tid,
                pid, tid, "ComboLBox", true, true, true));
            Assert.IsFalse(NativeComboListDiscovery.MatchesNativeList(pid, tid + 1, pid, tid,
                pid, tid, "ComboLBox", true, true, true));
            Assert.IsFalse(NativeComboListDiscovery.MatchesNativeList(pid, tid, pid, tid + 1,
                pid, tid, "ComboLBox", true, true, true));
            Assert.IsFalse(NativeComboListDiscovery.MatchesNativeList(pid, tid, pid, tid,
                pid, tid, "Button", true, true, true));
            Assert.IsFalse(NativeComboListDiscovery.MatchesNativeList(pid, tid, pid, tid,
                pid, tid, "ComboLBox", false, true, true));
            Assert.IsFalse(NativeComboListDiscovery.MatchesNativeList(pid, tid, pid, tid,
                pid, tid, "ComboLBox", true, false, true));
            Assert.IsFalse(NativeComboListDiscovery.MatchesNativeList(pid, tid, pid, tid,
                pid, tid, "ComboLBox", true, true, false));
        }
    }
}
