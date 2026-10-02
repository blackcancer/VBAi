using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Desktop.Helper;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class NativeComboListDiscoveryTests
    {
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
