using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Windows.Automation;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class PrivateDesktopUiActionTests
    {
        [TestMethod]
        public void NativeButtonRequiresExactLeafIdentityAndOwnerRelationship()
        {
            const int pid = 422;
            const uint tid = 73;
            string role = ControlType.Button.ProgrammaticName;
            Assert.IsTrue(PrivateDesktopUiAction.MatchesButton("compare", "compare", "Compare", "Compare",
                role, pid, pid, tid, pid, tid, 0x123, 0x123, true, false, true, true, true));
            Assert.IsFalse(PrivateDesktopUiAction.MatchesButton("other", "compare", "Compare", "Compare",
                role, pid, pid, tid, pid, tid, 0x123, 0x123, true, false, true, true, true));
            Assert.IsFalse(PrivateDesktopUiAction.MatchesButton("compare", "compare", "Other", "Compare",
                role, pid, pid, tid, pid, tid, 0x123, 0x123, true, false, true, true, true));
            Assert.IsFalse(PrivateDesktopUiAction.MatchesButton("compare", "compare", "Compare", "Compare",
                ControlType.MenuItem.ProgrammaticName, pid, pid, tid, pid, tid, 0x123, 0x123, true, false, true, true, true));
            Assert.IsFalse(PrivateDesktopUiAction.MatchesButton("compare", "compare", "Compare", "Compare",
                role, pid, pid + 1, tid, pid, tid, 0x123, 0x123, true, false, true, true, true));
            Assert.IsFalse(PrivateDesktopUiAction.MatchesButton("compare", "compare", "Compare", "Compare",
                role, pid, pid, tid + 1, pid, tid, 0x123, 0x123, true, false, true, true, true));
            Assert.IsFalse(PrivateDesktopUiAction.MatchesButton("compare", "compare", "Compare", "Compare",
                role, pid, pid, tid, pid, tid, 0, 0x123, true, false, true, true, true));
            Assert.IsFalse(PrivateDesktopUiAction.MatchesButton("compare", "compare", "Compare", "Compare",
                role, pid, pid, tid, pid, tid, 0x123, 0x124, true, false, true, true, true));
            Assert.IsFalse(PrivateDesktopUiAction.MatchesButton("compare", "compare", "Compare", "Compare",
                role, pid, pid, tid, pid, tid, 0x123, 0x123, true, false, true, true, false));
            Assert.IsFalse(PrivateDesktopUiAction.MatchesButton("compare", "compare", "Compare", "Compare",
                role, pid, pid, tid, pid, tid, 0x123, 0x123, false, false, true, true, true));
        }

        [TestMethod]
        public void VirtualGitItemRequiresExactPopupAndZeroLeafHandle()
        {
            const int pid = 422;
            const uint tid = 73;
            const string label = "GitHub · synchronize VBA…";
            string role = ControlType.MenuItem.ProgrammaticName;
            Assert.IsTrue(PrivateDesktopUiAction.MatchesVirtualGit(label, label, role,
                pid, pid, tid, pid, tid, 0, 0x222, 0x222, true, false, true));
            Assert.IsFalse(PrivateDesktopUiAction.MatchesVirtualGit("GitHub", label, role,
                pid, pid, tid, pid, tid, 0, 0x222, 0x222, true, false, true));
            Assert.IsFalse(PrivateDesktopUiAction.MatchesVirtualGit(label, label, ControlType.Button.ProgrammaticName,
                pid, pid, tid, pid, tid, 0, 0x222, 0x222, true, false, true));
            Assert.IsFalse(PrivateDesktopUiAction.MatchesVirtualGit(label, label, role,
                pid, pid + 1, tid, pid, tid, 0, 0x222, 0x222, true, false, true));
            Assert.IsFalse(PrivateDesktopUiAction.MatchesVirtualGit(label, label, role,
                pid, pid, tid + 1, pid, tid, 0, 0x222, 0x222, true, false, true));
            Assert.IsFalse(PrivateDesktopUiAction.MatchesVirtualGit(label, label, role,
                pid, pid, tid, pid, tid, 17, 0x222, 0x222, true, false, true));
            Assert.IsFalse(PrivateDesktopUiAction.MatchesVirtualGit(label, label, role,
                pid, pid, tid, pid, tid, 0, 0x333, 0x222, true, false, true));
            Assert.IsFalse(PrivateDesktopUiAction.MatchesVirtualGit(label, label, role,
                pid, pid, tid, pid, tid, 0, 0x222, 0x222, true, true, true));
            Assert.IsFalse(PrivateDesktopUiAction.MatchesVirtualGit(label, label, role,
                pid, pid, tid, pid, tid, 0, 0x222, 0x222, true, false, false));
        }
    }
}
