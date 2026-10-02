using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Desktop.Helper;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class NativeToolStripPopupIdentityTests
    {
        [TestMethod]
        public void LegacyToolbarRoleRequiresTheExactWinFormsDropdownClass()
        {
            const string dropdown = "WindowsForms10.Window.20808.app.0.123_r8_ad1";
            Assert.IsTrue(NativeToolStripPopupIdentity.Matches("ControlType.Menu", dropdown));
            Assert.IsTrue(NativeToolStripPopupIdentity.Matches("ControlType.ToolBar", dropdown));
            foreach (string role in new[] { null, "", "ControlType.Pane", "ControlType.MenuItem", "controltype.menu" })
                Assert.IsFalse(NativeToolStripPopupIdentity.Matches(role, dropdown));
            foreach (string cls in new[] { null, "", "ToolbarWindow32", "WindowsForms10.Window.0.app.0.123_r8_ad1",
                "WindowsForms10.Window.20808", "WindowsForms10.Window.208080.app.0.123_r8_ad1" })
                Assert.IsFalse(NativeToolStripPopupIdentity.Matches("ControlType.ToolBar", cls));
            Assert.IsFalse(NativeToolStripPopupIdentity.Matches("ControlType.Menu", "#32768"));
        }
    }
}
