using System.Dynamic;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbeNativeNavigationTests
    {
        private static dynamic Position()
        {
            dynamic window = new ExpandoObject(); window.Type = 0; window.Caption = "Module (Code)";
            dynamic position = new ExpandoObject(); position.ActiveWindow = window; position.ObjectBrowserVisible = false;
            position.CodePaneAvailable = true; position.Project = "project"; position.Module = "module";
            position.StartLine = 2; position.EndLine = 2; position.StartColumn = 5; position.EndColumn = 8;
            return position;
        }
        [TestMethod]
        public void CaptionAndMainFrameChangesAloneAreNotNavigation()
        {
            dynamic before = Position(), after = Position();
            after.ActiveWindow.Caption = "Workbook - Module (Code)";
            Assert.IsFalse(VbeDebug.NavigationChanged(before, after));
            after.ActiveWindow.Type = 12;
            Assert.IsFalse(VbeDebug.NavigationChanged(before, after));
        }
        [TestMethod]
        public void BrowserAppearanceIsObservedWithoutClaimingItsSelectedMember()
        {
            dynamic before = Position(), after = Position();
            after.ObjectBrowserVisible = true; after.ActiveWindow.Type = 12;
            Assert.IsTrue(VbeDebug.NavigationChanged(before, after));
            before.ObjectBrowserVisible = true;
            Assert.IsFalse(VbeDebug.NavigationChanged(before, after));
        }
        [TestMethod]
        public void SourcePositionAndModuleChangesAreObserved()
        {
            dynamic before = Position(), after = Position();
            after.Module = "target"; Assert.IsTrue(VbeDebug.NavigationChanged(before, after));
            after.Module = before.Module; after.StartLine = 4;
            Assert.IsTrue(VbeDebug.NavigationChanged(before, after));
        }
    }
}
