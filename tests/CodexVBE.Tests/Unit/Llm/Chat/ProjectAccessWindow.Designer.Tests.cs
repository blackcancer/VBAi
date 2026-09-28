using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    [TestClass,TestCategory("Unit")]
    public sealed class ProjectAccessDesignerTests
    {
        [STATestMethod]
        public void DialogDesignerContainsIndependentReadAndSharedPermissionControls()
        {
            using(var window=new ProjectAccessWindow()) {
                Assert.AreEqual(6,window.Controls.Count);
                Assert.IsTrue(window.Controls["projectList"] is CheckedListBox);
                Assert.IsTrue(window.Controls["sharedContext"] is CheckBox);
                Assert.AreEqual(DialogResult.OK,((Button)window.AcceptButton).DialogResult);
                Assert.AreEqual(DialogResult.Cancel,((Button)window.CancelButton).DialogResult);
                Assert.IsFalse(window.MaximizeBox);
            }
        }
    }
}
