using System.Collections.Generic;
using System.Windows.Forms;
using CodexVBE.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ProjectAccessWindowTests
    {
        [STATestMethod]
        public void StaticDesignerControlsRetainSelectionsAndSeparateSharedPermission()
        {
            using (var window = new ProjectAccessWindow())
            {
                window.Populate(new[] {
                    new KeyValuePair<string, string>("C:\\B.xlsm", "B"),
                    new KeyValuePair<string, string>("C:\\C.xlsm", "C") }, new[] { "C:\\B.xlsm" }, false);
                CollectionAssert.AreEqual(new[] { "C:\\B.xlsm" }, window.SelectedProjects);
                Assert.IsFalse(window.SharedContext);
                var list = (CheckedListBox)window.Controls["projectList"];
                list.SetItemChecked(1, true);
                Assert.AreEqual(2, window.SelectedProjects.Length);
                ((CheckBox)window.Controls["sharedContext"]).Checked = true;
                Assert.IsTrue(window.SharedContext);
                Assert.AreEqual(DialogResult.OK, ((Button)window.AcceptButton).DialogResult);
                Assert.AreEqual(DialogResult.Cancel, ((Button)window.CancelButton).DialogResult);
            }
        }
    }
}
