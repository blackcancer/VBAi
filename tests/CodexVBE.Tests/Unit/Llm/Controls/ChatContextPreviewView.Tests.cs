using System.Windows.Forms;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ChatContextPreviewViewTests
    {
        [STATestMethod]
        public void DesignerPreviewRebindsTextAndPreservesItsReadonlyEditor()
        {
            using (var view = new ChatContextPreviewView())
            {
                var content = (TextBox)view.Controls.Find("content", true)[0];
                view.ShowContent("Module", "Sub Example()\r\nEnd Sub");
                Assert.IsTrue(content.ReadOnly);
                Assert.AreEqual("Sub Example()\r\nEnd Sub", content.Text);
                view.ShowContent(null, null);
                Assert.AreEqual("", content.Text);
                Assert.AreSame(content, view.Controls.Find("content", true)[0]);
            }
        }
    }
}
