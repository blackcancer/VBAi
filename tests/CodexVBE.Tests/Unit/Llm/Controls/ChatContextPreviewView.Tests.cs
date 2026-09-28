using System.Windows.Forms;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ChatContextPreviewViewTests
    {
        /// <summary>Vérifie aussi la libération normale des ressources du cadre et de l'éditeur.</summary>
        [STATestMethod]
        public void NormalDesignerDisposalReleasesResources()
        {
            var view = new ChatContextPreviewView();
            var content = (TextBox)view.Controls.Find("content", true)[0];
            view.Dispose(); Assert.IsTrue(view.IsDisposed); Assert.IsTrue(content.IsDisposed);
        }

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
                UiInvoke.Call(typeof(ChatContextPreviewView), "Dispose", view, false); Assert.IsFalse(view.IsDisposed);
                var components = UiInvoke.Field<System.ComponentModel.IContainer>(view, "components"); components.Dispose();
                typeof(ChatContextPreviewView).GetField("components", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).SetValue(view, null);
            }
        }
    }
}
