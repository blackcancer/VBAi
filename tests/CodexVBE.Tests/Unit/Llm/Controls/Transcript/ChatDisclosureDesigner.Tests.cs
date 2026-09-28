using System.ComponentModel.Design;
using System.Windows.Forms;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    [TestClass,TestCategory("Unit")]
    public sealed class ChatDisclosureDesignerContractTests
    {
        [STATestMethod]
        public void DisclosureBodyAcceptsDesignerComponents()
        {
            using(var surface=new DesignSurface(typeof(UserControl))) {
                var host=(IDesignerHost)surface.GetService(typeof(IDesignerHost)); var disclosure=(ChatDisclosureView)host.CreateComponent(typeof(ChatDisclosureView));
                ((UserControl)host.RootComponent).Controls.Add(disclosure); Assert.IsInstanceOfType(host.GetDesigner(disclosure),typeof(ChatDisclosureDesigner));
                Assert.IsNotNull(disclosure.ContentPanel.Site); var button=(Button)host.CreateComponent(typeof(Button)); disclosure.ContentPanel.Controls.Add(button);
                Assert.AreSame(disclosure.ContentPanel,button.Parent); Assert.IsNotNull(host.GetDesigner(button));
            }
        }
    }
}