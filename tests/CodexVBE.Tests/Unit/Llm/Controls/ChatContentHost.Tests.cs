using System.ComponentModel;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ChatContentHostTests
    {
        [STATestMethod]
        public void ContentHostUsesStandardWinFormsDesignerAndStartsEmpty()
        {
            using (var host = new ChatContentHost())
            {
                Assert.IsNull(host.Child);
                var attribute = (DesignerAttribute)TypeDescriptor.GetAttributes(host)[typeof(DesignerAttribute)];
                StringAssert.StartsWith(attribute.DesignerTypeName, "System.Windows.Forms.Design.ControlDesigner,");
            }
        }
    }
}
