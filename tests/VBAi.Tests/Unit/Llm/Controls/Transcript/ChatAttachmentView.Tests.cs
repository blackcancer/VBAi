using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using VBAi;
using VBAi.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ChatAttachmentViewContractTests
    {
        [STATestMethod]
        public void ConstructorCreatesNamedThemedControls() { using(var theme=new ThemeScope()) using(var view=new ChatAttachmentView()) { Assert.AreEqual("ChatAttachmentView",view.Name); Assert.IsTrue(view.Controls.Count>0); Assert.AreEqual(UiTheme.Background,view.BackColor); Assert.AreEqual(UiTheme.Foreground,view.ForeColor); } }
    }
}