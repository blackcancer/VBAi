using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ChatFormRecoveryViewContractTests
    {
        [STATestMethod]
        public void ConstructorCreatesNamedThemedControls() { using(var theme=new ThemeScope()) using(var view=new ChatFormRecoveryView()) { Assert.AreEqual("ChatFormRecoveryView",view.Name); Assert.IsTrue(view.Controls.Count>0); Assert.AreEqual(UiTheme.Background,view.BackColor); Assert.AreEqual(UiTheme.Foreground,view.ForeColor); } }
    }
}