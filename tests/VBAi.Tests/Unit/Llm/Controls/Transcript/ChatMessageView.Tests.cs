using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Infrastructure;
namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ChatMessageViewContractTests
    {
        [STATestMethod]
        public void ConstructorCreatesNamedThemedControls() { using (var theme = new ThemeScope()) using (var view = new ChatMessageView()) { Assert.AreEqual("ChatMessageView", view.Name); Assert.IsTrue(view.Controls.Count > 0); Assert.AreEqual(UiTheme.Background, view.BackColor); Assert.AreEqual(UiTheme.Foreground, view.ForeColor); } }
        [STATestMethod]
        public void MessageActionPropertiesReturnDesignerButtons()
        {
            using (var view = new ChatMessageView()) { Assert.AreSame(view.copy, view.CopyButton); Assert.AreSame(view.fork, view.ForkButton); }
        }
    }
}
