using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Drawing;
using System.Windows.Forms;
using VBAi.Tests.Infrastructure;
namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ChatLinkViewContractTests
    {
        [STATestMethod]
        public void ConstructorCreatesNamedThemedControls() { using (var theme = new ThemeScope()) using (var view = new ChatLinkView()) { Assert.AreEqual("ChatLinkView", view.Name); Assert.IsTrue(view.Controls.Count > 0); Assert.AreEqual(UiTheme.Background, view.BackColor); Assert.AreEqual(UiTheme.Foreground, view.ForeColor); } }
        [STATestMethod]
        public void TokenPreferredSizeIncludesPaddingAndMarginsAndSupportsInitialization()
        {
            using (var view = new ChatLinkView())
            {
                view.link.Text = "@Project.Module.Procedure"; var size = view.link.GetPreferredSize(Size.Empty);
                var layout = UiInvoke.Field<TableLayoutPanel>(view, "layout"); var actual = view.GetPreferredSize(new Size(500, 500));
                Assert.AreEqual(size.Width + layout.Padding.Horizontal + view.link.Margin.Horizontal, actual.Width);
                Assert.AreEqual(size.Height + layout.Padding.Vertical + view.link.Margin.Vertical, actual.Height);
                var link = view.link; view.link = null; Assert.IsTrue(view.GetPreferredSize(Size.Empty).Width >= 0); view.link = link;
            }
        }
    }
}
