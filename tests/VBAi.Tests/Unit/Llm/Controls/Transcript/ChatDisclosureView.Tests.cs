using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using VBAi.Tests.Infrastructure;
namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ChatDisclosureViewContractTests
    {
        /// <summary>Caption changes leave the host's current body visibility alone and do not publish expansion events.</summary>
        [STATestMethod]
        [DataRow(false)][DataRow(true)]
        public void CaptionRefreshDoesNotOverrideTheCurrentBodyVisibility(bool expanded)
        {
            using (var view = new ChatDisclosureView())
            {
                view.Expanded = expanded;
                // A host may hide nested content while arranging or unrealizing it.
                view.body.Visible = false;
                int transitions = 0;
                view.ExpansionChanged += (sender, args) => transitions++;
                foreach (string title in new[] { "First action", "Next action", "Next action", null })
                {
                    view.Title = title;
                    Assert.AreEqual(expanded, view.Expanded);
                    Assert.IsFalse(view.body.Visible);
                }
                Assert.AreEqual(0, transitions);
                view.Expanded = !expanded;
                Assert.AreEqual(!expanded, view.body.Visible);
                Assert.AreEqual(1, transitions);
            }
        }

        [STATestMethod]
        public void ConstructorCreatesNamedThemedControls() { using (var theme = new ThemeScope()) using (var view = new ChatDisclosureView()) { Assert.AreEqual("ChatDisclosureView", view.Name); Assert.IsTrue(view.Controls.Count > 0); Assert.AreEqual(UiTheme.Background, view.BackColor); Assert.AreEqual(UiTheme.Foreground, view.ForeColor); } }
        [STATestMethod]
        public void ExpansionChangesCaptionBodyAndNotifiesOnlyTransitions()
        {
            using (var view = new ChatDisclosureView())
            {
                Assert.AreSame(view.body, view.ContentPanel); Assert.IsFalse(view.Expanded);
                view.Title = "custom title"; Assert.AreEqual("custom title", view.Title); StringAssert.Contains(view.toggle.Text, "custom title");
                view.Expanded = false; view.Expanded = true; Assert.IsTrue(view.body.Visible); StringAssert.StartsWith(view.toggle.Text, "▾ ");
                int changes = 0; view.ExpansionChanged += (s, e) => { Assert.AreSame(view, s); Assert.AreSame(EventArgs.Empty, e); changes++; };
                view.Expanded = true; Assert.AreEqual(0, changes); view.toggle.PerformClick(); Assert.IsFalse(view.Expanded); Assert.IsFalse(view.body.Visible); Assert.AreEqual(1, changes);
                view.toggle.PerformClick(); Assert.IsTrue(view.Expanded); Assert.AreEqual(2, changes);
                var toggle = view.toggle; view.toggle = null; view.Title = null; Assert.IsNull(view.Title); view.toggle = toggle;
            }
        }
    }
}
