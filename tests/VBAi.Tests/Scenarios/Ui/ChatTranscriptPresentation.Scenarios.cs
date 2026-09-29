using System;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Linq;
using System.Windows.Forms;
using VBAi;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ChatTranscriptPresentationScenarios
    {
        [STATestMethod]
        public void LoadedTranscriptCardDisplaysNativeMessageAndExpandedDiff()
        {
            var card = new ChatChangeCardView(); card.diff.ShowDiff("old", "new"); card.section.Expanded = true;
            var message = new ChatMessageView(); message.message.ShowPlain("A visible native message");
            var reference = new ChatLinkView(); reference.link.Text = "@Budget.Module.Calculer"; message.references.Controls.Add(reference);
            using (var first = new ChatDesignerHost(message))
            using (var second = new ChatDesignerHost(card)) {
                var stack = new System.Windows.Controls.StackPanel(); stack.Children.Add(first); stack.Children.Add(second);
                var window = new System.Windows.Window { Content = stack, Width = 600, Height = 900, ShowInTaskbar = false };
                try {
                    window.Show(); window.UpdateLayout(); Application.DoEvents(); window.UpdateLayout();
                    Assert.IsTrue(message.Visible); Assert.IsTrue(message.Height > 40, "Message height " + message.Height);
                    Assert.IsTrue(message.message.Height >= 24, "Message body height " + message.message.Height);
                    Assert.IsTrue(message.message.content.Height >= 24);
                    Assert.IsTrue(message.message.content.Visible); Assert.IsTrue(message.message.content.IsHandleCreated);
                    Assert.IsTrue(message.headingActions.ClientRectangle.Contains(message.fork.Bounds), "Fork action clipped");
                    Assert.IsTrue(reference.Width > 50, "Reference width " + reference.Width);
                    Assert.IsTrue(reference.Visible); Assert.IsTrue(reference.Height > 20, "Reference height " + reference.Height);
                    Assert.IsTrue(card.diff.Visible, "Diff not visible");
                    Assert.IsTrue(card.diff.Height > 100, "Diff height " + card.diff.Height);
                    Assert.IsTrue(card.section.Height > 100, "Section height " + card.section.Height);
                } finally { window.Close(); }
            }
        }
    }
}
