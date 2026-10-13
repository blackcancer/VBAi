using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Unit
{
    /// <summary>Verifies that transcript metadata consumes space only while it contains visible rows.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class ChatMessageViewMetadataTests
    {
        [STATestMethod]
        [DataRow("attachments", 300)]
        [DataRow("attachments", 500)]
        [DataRow("attachments", 800)]
        [DataRow("references", 300)]
        [DataRow("references", 500)]
        [DataRow("references", 800)]
        [DataRow("targets", 300)]
        [DataRow("targets", 500)]
        [DataRow("targets", 800)]
        public void MetadataAppearsForAddedRowsAndReclaimsItsSpaceAfterTheLastRemoval(string name, int width)
        {
            using (var card = new ChatMessageView())
            using (var first = new Panel { Size = new Size(100, 80), Margin = Padding.Empty })
            using (var second = new Panel { Size = new Size(100, 80), Margin = Padding.Empty })
            {
                card.Width = width;
                card.CreateControl();
                var row = (FlowLayoutPanel)card.Controls.Find(name, true).Single();
                var constraint = new Size(width, 0);
                Assert.IsFalse(row.Visible, "An empty metadata section must not occupy a transcript row.");
                int emptyHeight = card.GetPreferredSize(constraint).Height;
                row.Controls.Add(first);
                card.PerformLayout();
                Assert.IsTrue(row.Visible);
                Assert.IsTrue(card.GetPreferredSize(constraint).Height > emptyHeight);
                row.Controls.Add(second);
                row.Controls.Remove(first);
                Assert.IsTrue(row.Visible, "Removing one of several rows must preserve the remaining metadata.");
                row.Controls.Remove(second);
                card.PerformLayout();
                Assert.IsFalse(row.Visible);
                Assert.AreEqual(emptyHeight, card.GetPreferredSize(constraint).Height,
                    "The last removal must reclaim the row and its margins.");
            }
        }

        [STATestMethod]
        public void EmptyMetadataDoesNotHideMessageContentOrAvailableActions()
        {
            using (var card = new ChatMessageView())
            {
                foreach (string name in new[] { "attachments", "references", "targets" })
                    Assert.IsFalse(card.Controls.Find(name, true).Single().Visible);
                Assert.IsTrue(card.Controls.Find("message", true).Single().Visible);
                Assert.IsTrue(card.CopyButton.Visible);
                Assert.IsTrue(card.ForkButton.Visible);
                var reference = new ChatLinkView();
                card.Controls.Find("references", true).Single().Controls.Add(reference);
                card.Dispose();
                Assert.IsTrue(reference.IsDisposed, "The card continues to own its dynamically added controls.");
            }
        }
    }
}
