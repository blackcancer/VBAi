using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using VBAi.Tests.Infrastructure;
namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ChatMarkdownTests
    {
        [STATestMethod]
        public void MarkdownBlocksPreserveCodeTableQuoteListsAndDirection()
        {
            using (var theme = new ThemeScope())
            using (var culture = new LocalizationScope())
            using (var view = new ChatTextContentView())
            {
                view.ShowMarkdown(null, null, r => { }, e => { }); Assert.AreEqual("", view.MarkdownView.SourceText);
                view.ShowMarkdown("# Heading\n\nparagraph `code` *italic* **bold** ~~deleted~~\nsoft\nline  \nhard\n\n> quoted\n\n---\n\n3. third\n4. fourth\n\n- [x] done\n- [ ] pending\n\n| A | B |\n|---|---|\n| x | y |\n\n```vba\nSub Test()\n' comment\nEnd Sub\n```\n\n<div>raw</div>", null, r => { }, e => { });
                var textViews = SharedMarkdownViewFixture.TextViews(view);
                Assert.IsTrue(textViews.All(text => text.Content.ReadOnly));
                Assert.IsTrue(textViews.All(text => text.Content.RightToLeft == RightToLeft.No));
                string rendered = string.Join("\n", textViews.Select(text => text.Content.Text));
                StringAssert.Contains(rendered, "3. third"); StringAssert.Contains(rendered, "4. fourth");
                StringAssert.Contains(rendered, "quoted"); StringAssert.Contains(rendered, "Sub Test()");
                var grid = SharedMarkdownViewFixture.Descendants(view.MarkdownView).OfType<Blackcancer.ChatUI.Views.ChatTableView>().Single().Grid;
                CollectionAssert.AreEqual(new[] { "A", "B" }, grid.Columns.Cast<DataGridViewColumn>().Select(column => column.HeaderText).ToArray());
                CollectionAssert.AreEqual(new[] { "x", "y" }, grid.Rows[0].Cells.Cast<DataGridViewCell>().Select(cell => Convert.ToString(cell.Value)).ToArray());
                var paragraph = textViews.Single(text => text.Content.Text.Contains("bold")).Content;
                paragraph.Select(paragraph.Text.IndexOf("bold"), 4); Assert.IsTrue(paragraph.SelectionFont.Bold);
                paragraph.Select(paragraph.Text.IndexOf("italic"), 6); Assert.IsTrue(paragraph.SelectionFont.Italic);
                paragraph.Select(paragraph.Text.IndexOf("deleted"), 7); Assert.IsTrue(paragraph.SelectionFont.Strikeout);
                var code = SharedMarkdownViewFixture.CodeViews(view).Single();
                code.Content.Select(code.Content.Text.IndexOf("Sub Test()"), 3); Assert.AreEqual("Consolas", code.Content.SelectionFont.FontFamily.Name);
                LocalizationScope.Set("ar-SA"); view.ShowMarkdown("text", null, r => { }, e => { }); Assert.AreEqual(RightToLeft.Yes, SharedMarkdownViewFixture.TextViews(view).Single().Content.RightToLeft);
                view.ShowPlain("code", true); Assert.AreEqual(RightToLeft.No, view.content.RightToLeft);
            }
        }
        [STATestMethod]
        public void ReferenceLinksWebLinksAndCopyActionsReportErrorsWithoutExternalEffects()
        {
            using (var theme = new ThemeScope())
            using (var culture = new LocalizationScope())
            using (var view = new ChatTextContentView())
            {
                var copy = ChatMarkdown.CopyText; var open = ChatMarkdown.OpenLink;
                try
                {
                    string copied = null, opened = null, error = null; VbeChatReference navigated = null;
                    ChatMarkdown.CopyText = s => copied = s; ChatMarkdown.OpenLink = s => opened = s;
                    var reference = new VbeChatReference { Project = "Project" };
                    view.ShowMarkdown("#Project #missing [web](https://example.invalid) [blocked](file:///C:/x) <https://auto.invalid> <b>html</b>\n\n```\ncode\n```", new Dictionary<string, VbeChatReference> { { "#Project", reference } }, r => navigated = r, e => error = e);
                    var links = SharedMarkdownViewFixture.Links(view); Assert.AreEqual(3, links.Length);
                    links[0].Activate(); Assert.AreSame(reference, navigated);
                    links[1].Activate(); Assert.AreEqual("https://example.invalid", opened);
                    links[2].Activate(); Assert.AreEqual("https://auto.invalid", opened);
                    var code = SharedMarkdownViewFixture.CodeViews(view).Single(); SharedMarkdownViewFixture.CopyCode(code); Assert.AreEqual("code", copied);
                    ChatMarkdown.OpenLink = s => { throw new InvalidOperationException("open failed"); }; ChatMarkdown.CopyText = s => { throw new InvalidOperationException("copy failed"); };
                    foreach (var link in links.Skip(1)) { error = null; link.Activate(); Assert.AreEqual("open failed", error); }
                    SharedMarkdownViewFixture.CopyCode(code); Assert.AreEqual("copy failed", error);
                }
                finally { ChatMarkdown.CopyText = copy; ChatMarkdown.OpenLink = open; }
            }
        }
    }
}
