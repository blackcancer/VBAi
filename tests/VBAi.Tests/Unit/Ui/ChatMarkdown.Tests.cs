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
                view.ShowMarkdown(null, null, r => { }, e => { }); Assert.AreEqual("", view.content.Text);
                view.ShowMarkdown("# Heading\n\nparagraph `code` *italic* **bold** ~~deleted~~\nsoft\nline  \nhard\n\n> quoted\n\n---\n\n3. third\n4. fourth\n\n- [x] done\n- [ ] pending\n\n| A | B |\n|---|---|\n| x | y |\n\n```vba\nSub Test()\n' comment\nEnd Sub\n```\n\n<div>raw</div>", null, r => { }, e => { });
                Assert.IsTrue(view.content.ReadOnly); Assert.AreEqual(RightToLeft.No, view.content.RightToLeft);
                StringAssert.Contains(view.content.Text, "3. third"); StringAssert.Contains(view.content.Text, "4. fourth");
                StringAssert.Contains(view.content.Text, "A\tB"); StringAssert.Contains(view.content.Text, "x\ty");
                StringAssert.Contains(view.content.Text, "quoted"); StringAssert.Contains(view.content.Text, "Sub Test()");
                view.content.Select(view.content.Text.IndexOf("bold"), 4); Assert.IsTrue(view.content.SelectionFont.Bold);
                view.content.Select(view.content.Text.IndexOf("italic"), 6); Assert.IsTrue(view.content.SelectionFont.Italic);
                view.content.Select(view.content.Text.IndexOf("deleted"), 7); Assert.IsTrue(view.content.SelectionFont.Strikeout);
                view.content.Select(view.content.Text.IndexOf("Sub Test()"), 3); Assert.AreEqual("Consolas", view.content.SelectionFont.FontFamily.Name);
                Assert.AreEqual(1, view.actions.Count(a => a.Code != null));
                LocalizationScope.Set("ar-SA"); view.ShowMarkdown("text", null, r => { }, e => { }); Assert.AreEqual(RightToLeft.Yes, view.content.RightToLeft);
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
                    var links = view.actions.Where(a => a.Invoke != null).ToArray(); Assert.AreEqual(3, links.Length);
                    view.ActivateAt(links[0].Start); Assert.AreSame(reference, navigated);
                    view.ActivateAt(links[1].Start); Assert.AreEqual("https://example.invalid", opened);
                    view.ActivateAt(links[2].Start); Assert.AreEqual("https://auto.invalid", opened);
                    var code = view.actions.Single(a => a.Code != null); view.content.Select(code.Start, 0); view.copyCode.PerformClick(); Assert.AreEqual("code", copied);
                    ChatMarkdown.OpenLink = s => { throw new InvalidOperationException("open failed"); }; ChatMarkdown.CopyText = s => { throw new InvalidOperationException("copy failed"); };
                    foreach (var link in links.Skip(1)) { error = null; view.ActivateAt(link.Start); Assert.AreEqual("open failed", error); }
                    view.copyCode.PerformClick(); Assert.AreEqual("copy failed", error);
                }
                finally { ChatMarkdown.CopyText = copy; ChatMarkdown.OpenLink = open; }
            }
        }
    }
}
