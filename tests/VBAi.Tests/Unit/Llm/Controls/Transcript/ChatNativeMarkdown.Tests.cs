using Markdig.Extensions.Tables;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;
using System.Drawing;
using VBAi.Tests.Infrastructure;
namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ChatNativeMarkdownContractTests
    {
        /// <summary>Prose, emphasis and inline code bind the captured target; unknown or fenced tokens do not create actions.</summary>
        [STATestMethod]
        [DataRow("prose", 1)]
        [DataRow("inline-code", 1)]
        [DataRow("emphasis", 1)]
        [DataRow("repeated", 2)]
        [DataRow("unknown", 0)]
        [DataRow("fenced-code", 0)]
        public void ReferenceActionsUseTheCapturedEntryInProseAndInlineCode(string format, int count)
        {
            var original = new VbeChatReference { Project = "Project1", Module = "Module1", Name = "Calculate", Sha256 = "old" };
            var prompt = new ChatEntry { Speaker = "Vous", TurnId = "first", References = new[] { original } };
            var answer = new ChatEntry { Speaker = "Assistant", TurnId = "first" };
            var later = new ChatEntry { Speaker = "Vous", TurnId = "second", References = new[]
                { new VbeChatReference { Project = original.Project, Module = original.Module, Name = original.Name, Sha256 = "new" } } };
            string text = original.Token;
            if (format == "inline-code") text = "`" + text + "`";
            if (format == "emphasis") text = "**" + text + "**";
            if (format == "repeated") text += " then `" + text + "`";
            if (format == "unknown") text = "@Other.Module.Unknown";
            if (format == "fenced-code") text = "```vba\n" + text + "\n```";
            using (var theme = new ThemeScope())
            using (var view = new ChatTextContentView())
            {
                VbeChatReference selected = null;
                view.ShowMarkdown(text, ChatTranscriptReferences.Resolve(new[] { prompt, answer, later }, answer), r => selected = r, error => Assert.Fail(error));
                var navigation = SharedMarkdownViewFixture.Links(view);
                Assert.AreEqual(count, navigation.Length, format);
                foreach (var action in navigation)
                {
                    Assert.AreEqual(original.Token, action.Content.Text.Substring(action.Start, action.Length));
                    action.Activate(); Assert.AreSame(original, selected);
                }
                if (format == "fenced-code")
                {
                    var code = SharedMarkdownViewFixture.CodeViews(view).Single();
                    Assert.AreEqual(original.Token, code.SourceText.TrimEnd('\r', '\n'));
                    string copied = null;
                    code.Services.CopyText = value => copied = value;
                    SharedMarkdownViewFixture.CopyCode(code);
                    Assert.AreEqual(code.SourceText, copied, "The visible code action must copy its own source.");
                    Assert.IsNull(selected);
                }
                if (format == "inline-code")
                {
                    navigation[0].Content.Select(navigation[0].Start, navigation[0].Length);
                    Assert.AreEqual("Consolas", navigation[0].Content.SelectionFont.FontFamily.Name);
                }
            }
        }

        [STATestMethod]
        public void AdditionalAstContainersAndNonLeafTableCellsHaveDefinedRendering()
        {
            using (var theme = new ThemeScope())
            using (var view = new ChatTextContentView())
            {
                var document = new MarkdownDocument(); var nested = new MarkdownDocument(); nested.Add(new ParagraphBlock { Inline = new ContainerInline().AppendChild(new LiteralInline("nested")) }); document.Add(nested);
                document.Add(new HtmlBlock(null) { Lines = new Markdig.Helpers.StringLineGroup("html block") });
                var table = new Table(); var row = new TableRow(); var cell = new TableCell(); cell.Add(new MarkdownDocument()); cell.Add(new ParagraphBlock()); row.Add(cell); table.Add(row); document.Add(table);
                var emptyList = new ListBlock(null) { OrderedStart = "0", IsOrdered = true }; document.Add(emptyList);
                UiInvoke.Call(typeof(ChatNativeMarkdown), "Blocks", null, view, document, null, new Action<VbeChatReference>(r => { }), new Action<string>(e => Assert.Fail(e)));
                StringAssert.Contains(view.content.Text, "nested"); StringAssert.Contains(view.content.Text, "html block"); StringAssert.Contains(view.content.Text, "\t");
                var inlines = new ContainerInline(); var inlineContainer = new ContainerInline(); inlineContainer.AppendChild(new LiteralInline("nested inline")); inlines.AppendChild(inlineContainer);
                inlines.AppendChild(new AutolinkInline("file:///C:/blocked"));
                UiInvoke.Call(typeof(ChatNativeMarkdown), "Inlines", null, view, inlines, null, new Action<VbeChatReference>(r => { }), new Action<string>(e => Assert.Fail(e)), FontStyle.Regular, 9.5f);
                StringAssert.Contains(view.content.Text, "nested inline"); Assert.AreEqual(0, view.actions.Count);
            }
        }
    }
}
