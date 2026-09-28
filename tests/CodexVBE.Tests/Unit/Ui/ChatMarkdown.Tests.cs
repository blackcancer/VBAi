using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    /// <summary>Vérifie le rendu Markdown de la conversation et ses actions de lien.</summary>
    [TestClass]
    public sealed class ChatMarkdownTests
    {
        /// <summary>Préserve blocs de code, tableaux, citations, listes et direction du texte.</summary>
        [STATestMethod,TestCategory("Unit")]
        public void MarkdownBlocksPreserveCodeTableQuoteListsAndDirection()
        {
            using(var theme=new ThemeScope())
            using(var culture=new LocalizationScope())
            {
                Assert.AreEqual(0,ChatMarkdown.Render(null,null,r=>{},e=>{}).Document.Blocks.Count);
                var view=ChatMarkdown.Render("# Heading\n\nparagraph `code` *italic* **bold** ~~deleted~~\nsoft\nline  \nhard\n\n> quoted\n\n---\n\n3. third\n4. fourth\n\n- [x] done\n- [ ] pending\n\n| A | B |\n|---|---|\n| x | y |\n\n```vba\nSub Test()\n' comment\nEnd Sub\n```\n\n<div>raw</div>",null,r=>{},e=>{});
                Assert.IsTrue(view.IsReadOnly); Assert.IsTrue(view.IsDocumentEnabled);
                Assert.AreEqual(FlowDirection.LeftToRight,view.Document.FlowDirection);
                Assert.AreEqual(22,((Paragraph)view.Document.Blocks.FirstBlock).FontSize);
                var ordered=view.Document.Blocks.OfType<System.Windows.Documents.List>().First(); Assert.AreEqual(TextMarkerStyle.Decimal,ordered.MarkerStyle); Assert.AreEqual(3,ordered.StartIndex);
                var table=view.Document.Blocks.OfType<System.Windows.Documents.Table>().Single(); Assert.AreEqual(2,table.RowGroups[0].Rows.Count); Assert.AreEqual(FontWeights.SemiBold,table.RowGroups[0].Rows[0].Cells[0].FontWeight); Assert.AreEqual(FontWeights.Normal,table.RowGroups[0].Rows[1].Cells[0].FontWeight);
                Assert.AreEqual(1,view.Document.Blocks.OfType<Section>().Count());
                var panel=(StackPanel)view.Document.Blocks.OfType<BlockUIContainer>().Single().Child;
                var code=(RichTextBox)panel.Children[1]; Assert.AreEqual(FlowDirection.LeftToRight,code.Document.FlowDirection); Assert.IsTrue(new TextRange(code.Document.ContentStart,code.Document.ContentEnd).Text.Contains("Sub Test()"));
                LocalizationScope.Set("ar-SA"); Assert.AreEqual(FlowDirection.RightToLeft,ChatMarkdown.Render("text",null,r=>{},e=>{}).Document.FlowDirection);
                var document=new FlowDocument(); var ast=new MarkdownDocument(); ast.Add(new LinkReferenceDefinitionGroup()); ast.Add(new ParagraphBlock());
                foreach(var start in new[]{"0","invalid","1"}) ast.Add(new ListBlock(null){IsOrdered=true,OrderedStart=start});
                UiInvoke.Call(typeof(ChatMarkdown),"AddBlocks",null,document.Blocks,ast,null,new Action<VbeChatReference>(r=>{}),new Action<string>(e=>{}));
                Assert.AreEqual(4,document.Blocks.Count);
            }
        }
        /// <summary>Traite liens de références, liens Web et erreurs de copie sans effet externe.</summary>
        [STATestMethod,TestCategory("Unit")]
        public void ReferenceLinksWebLinksAndCopyActionsReportErrorsWithoutExternalEffects()
        {
            using(var theme=new ThemeScope())
            using(var culture=new LocalizationScope())
            {
                var copy=ChatMarkdown.CopyText; var open=ChatMarkdown.OpenLink;
                try
                {
                    string copied=null,opened=null,error=null; VbeChatReference navigated=null;
                    ChatMarkdown.CopyText=s=>copied=s; ChatMarkdown.OpenLink=s=>opened=s;
                    var reference=new VbeChatReference {Project="Project"}; var refs=new Dictionary<string,VbeChatReference>{{"#Project",reference}};
                    var view=ChatMarkdown.Render("#Project #missing [web](https://example.invalid) [blocked](file:///C:/x) <https://auto.invalid> <b>html</b>\n\n```\ncode\n```",refs,r=>navigated=r,e=>error=e);
                    var paragraph=(Paragraph)view.Document.Blocks.FirstBlock; var links=paragraph.Inlines.OfType<Hyperlink>().ToArray(); Assert.AreEqual(3,links.Length);
                    links[0].RaiseEvent(new RoutedEventArgs(Hyperlink.ClickEvent)); Assert.AreSame(reference,navigated);
                    links[1].RaiseEvent(new RoutedEventArgs(Hyperlink.ClickEvent)); Assert.AreEqual("https://example.invalid",opened);
                    links[2].RaiseEvent(new RoutedEventArgs(Hyperlink.ClickEvent)); Assert.AreEqual("https://auto.invalid",opened);
                    var button=(Button)((StackPanel)view.Document.Blocks.OfType<BlockUIContainer>().Single().Child).Children[0]; button.RaiseEvent(new RoutedEventArgs(Button.ClickEvent)); Assert.AreEqual("code",copied);
                    ChatMarkdown.OpenLink=s=>{throw new InvalidOperationException("open failed");}; ChatMarkdown.CopyText=s=>{throw new InvalidOperationException("copy failed");};
                    foreach(var link in links.Skip(1)) {error=null; link.RaiseEvent(new RoutedEventArgs(Hyperlink.ClickEvent)); Assert.AreEqual("open failed",error);}
                    button.RaiseEvent(new RoutedEventArgs(Button.ClickEvent)); Assert.AreEqual("copy failed",error);
                    var source=new ContainerInline(); var nested=new ContainerInline(); nested.AppendChild(new LiteralInline("nested")); source.AppendChild(nested); source.AppendChild(new HtmlEntityInline());
                    var destination=new Paragraph(); UiInvoke.Call(typeof(ChatMarkdown),"AddInlines",null,destination.Inlines,source,null,new Action<VbeChatReference>(r=>{}),new Action<string>(e=>{})); Assert.AreEqual("nested",((Run)destination.Inlines.Single()).Text);
                }
                finally {ChatMarkdown.CopyText=copy;ChatMarkdown.OpenLink=open;}
            }
        }
    }
}
