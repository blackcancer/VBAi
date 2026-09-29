using System;
using System.Collections.Generic;
using System.Drawing;
using VBAi;
using VBAi.Tests.Infrastructure;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using Markdig.Extensions.Tables;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace VBAi.Tests.Unit
{
    [TestClass,TestCategory("Unit")]
    public sealed class ChatNativeMarkdownContractTests
    {
        [STATestMethod]
        public void AdditionalAstContainersAndNonLeafTableCellsHaveDefinedRendering()
        {
            using(var theme=new ThemeScope())
            using(var view=new ChatTextContentView()) {
                var document=new MarkdownDocument(); var nested=new MarkdownDocument(); nested.Add(new ParagraphBlock {Inline=new ContainerInline().AppendChild(new LiteralInline("nested"))}); document.Add(nested);
                document.Add(new HtmlBlock(null) { Lines=new Markdig.Helpers.StringLineGroup("html block") });
                var table=new Table(); var row=new TableRow(); var cell=new TableCell(); cell.Add(new MarkdownDocument()); cell.Add(new ParagraphBlock()); row.Add(cell); table.Add(row); document.Add(table);
                var emptyList=new ListBlock(null) {OrderedStart="0",IsOrdered=true}; document.Add(emptyList);
                UiInvoke.Call(typeof(ChatNativeMarkdown),"Blocks",null,view,document,null,new Action<VbeChatReference>(r=>{}),new Action<string>(e=>Assert.Fail(e)));
                StringAssert.Contains(view.content.Text,"nested"); StringAssert.Contains(view.content.Text,"html block"); StringAssert.Contains(view.content.Text,"\t");
                var inlines=new ContainerInline(); var inlineContainer=new ContainerInline(); inlineContainer.AppendChild(new LiteralInline("nested inline")); inlines.AppendChild(inlineContainer);
                inlines.AppendChild(new AutolinkInline("file:///C:/blocked"));
                UiInvoke.Call(typeof(ChatNativeMarkdown),"Inlines",null,view,inlines,null,new Action<VbeChatReference>(r=>{}),new Action<string>(e=>Assert.Fail(e)),FontStyle.Regular,9.5f);
                StringAssert.Contains(view.content.Text,"nested inline"); Assert.AreEqual(0,view.actions.Count);
            }
        }
    }
}