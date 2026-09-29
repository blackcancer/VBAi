using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using VBAi;
using VBAi.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ChatTextContentViewContractTests
    {
        [STATestMethod]
        public void ConstructorCreatesNamedThemedControls() { using(var theme=new ThemeScope()) using(var view=new ChatTextContentView()) { Assert.AreEqual("ChatTextContentView",view.Name); Assert.IsTrue(view.Controls.Count>0); Assert.AreEqual(UiTheme.Background,view.BackColor); Assert.AreEqual(UiTheme.Foreground,view.ForeColor); } }
        [STATestMethod]
        public void NativeTextInteractionCopiesSelectionAndCodeAndHandlesMouseKeyboardAndMenu()
        {
            using(var theme=new ThemeScope())
            using(var view=new ChatTextContentView()) {
                var copy=ChatMarkdown.CopyText; string copied=null,error=null; int activations=0;
                try {
                    ChatMarkdown.CopyText=s=>copied=s; view.ErrorHandler=s=>error=s;
                    view.ShowPlain(null); Assert.AreEqual("",view.content.Text); view.ShowPlain("abcdef");
                    view.copySelection.PerformClick(); Assert.AreEqual("abcdef",copied);
                    view.content.Select(1,2); view.copySelection.PerformClick(); Assert.AreEqual("bc",copied);
                    var handle=view.content.Handle; view.Append(null); Assert.AreEqual("abcdef",view.content.Text);
                    var first=view.OwnFont("Consolas",8,FontStyle.Bold); Assert.AreSame(first,view.OwnFont("Consolas",8,FontStyle.Bold));
                    Assert.AreNotSame(first,view.OwnFont("Segoe UI",8,FontStyle.Bold)); Assert.AreNotSame(first,view.OwnFont("Consolas",9,FontStyle.Bold)); Assert.AreNotSame(first,view.OwnFont("Consolas",8,FontStyle.Italic));
                    view.actions.Add(new ChatTextContentView.TextAction { Start=1,Length=2,Invoke=()=>activations++ });
                    view.ActivateAt(0); view.ActivateAt(1); view.ActivateAt(2); view.ActivateAt(3); Assert.AreEqual(2,activations);
                    view.actions.Add(new ChatTextContentView.TextAction {Start=4,Length=1}); view.ActivateAt(4);
                    view.actions.Add(new ChatTextContentView.TextAction {Start=5,Length=1,Invoke=()=>{throw new InvalidOperationException("action error");}}); view.ActivateAt(5); Assert.AreEqual("action error",error);
                    view.content.Select(1,0); var enter=new KeyEventArgs(Keys.Enter); TranscriptFixture.Event(view.content,"OnKeyDown",enter); Assert.IsTrue(enter.Handled); Assert.AreEqual(3,activations);
                    var other=new KeyEventArgs(Keys.A); TranscriptFixture.Event(view.content,"OnKeyDown",other); Assert.IsFalse(other.Handled);
                    var point=view.content.GetPositionFromCharIndex(1); TranscriptFixture.Event(view.content,"OnMouseUp",new MouseEventArgs(MouseButtons.Right,1,point.X,point.Y,0)); Assert.AreEqual(3,activations);
                    view.content.Select(1,1); TranscriptFixture.Event(view.content,"OnMouseUp",new MouseEventArgs(MouseButtons.Left,1,point.X,point.Y,0)); Assert.AreEqual(3,activations);
                    view.content.Select(1,0); TranscriptFixture.Event(view.content,"OnMouseUp",new MouseEventArgs(MouseButtons.Left,1,point.X,point.Y,0)); Assert.AreEqual(4,activations);
                    view.copyCode.PerformClick(); Assert.AreEqual("bc",copied);
                    UiInvoke.Call(typeof(ContextMenuStrip),"OnOpening",view.copyMenu,new CancelEventArgs()); Assert.IsFalse(view.copyCode.Available);
                    view.content.Select(0,0); UiInvoke.Call(typeof(ToolStripItem),"OnClick",view.copyCode,EventArgs.Empty); UiInvoke.Call(typeof(ContextMenuStrip),"OnOpening",view.copyMenu,new CancelEventArgs()); Assert.IsFalse(view.copyCode.Available);
                    view.actions.Insert(0,new ChatTextContentView.TextAction {Start=0,Length=1,Code="code"}); UiInvoke.Call(typeof(ContextMenuStrip),"OnOpening",view.copyMenu,new CancelEventArgs()); Assert.IsTrue(view.copyCode.Available);
                    view.copyCode.PerformClick(); Assert.AreEqual("code",copied); ChatMarkdown.CopyText=s=>{throw new InvalidOperationException("copy error");}; view.copySelection.PerformClick(); Assert.AreEqual("copy error",error);
                } finally { ChatMarkdown.CopyText=copy; }
            }
        }
        [STATestMethod]
        public void RichTextReflowBoundsHeightAndSupportsEmptyDisposedOrInitializingContent()
        {
            using(var theme=new ThemeScope())
            using(var culture=new LocalizationScope())
            using(var view=new ChatTextContentView()) {
                UiInvoke.Call(typeof(ChatTextContentView),"ResizeText",view); view.ShowPlain("without handle");
                view.ShowPlain("small"); var handle=view.content.Handle; view.Width=200; Assert.IsTrue(view.content.Height>=24); Assert.AreEqual(RichTextBoxScrollBars.None,view.content.ScrollBars);
                view.ShowPlain(string.Join("\n",new string[300]).Replace("\n","line\n")); Assert.AreEqual(1200,view.content.Height); Assert.AreEqual(RichTextBoxScrollBars.Vertical,view.content.ScrollBars);
                TranscriptFixture.Event(view.content,"OnContentsResized",new ContentsResizedEventArgs(new Rectangle(0,0,10,1))); Assert.AreEqual(24,view.content.Height);
                TranscriptFixture.Event(view.content,"OnContentsResized",new ContentsResizedEventArgs(new Rectangle(0,0,10,2000))); Assert.AreEqual(1200,view.content.Height);
                view.ShowPlain(""); Assert.AreEqual("",view.content.Text); LocalizationScope.Set("ar-SA"); view.ShowPlain("rtl"); Assert.AreEqual(RightToLeft.Yes,view.content.RightToLeft);
                view.ShowPlain("code",true); Assert.AreEqual(RightToLeft.No,view.content.RightToLeft); Assert.AreEqual("Consolas",view.content.Font.FontFamily.Name);
                var content=view.content; view.content=null; UiInvoke.Call(typeof(ChatTextContentView),"ResizeText",view); view.content=content;
                content.Dispose(); UiInvoke.Call(typeof(ChatTextContentView),"ResizeText",view); Assert.IsTrue(content.IsDisposed);
            }
        }
    }
}
