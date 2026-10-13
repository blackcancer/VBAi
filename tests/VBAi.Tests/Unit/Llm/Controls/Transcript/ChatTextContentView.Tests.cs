using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using VBAi.Tests.Infrastructure;
namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ChatTextContentViewContractTests
    {
        /// <summary>Checks append, rewrite, truncation, empty and unchanged snapshots with a retained selection.</summary>
        [STATestMethod]
        [DataRow("prefix tail", 1, 3)]
        [DataRow("replacement", 1, 3)]
        [DataRow("pr", 1, 1)]
        [DataRow("", 0, 0)]
        [DataRow(null, 0, 0)]
        [DataRow("prefix", 1, 3)]
        public void StreamSnapshotPreservesValidSelectionAndPublishesOnce(string snapshot, int start, int length)
        {
            using (var view = new ChatTextContentView())
            {
                view.Width = 400;
                var handle = view.content.Handle;
                view.ShowPlain("prefix");
                view.content.Select(1, 3);
                int changes = 0, initialHeight = view.content.Height;
                view.content.TextChanged += (s, e) =>
                {
                    changes++;
                    Assert.AreEqual(initialHeight, view.content.Height, "Height must settle after publication.");
                };
                view.UpdatePlainText(snapshot);
                Assert.AreEqual(snapshot ?? "", view.content.Text);
                Assert.AreEqual(start, view.content.SelectionStart);
                Assert.AreEqual(length, view.content.SelectionLength);
                Assert.AreEqual(snapshot == "prefix" ? 0 : 1, changes);
                Assert.AreEqual(handle, view.content.Handle);
            }
        }

        /// <summary>Proves that growing snapshots retain native prefix formatting instead of replacing the whole document.</summary>
        [STATestMethod]
        public void StreamAppendRetainsTheNativePrefixAndDisposalIsIgnored()
        {
            using (var control = new RichTextBox())
            using (var font = new Font("Consolas", 11, FontStyle.Bold))
            {
                var handle = control.Handle;
                control.Text = "été\nUnicode 😃";
                control.Select(0, 3);
                control.SelectionFont = font;
                Assert.IsTrue(ChatTextContentView.UpdatePlainTextControl(control, "été\nUnicode 😃\n追加"));
                Assert.AreEqual("été", control.SelectedText);
                Assert.IsTrue(control.SelectionFont.Bold, "Appending must not replace the formatted prefix.");
                Assert.IsFalse(ChatTextContentView.UpdatePlainTextControl(control, control.Text));
                control.Dispose();
                Assert.IsFalse(ChatTextContentView.UpdatePlainTextControl(control, "ignored"));
            }
        }

        /// <summary>Checks final native text and layout in a nested activity hierarchy, including no-op snapshots.</summary>
        [STATestMethod]
        public void StreamedActivityTextSettlesItsNestedLayoutAndSkipsUnchangedSnapshots()
        {
            using (var group = new ChatActivityGroupView())
            using (var step = new ChatActivityStepView())
            {
                group.section.body.Controls.Add(step);
                group.section.Expanded = step.section.Expanded = true;
                var handle = step.detail.content.Handle;
                step.detail.ShowPlain("Seed");
                int layouts = 0;
                group.Layout += (s, e) => layouts++;
                string text = string.Concat(System.Linq.Enumerable.Repeat("A growing streamed activity line.\n", 100));
                step.detail.UpdatePlainText(text);
                Assert.AreEqual(text, step.detail.content.Text);
                Assert.AreEqual(1200, step.detail.content.Height);
                Assert.IsTrue(layouts > 0, "The settled layout must be published to the outer group.");
                int afterGrowth = layouts;
                step.detail.UpdatePlainText(text);
                Assert.AreEqual(afterGrowth, layouts, "An unchanged stream must not reflow the outer activity group.");
                step.detail.UpdatePlainText("Short result");
                Assert.IsTrue(step.detail.content.Height < 80);
                Assert.IsFalse(group.IsDisposed);
            }
        }

        [STATestMethod]
        public void ConstructorCreatesNamedThemedControls() { using (var theme = new ThemeScope()) using (var view = new ChatTextContentView()) { Assert.AreEqual("ChatTextContentView", view.Name); Assert.IsTrue(view.Controls.Count > 0); Assert.AreEqual(UiTheme.Background, view.BackColor); Assert.AreEqual(UiTheme.Foreground, view.ForeColor); } }
        [STATestMethod]
        public void NativeTextInteractionCopiesSelectionAndCodeAndHandlesMouseKeyboardAndMenu()
        {
            using (var theme = new ThemeScope())
            using (var view = new ChatTextContentView())
            {
                var copy = ChatMarkdown.CopyText; string copied = null, error = null; int activations = 0;
                try
                {
                    ChatMarkdown.CopyText = s => copied = s; view.ErrorHandler = s => error = s;
                    view.ShowPlain(null); Assert.AreEqual("", view.content.Text); view.ShowPlain("abcdef");
                    view.copySelection.PerformClick(); Assert.AreEqual("abcdef", copied);
                    view.content.Select(1, 2); view.copySelection.PerformClick(); Assert.AreEqual("bc", copied);
                    var handle = view.content.Handle; view.Append(null); Assert.AreEqual("abcdef", view.content.Text);
                    var first = view.OwnFont("Consolas", 8, FontStyle.Bold); Assert.AreSame(first, view.OwnFont("Consolas", 8, FontStyle.Bold));
                    Assert.AreNotSame(first, view.OwnFont("Segoe UI", 8, FontStyle.Bold)); Assert.AreNotSame(first, view.OwnFont("Consolas", 9, FontStyle.Bold)); Assert.AreNotSame(first, view.OwnFont("Consolas", 8, FontStyle.Italic));
                    view.actions.Add(new ChatTextContentView.TextAction { Start = 1, Length = 2, Invoke = () => activations++ });
                    view.ActivateAt(0); view.ActivateAt(1); view.ActivateAt(2); view.ActivateAt(3); Assert.AreEqual(2, activations);
                    view.actions.Add(new ChatTextContentView.TextAction { Start = 4, Length = 1 }); view.ActivateAt(4);
                    view.actions.Add(new ChatTextContentView.TextAction { Start = 5, Length = 1, Invoke = () => { throw new InvalidOperationException("action error"); } }); view.ActivateAt(5); Assert.AreEqual("action error", error);
                    view.content.Select(1, 0); var enter = new KeyEventArgs(Keys.Enter); TranscriptFixture.Event(view.content, "OnKeyDown", enter); Assert.IsTrue(enter.Handled); Assert.AreEqual(3, activations);
                    var other = new KeyEventArgs(Keys.A); TranscriptFixture.Event(view.content, "OnKeyDown", other); Assert.IsFalse(other.Handled);
                    var point = view.content.GetPositionFromCharIndex(1); TranscriptFixture.Event(view.content, "OnMouseUp", new MouseEventArgs(MouseButtons.Right, 1, point.X, point.Y, 0)); Assert.AreEqual(3, activations);
                    view.content.Select(1, 1); TranscriptFixture.Event(view.content, "OnMouseUp", new MouseEventArgs(MouseButtons.Left, 1, point.X, point.Y, 0)); Assert.AreEqual(3, activations);
                    view.content.Select(1, 0); TranscriptFixture.Event(view.content, "OnMouseUp", new MouseEventArgs(MouseButtons.Left, 1, point.X, point.Y, 0)); Assert.AreEqual(4, activations);
                    view.copyCode.PerformClick(); Assert.AreEqual("bc", copied);
                    UiInvoke.Call(typeof(ContextMenuStrip), "OnOpening", view.copyMenu, new CancelEventArgs()); Assert.IsFalse(view.copyCode.Available);
                    view.content.Select(0, 0); UiInvoke.Call(typeof(ToolStripItem), "OnClick", view.copyCode, EventArgs.Empty); UiInvoke.Call(typeof(ContextMenuStrip), "OnOpening", view.copyMenu, new CancelEventArgs()); Assert.IsFalse(view.copyCode.Available);
                    view.actions.Insert(0, new ChatTextContentView.TextAction { Start = 0, Length = 1, Code = "code" }); UiInvoke.Call(typeof(ContextMenuStrip), "OnOpening", view.copyMenu, new CancelEventArgs()); Assert.IsTrue(view.copyCode.Available);
                    view.copyCode.PerformClick(); Assert.AreEqual("code", copied); ChatMarkdown.CopyText = s => { throw new InvalidOperationException("copy error"); }; view.copySelection.PerformClick(); Assert.AreEqual("copy error", error);
                }
                finally { ChatMarkdown.CopyText = copy; }
            }
        }
        /// <summary>Markdown uses the visible shared blocks; the retained plain activity surface is untouched.</summary>
        [STATestMethod]
        public void MarkdownDefersHeightChangesUntilAllStyledRunsAreRendered()
        {
            using (var host = new MarkdownTestHost())
            using (var view = new ChatTextContentView { Dock = DockStyle.Top })
            {
                host.Controls.Add(view);
                view.ShowPlain("seed");
                var handle = view.content.Handle;
                int plainChanges = 0;
                view.content.TextChanged += (s, e) => plainChanges++;
                string markdown = string.Concat(System.Linq.Enumerable.Repeat("**Heading** and `code`\n\n", 100));
                view.ShowMarkdown(markdown, null, reference => { }, error => Assert.Fail(error));
                host.Show(); Application.DoEvents(); view.PerformLayout();
                var shared = view.MarkdownView;
                Assert.IsTrue(shared.Visible, "The actual shared Markdown surface must be displayed.");
                Assert.IsTrue(shared.IsHandleCreated);
                Assert.IsFalse(view.content.Visible);
                Assert.AreEqual(markdown, shared.SourceText);
                Assert.AreEqual(100, shared.ContentPanel.Controls.Count);
                var first = (Blackcancer.ChatUI.Views.ChatTextView)shared.ContentPanel.Controls[0];
                StringAssert.Contains(first.Content.Text, "Heading and code");
                first.Content.Select(0, 7);
                Assert.IsTrue(first.Content.SelectionFont.Bold);
                Assert.AreEqual(0, plainChanges, "Markdown must never be mirrored into a hidden legacy RichTextBox.");
                Assert.AreEqual("seed", view.content.Text);
                Assert.AreEqual(handle, view.content.Handle);
                view.ShowPlain("");
                Assert.IsFalse(shared.Visible); Assert.IsTrue(view.content.Visible);
                Assert.AreEqual(24, view.content.Height);
                Assert.AreEqual(RichTextBoxScrollBars.None, view.content.ScrollBars);
            }
        }

        private sealed class MarkdownTestHost : Form
        {
            internal MarkdownTestHost()
            { StartPosition = FormStartPosition.Manual; Location = new Point(-30000, -30000); ShowInTaskbar = false; ClientSize = new Size(400, 900); }
            protected override bool ShowWithoutActivation => true;
        }

        [STATestMethod]
        public void NativeViewportResizeCannotInflateOnePublishedLineIntoLargeBlankSpace()
        {
            using (var form = new Form { Left = -10000, Top = -10000, ShowInTaskbar = false, ClientSize = new Size(660, 300) })
            using (var view = new ChatTextContentView { Dock = DockStyle.Top })
            {
                form.Controls.Add(view); view.ShowPlain("Published public summary");
                form.Show(); Application.DoEvents(); view.PerformLayout();
                int naturalHeight = view.content.Height;
                Assert.IsTrue(naturalHeight >= 24 && naturalHeight < 60);
                TranscriptFixture.Event(view.content, "OnContentsResized", new ContentsResizedEventArgs(new Rectangle(0, 0, 456, 1024)));
                Application.DoEvents();
                Assert.AreEqual(naturalHeight, view.content.Height, "A previous native viewport height must not displace later tool rows.");
                Assert.AreEqual("Published public summary", view.content.Text);
                Assert.AreEqual(RichTextBoxScrollBars.None, view.content.ScrollBars);
            }
        }

        [STATestMethod]
        public void RichTextReflowBoundsHeightAndSupportsEmptyDisposedOrInitializingContent()
        {
            using (var theme = new ThemeScope())
            using (var culture = new LocalizationScope())
            using (var view = new ChatTextContentView())
            {
                UiInvoke.Call(typeof(ChatTextContentView), "ResizeText", view); view.ShowPlain("without handle");
                view.ShowPlain("small"); var handle = view.content.Handle; view.Width = 200; Assert.IsTrue(view.content.Height >= 24); Assert.AreEqual(RichTextBoxScrollBars.None, view.content.ScrollBars);
                view.ShowPlain(string.Join("\n", new string[300]).Replace("\n", "line\n")); Assert.AreEqual(1200, view.content.Height); Assert.AreEqual(RichTextBoxScrollBars.Vertical, view.content.ScrollBars);
                TranscriptFixture.Event(view.content, "OnContentsResized", new ContentsResizedEventArgs(new Rectangle(0, 0, 10, 1))); Assert.AreEqual(1200, view.content.Height, "A stale native viewport rectangle must not collapse actual long content.");
                TranscriptFixture.Event(view.content, "OnContentsResized", new ContentsResizedEventArgs(new Rectangle(0, 0, 10, 2000))); Assert.AreEqual(1200, view.content.Height);
                view.ShowPlain(""); Assert.AreEqual("", view.content.Text); LocalizationScope.Set("ar-SA"); view.ShowPlain("rtl"); Assert.AreEqual(RightToLeft.Yes, view.content.RightToLeft);
                view.ShowPlain("code", true); Assert.AreEqual(RightToLeft.No, view.content.RightToLeft); Assert.AreEqual("Consolas", view.content.Font.FontFamily.Name);
                var content = view.content; view.content = null; UiInvoke.Call(typeof(ChatTextContentView), "ResizeText", view); view.content = content;
                content.Dispose(); UiInvoke.Call(typeof(ChatTextContentView), "ResizeText", view); Assert.IsTrue(content.IsDisposed);
            }
        }
    }
}
