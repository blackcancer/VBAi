using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
namespace VBAi
{

    /// <summary>Displays selectable transcript text and Markdown with clickable references and code copying.</summary>
    public sealed partial class ChatTextContentView : ChatDesignerView
    {

        /// <summary>Receives errors raised while activating transcript links or actions.</summary>
        internal Action<string> ErrorHandler = LoadLog.Write;

        /// <summary>Tracks fonts created by this view so they can be disposed with it.</summary>
        private readonly List<Font> ownedFonts = new List<Font>();

        /// <summary>Suppresses intermediate resize work while Markdown or plain content is being replaced.</summary>
        private int contentUpdateDepth;

        /// <summary>Prevents reentrant measurement while text height or native scrollbar state changes.</summary>
        private bool resizingText;

        /// <summary>Maps rendered character ranges to navigation, link, and code copy actions.</summary>
        internal readonly List<TextAction> actions = new List<TextAction>();

        /// <summary>Character range and callback associated with a rendered transcript link or code-copy action.</summary>
        internal sealed class TextAction {

/// <summary>Start index and number of rendered characters covered by this action.</summary>
internal int Start, Length;

/// <summary>Callback invoked when the user activates a character within the range.</summary>
internal Action Invoke;

/// <summary>Optional source text copied by the code-copy menu for this action.</summary>
internal string Code; }

        /// <summary>Creates the native read-only text field and context menu.</summary>
        public ChatTextContentView()
        {
            InitializeComponent(); UiText.Apply(this, components); UiTheme.Apply(this);
            content.ContentsResized += (s,e) => { if (contentUpdateDepth == 0 && !resizingText) content.Height = content.TextLength == 0 ? 24 : Math.Max(24, Math.Min(1200, e.NewRectangle.Height + 8)); };
            content.TextChanged += (s,e) => ResizeText();
            content.HandleCreated += (s,e) => ResizeText();
            content.MouseUp += (s,e) => { if (e.Button != MouseButtons.Left || content.SelectionLength != 0) return; ActivateAt(content.GetCharIndexFromPosition(e.Location)); };
            content.KeyDown += (s,e) => { if (e.KeyCode == Keys.Enter) { ActivateAt(content.SelectionStart); e.Handled = true; } };
            copySelection.Click += (s,e) => Copy(content.SelectionLength > 0 ? content.SelectedText : content.Text);
            copyCode.Click += (s,e) => { var action = ActionAt(content.SelectionStart); if (action?.Code != null) Copy(action.Code); };
            copyMenu.Opening += (s,e) => copyCode.Visible = ActionAt(content.SelectionStart)?.Code != null;
        }

        /// <summary>Replaces the transcript content with plain text and applies code or interface formatting.</summary>
        /// <param name="text">Plain message text to display.</param>
        /// <param name="code">Whether to use code formatting for the text.</param>
        internal void ShowPlain(string text, bool code = false)
        {
            contentUpdateDepth++;
            try
            {
                actions.Clear(); content.Clear();
                content.RightToLeft = !code && UiText.Culture.TextInfo.IsRightToLeft ? RightToLeft.Yes : RightToLeft.No;
                content.Font = OwnFont(code ? "Consolas" : "Segoe UI", 9.5f, FontStyle.Regular);
                content.Text = text ?? ""; content.Select(0,0);
            }
            finally { contentUpdateDepth--; ResizeText(); }
        }

        /// <summary>Renders Markdown into the transcript and installs the callbacks for references and errors.</summary>
        /// <param name="text">Markdown source to render in the transcript.</param>
        /// <param name="references">VBA references that should become interactive transcript links.</param>
        /// <param name="navigate">Callback used to open a recognized VBA reference.</param>
        /// <param name="error">Callback used to report link and action errors.</param>
        internal void ShowMarkdown(string text, IDictionary<string,VbeChatReference> references, Action<VbeChatReference> navigate, Action<string> error)
        {
            contentUpdateDepth++;
            try
            {
                ErrorHandler = error;
                content.RightToLeft = UiText.Culture.TextInfo.IsRightToLeft ? RightToLeft.Yes : RightToLeft.No;
                actions.Clear(); content.Clear();
                ChatNativeMarkdown.Render(this, text ?? "", references, navigate, error);
                content.Select(0,0);
            }
            finally { contentUpdateDepth--; ResizeText(); }
        }

        /// <summary>Returns a cached view owned font matching the requested family, size, and style.</summary>
        /// <param name="family">Font family to reuse or create.</param>
        /// <param name="size">Font size in points.</param>
        /// <param name="style">Font style to apply.</param>
        /// <returns>The matching font instance, created and retained by this view when necessary.</returns>
        internal Font OwnFont(string family, float size, FontStyle style) { var existing = ownedFonts.Find(f => f.FontFamily.Name == family && f.Size == size && f.Style == style); if (existing != null) return existing; var font = new Font(family,size,style); ownedFonts.Add(font); return font; }

        /// <summary>Appends a styled text run to the native transcript control.</summary>
        /// <param name="text">Text to append to the transcript.</param>
        /// <param name="family">Font family for the appended text.</param>
        /// <param name="size">Font size in points.</param>
        internal void Append(string text, string family = "Segoe UI", float size = 9.5f) => Append(text, family, size, FontStyle.Regular);

        /// <summary>Appends a styled text run to the native transcript control.</summary>
        /// <param name="text">Text to append to the transcript.</param>
        /// <param name="family">Font family for the appended text.</param>
        /// <param name="size">Font size in points.</param>
        /// <param name="style">Font style for the appended text.</param>
        /// <param name="color">Optional foreground color for the appended text.</param>
        internal void Append(string text, string family, float size, FontStyle style, Color? color = null)
        {
            content.Select(content.TextLength,0); content.SelectionFont = OwnFont(family,size,style);
            content.SelectionColor = color ?? UiTheme.Foreground; content.AppendText(text ?? "");
        }

        /// <summary>Invokes the transcript action associated with the character at the specified position.</summary>
        /// <param name="index">Character position whose associated action should be invoked.</param>
        internal void ActivateAt(int index) { try { ActionAt(index)?.Invoke?.Invoke(); } catch (Exception ex) { ErrorHandler(ex.Message); } }

        /// <summary>Copies the supplied transcript text through the native clipboard helper.</summary>
        /// <param name="text">Text to place on the clipboard.</param>
        private void Copy(string text) { try { ChatMarkdown.CopyText(text); } catch (Exception ex) { ErrorHandler(ex.Message); } }

        /// <summary>Finds the action whose rendered character range contains the specified position.</summary>
        /// <param name="index">Character position to test against the rendered action ranges.</param>
        /// <returns>The action covering the character position, or null when no action covers it.</returns>
        private TextAction ActionAt(int index) => actions.Find(a => index >= a.Start && index < a.Start + a.Length);

        /// <summary>Reflows the native rich text on width changes.</summary>
        /// <param name="e">Resize event.</param>
        protected override void OnSizeChanged(EventArgs e) { base.OnSizeChanged(e); ResizeText(); }

        /// <summary>Measures the rendered text and updates the rich text control height and scroll bars.</summary>
        private void ResizeText()
        {
            if (resizingText || contentUpdateDepth != 0 || content == null || content.IsDisposed || !content.IsHandleCreated) return;
            resizingText = true;
            try
            {
                int height = 24;
                if (content.TextLength > 0)
                {
                    var end = content.GetPositionFromCharIndex(content.TextLength - 1);
                    height = Math.Max(24, Math.Min(1200, end.Y + (int)Math.Ceiling(content.Font.GetHeight()) + 12));
                }
                content.Height = height;
                // ScrollBars recreates the native handle and synchronously raises HandleCreated.
                // Never start another measurement while that recreation is still in progress.
                var scrollBars = height >= 1200 ? RichTextBoxScrollBars.Vertical : RichTextBoxScrollBars.None;
                if (content.ScrollBars != scrollBars) content.ScrollBars = scrollBars;
            }
            finally { resizingText = false; }
        }

        /// <summary>Disposes fonts created by this view and clears its font cache.</summary>
        private void DisposeTextResources() { foreach (var font in ownedFonts) font.Dispose(); ownedFonts.Clear(); }
    }
}
