using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
namespace CodexVBE
{
    /// <summary>Designer-backed selectable rich text, with native links and code copying.</summary>
    public sealed partial class ChatTextContentView : ChatDesignerView
    {
        internal Action<string> ErrorHandler = LoadLog.Write;
        private readonly List<Font> ownedFonts = new List<Font>();
        internal readonly List<TextAction> actions = new List<TextAction>();
        internal sealed class TextAction { internal int Start, Length; internal Action Invoke; internal string Code; }
        /// <summary>Creates the native read-only text field and context menu.</summary>
        public ChatTextContentView()
        {
            InitializeComponent(); UiText.Apply(this, components); UiTheme.Apply(this);
            content.ContentsResized += (s,e) => content.Height = Math.Max(24, Math.Min(1200, e.NewRectangle.Height + 8));
            content.TextChanged += (s,e) => ResizeText();
            content.HandleCreated += (s,e) => ResizeText();
            content.MouseUp += (s,e) => { if (e.Button != MouseButtons.Left || content.SelectionLength != 0) return; ActivateAt(content.GetCharIndexFromPosition(e.Location)); };
            content.KeyDown += (s,e) => { if (e.KeyCode == Keys.Enter) { ActivateAt(content.SelectionStart); e.Handled = true; } };
            copySelection.Click += (s,e) => Copy(content.SelectionLength > 0 ? content.SelectedText : content.Text);
            copyCode.Click += (s,e) => { var action = ActionAt(content.SelectionStart); if (action?.Code != null) Copy(action.Code); };
            copyMenu.Opening += (s,e) => copyCode.Visible = ActionAt(content.SelectionStart)?.Code != null;
        }
        internal void ShowPlain(string text, bool code = false)
        {
            actions.Clear(); content.Clear();
            content.RightToLeft = !code && UiText.Culture.TextInfo.IsRightToLeft ? RightToLeft.Yes : RightToLeft.No;
            content.Font = OwnFont(code ? "Consolas" : "Segoe UI", 9.5f, FontStyle.Regular);
            content.Text = text ?? ""; content.Select(0,0); ResizeText();
        }
        internal void ShowMarkdown(string text, IDictionary<string,VbeChatReference> references, Action<VbeChatReference> navigate, Action<string> error)
        {
            ErrorHandler = error;
            content.RightToLeft = UiText.Culture.TextInfo.IsRightToLeft ? RightToLeft.Yes : RightToLeft.No;
            actions.Clear(); content.Clear();
            ChatNativeMarkdown.Render(this, text ?? "", references, navigate, error);
            content.Select(0,0); ResizeText();
        }
        internal Font OwnFont(string family, float size, FontStyle style) { var existing = ownedFonts.Find(f => f.FontFamily.Name == family && f.Size == size && f.Style == style); if (existing != null) return existing; var font = new Font(family,size,style); ownedFonts.Add(font); return font; }
        internal void Append(string text, string family = "Segoe UI", float size = 9.5f, FontStyle style = FontStyle.Regular, Color? color = null)
        {
            content.Select(content.TextLength,0); content.SelectionFont = OwnFont(family,size,style);
            content.SelectionColor = color ?? UiTheme.Foreground; content.AppendText(text ?? "");
        }
        internal void ActivateAt(int index) { try { ActionAt(index)?.Invoke?.Invoke(); } catch (Exception ex) { ErrorHandler(ex.Message); } }
        private void Copy(string text) { try { ChatMarkdown.CopyText(text); } catch (Exception ex) { ErrorHandler(ex.Message); } }
        private TextAction ActionAt(int index) => actions.Find(a => index >= a.Start && index < a.Start + a.Length);
        /// <summary>Reflows the native rich text on width changes.</summary>
        /// <param name="e">Resize event.</param>
        protected override void OnSizeChanged(EventArgs e) { base.OnSizeChanged(e); ResizeText(); }
        private void ResizeText()
        {
            if (content == null || content.IsDisposed || !content.IsHandleCreated || content.TextLength == 0) return;
            var end = content.GetPositionFromCharIndex(content.TextLength - 1);
            content.Height = Math.Max(24, Math.Min(1200, end.Y + (int)Math.Ceiling(content.Font.GetHeight()) + 12));
            content.ScrollBars = content.Height >= 1200 ? RichTextBoxScrollBars.Vertical : RichTextBoxScrollBars.None;
        }
        private void DisposeTextResources() { foreach (var font in ownedFonts) font.Dispose(); ownedFonts.Clear(); }
    }
}
