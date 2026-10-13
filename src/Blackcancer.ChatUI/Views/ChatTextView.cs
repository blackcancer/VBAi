// SPDX-License-Identifier: MPL-2.0
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Blackcancer.ChatUI.Presentation;

namespace Blackcancer.ChatUI.Views
{
    /// <summary>Selectable prose with RichEdit document measurement independent of caret scrolling.</summary>
    public partial class ChatTextView : ChatView
    {
        private readonly List<Font> fonts = new List<Font>();
        private readonly List<TextAction> actions = new List<TextAction>();
        private string renderedSource;
        private string renderedPrefix;
        private int renderedListLevel, renderedQuoteDepth;
        private int updateDepth, requestedHeight = 24, measuredWidth = -1;
        private bool measuring;
        private ChatServices services = new ChatServices();
        private SharedChatStyle style = new SharedChatStyle();
        private IDictionary<string, Action> tokenActions;
        private Action<string> error = message => DebugError(message);
        private sealed class TextAction { internal int Start, Length; internal bool IsToken; internal Action Invoke; }
        public ChatTextView()
        {
            InitializeComponent();
            content.ContentsResized += (s, e) => { requestedHeight = Math.Max(24, Math.Min(MaximumTextHeight, e.NewRectangle.Height + 8)); if (!measuring && updateDepth == 0) MeasureText(); };
            content.HandleCreated += (s, e) => MeasureText();
            content.MouseUp += (s, e) => { if (e.Button == MouseButtons.Left && content.SelectionLength == 0) ActivateAt(content.GetCharIndexFromPosition(e.Location)); };
            content.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { ActivateAt(content.SelectionStart); e.Handled = true; } };
            copySelection.Click += (s, e) => InvokeSafely(() => services.CopyText?.Invoke(content.SelectionLength == 0 ? content.Text : content.SelectedText));
            BindStyle();
        }
        public RichTextBox Content => content;
        public int MaximumTextHeight { get; set; } = 1200;
        public ChatServices Services { get => services; set { if (ReferenceEquals(services, value)) return; services = value ?? new ChatServices(); BindStyle(); } }
        public SharedChatStyle Style { get => style; set { style = value ?? new SharedChatStyle(); renderedSource = null; BindStyle(); } }
        internal void Configure(Action<string> onError, IDictionary<string, Action> tokens)
        { error = onError ?? new Action<string>(DebugError); tokenActions = tokens; RebuildTokenActions(); }
        public void BindStyle()
        {
            BackColor = content.BackColor = style.Background; ForeColor = content.ForeColor = style.Foreground;
            content.Font = style.BodyFont ?? SystemFonts.MessageBoxFont;
            content.RightToLeft = style.RightToLeft ? RightToLeft.Yes : RightToLeft.No;
            copySelection.Text = services.Text("Copy");
        }
        internal void ShowBlock(ChatMarkdownBlock block)
        {
            if (renderedSource == block.Source && renderedPrefix == block.ListPrefix && renderedListLevel == block.ListLevel && renderedQuoteDepth == block.QuoteDepth) return;
            int selectionStart = content.SelectionStart, selectionLength = content.SelectionLength;
            NativePoint scroll = CaptureScroll();
            updateDepth++;
            try
            {
                actions.Clear(); content.Clear();
                string prefix = block.Kind == ChatMarkdownBlockKind.ListItem ? new string(' ', Math.Min(24, Math.Max(0, block.ListLevel - 1) * 2)) + block.ListPrefix : block.QuoteDepth > 0 ? new string('│', Math.Min(12, block.QuoteDepth)) + " " : "";
                if (block.Kind == ChatMarkdownBlockKind.Rule) Append("────────", FontStyle.Regular, false, 0, null);
                else
                {
                    if (prefix.Length > 0) Append(prefix, FontStyle.Regular, false, 0, null);
                    foreach (var span in block.Spans)
                    {
                        int start = content.TextLength;
                        var fontStyle = span.FontStyle | (block.Kind == ChatMarkdownBlockKind.Heading ? FontStyle.Bold : FontStyle.Regular);
                        float size = block.Kind == ChatMarkdownBlockKind.Heading ? Math.Max(content.Font.Size, 18 - block.HeadingLevel) : 0;
                        bool link = services.IsAllowedLink?.Invoke(span.Href) == true;
                        Append(span.Text, fontStyle | (link ? FontStyle.Underline : 0), span.IsCode, size, link ? (Color?)style.Link : null);
                        if (link) { string href = span.Href; actions.Add(new TextAction { Start = start, Length = span.Text.Length, Invoke = () => services.OpenLink?.Invoke(href) }); }
                    }
                }
                renderedSource = block.Source; renderedPrefix = block.ListPrefix; renderedListLevel = block.ListLevel; renderedQuoteDepth = block.QuoteDepth; RebuildTokenActions();
                content.Select(Math.Min(selectionStart, content.TextLength), Math.Min(selectionLength, Math.Max(0, content.TextLength - selectionStart)));
                RestoreScroll(scroll);
            }
            finally { updateDepth--; MeasureText(); }
        }
        public void ShowPlain(string text, bool code = false)
        {
            text = text ?? "";
            if (renderedSource == text && content.Text == text) return;
            int start = content.SelectionStart, length = content.SelectionLength;
            NativePoint scroll = CaptureScroll();
            updateDepth++;
            try
            {
                actions.Clear();
                content.Font = code ? style.CodeFont ?? OwnFont("Consolas", content.Font.Size, FontStyle.Regular) : style.BodyFont ?? SystemFonts.MessageBoxFont;
                content.RightToLeft = code ? RightToLeft.No : style.RightToLeft ? RightToLeft.Yes : RightToLeft.No;
                if (text.StartsWith(content.Text, StringComparison.Ordinal)) { content.Select(content.TextLength, 0); content.AppendText(text.Substring(content.TextLength)); }
                else content.Text = text;
                renderedSource = text; RebuildTokenActions();
                content.Select(Math.Min(start, content.TextLength), Math.Min(length, Math.Max(0, content.TextLength - start)));
                RestoreScroll(scroll);
            }
            finally { updateDepth--; MeasureText(); }
        }
        private void Append(string text, FontStyle fontStyle, bool code, float size, Color? color)
        {
            Font basis = code ? style.CodeFont : style.BodyFont;
            string family = basis?.FontFamily.Name ?? (code ? "Consolas" : content.Font.FontFamily.Name);
            content.Select(content.TextLength, 0);
            content.SelectionFont = OwnFont(family, size > 0 ? size : basis?.Size ?? content.Font.Size, fontStyle);
            content.SelectionColor = color ?? style.Foreground;
            content.AppendText(text ?? "");
        }
        private Font OwnFont(string family, float size, FontStyle fontStyle)
        { var font = fonts.Find(f => f.FontFamily.Name == family && f.Size == size && f.Style == fontStyle); if (font != null) return font; font = new Font(family, size, fontStyle); fonts.Add(font); return font; }
        private void RebuildTokenActions()
        {
            actions.RemoveAll(a => a.IsToken);
            // Host tokens are matched literally, longest first, without changing rendered text.
            if (tokenActions == null) return;
            var occupied = new List<Tuple<int, int>>();
            foreach (var token in tokenActions.OrderByDescending(t => t.Key?.Length ?? 0))
            {
                if (string.IsNullOrEmpty(token.Key) || token.Value == null) continue;
                int offset = 0;
                while ((offset = content.Text.IndexOf(token.Key, offset, StringComparison.Ordinal)) >= 0)
                {
                    int end = offset + token.Key.Length;
                    if (!occupied.Any(range => offset < range.Item2 && end > range.Item1))
                    { occupied.Add(Tuple.Create(offset, end)); var callback = token.Value; actions.Insert(0, new TextAction { Start = offset, Length = token.Key.Length, IsToken = true, Invoke = callback }); }
                    offset = end;
                }
            }
        }
        private void ActivateAt(int position) => InvokeSafely(() => actions.Find(a => position >= a.Start && position < a.Start + a.Length)?.Invoke?.Invoke());
        private void InvokeSafely(Action action) { try { action(); } catch (Exception ex) { error(ex.Message); } }
        private static void DebugError(string value) => System.Diagnostics.Debug.WriteLine(value);
        protected override void OnSizeChanged(EventArgs e)
        { base.OnSizeChanged(e); if (measuredWidth == ClientSize.Width) return; measuredWidth = ClientSize.Width; MeasureText(); }
        private void MeasureText()
        {
            if (measuring || updateDepth != 0 || content == null || content.IsDisposed) return;
            measuring = true;
            try
            {
                if (content.IsHandleCreated) SendMessage(content.Handle, 0x0441 /* EM_REQUESTRESIZE */, IntPtr.Zero, IntPtr.Zero);
                else requestedHeight = Math.Max(24, Math.Min(MaximumTextHeight, content.GetPreferredSize(new Size(Math.Max(1, content.Width), 0)).Height + 8));
                int height = content.TextLength == 0 ? 24 : Math.Min(MaximumTextHeight, requestedHeight);
                if (content.Height != height) content.Height = height;
            }
            finally { measuring = false; }
        }
        private NativePoint CaptureScroll() { var point = new NativePoint(); if (content.IsHandleCreated) SendMessagePoint(content.Handle, 0x04DD, IntPtr.Zero, ref point); return point; }
        private void RestoreScroll(NativePoint point) { if (content.IsHandleCreated) SendMessagePoint(content.Handle, 0x04DE, IntPtr.Zero, ref point); }
        [StructLayout(LayoutKind.Sequential)] private struct NativePoint { public int X, Y; }
        [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr window, int message, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll", EntryPoint = "SendMessageW")] private static extern IntPtr SendMessagePoint(IntPtr window, int message, IntPtr wParam, ref NativePoint point);
        private void DisposeFonts() { foreach (var font in fonts) font.Dispose(); fonts.Clear(); }
    }
}
