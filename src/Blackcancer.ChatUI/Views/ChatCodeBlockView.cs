// SPDX-License-Identifier: MPL-2.0
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Blackcancer.ChatUI.Views
{
    /// <summary>Code with a persistent disclosure choice, language caption and always visible copy action.</summary>
    public partial class ChatCodeBlockView : ChatView
    {
        private ChatServices services = new ChatServices();
        private SharedChatStyle style = new SharedChatStyle();
        private string source = "", language = "";
        private bool initialized, expanded = true, expansionChosen;
        private Action<string> error = value => System.Diagnostics.Debug.WriteLine(value);
        public ChatCodeBlockView()
        {
            InitializeComponent();
            toggle.Click += (s, e) => Expanded = !Expanded;
            copy.Click += (s, e) => { try { services.CopyText?.Invoke(source); } catch (Exception ex) { error(ex.Message); } };
            code.Content.WordWrap = false; code.Content.ScrollBars = RichTextBoxScrollBars.Both; code.MaximumTextHeight = 420;
            BindStyle();
        }
        public string SourceText => source;
        public RichTextBox Content => code.Content;
        public bool Expanded
        {
            get => expanded;
            set { expansionChosen = true; if (expanded == value) return; expanded = value; if (code.Visible != value) code.Visible = value; UpdateCaption(); ExpansionChanged?.Invoke(this, EventArgs.Empty); }
        }
        public event EventHandler ExpansionChanged;
        public ChatServices Services { get => services; set { if (ReferenceEquals(services, value)) return; services = value ?? new ChatServices(); code.Services = services; BindStyle(); } }
        public SharedChatStyle Style { get => style; set { style = value ?? new SharedChatStyle(); code.Style = style; BindStyle(); if (initialized) { code.ShowPlain(source, true); Colorize(); } } }
        public void BindStyle()
        {
            BackColor = header.BackColor = style.CodeBackground; ForeColor = style.Foreground;
            toggle.BackColor = copy.BackColor = style.CodeBackground; toggle.ForeColor = copy.ForeColor = style.Foreground;
            Font = style.BodyFont; code.BindStyle(); code.Content.BackColor = style.CodeBackground;
            copy.Text = services.Text("Copy"); UpdateCaption();
        }
        public void ShowCode(string text, string codeLanguage = null, Action<string> onError = null)
        {
            text = text ?? ""; codeLanguage = codeLanguage ?? "";
            error = onError ?? (value => System.Diagnostics.Debug.WriteLine(value));
            if (initialized && source == text && language == codeLanguage) return;
            source = text; language = codeLanguage; initialized = true;
            // Streaming may cross the long-code threshold before a user has chosen a state.
            if (!expansionChosen) expanded = CountLines(text) <= 12;
            code.Configure(error, null); code.ShowPlain(text, true);
            Colorize();
            if (code.Visible != expanded) code.Visible = expanded;
            UpdateCaption();
        }
        private void UpdateCaption()
        { string caption = (expanded ? "▾ " : "▸ ") + (language.Length == 0 ? services.Text("Code") : language); if (toggle.Text != caption) toggle.Text = caption; }
        private static int CountLines(string text) { int count = 1; foreach (char value in text) if (value == '\n') count++; return count; }
        private void Colorize()
        {
            var colorizer = services.ColorizeCode;
            if (colorizer == null) return;
            int start = Content.SelectionStart, length = Content.SelectionLength;
            var scroll = new NativePoint();
            if (Content.IsHandleCreated) SendMessagePoint(Content.Handle, 0x04DD, IntPtr.Zero, ref scroll);
            try { colorizer(Content, source, language); }
            catch (Exception ex) { error(ex.Message); }
            finally
            {
                start = Math.Min(start, Content.TextLength);
                Content.Select(start, Math.Min(length, Content.TextLength - start));
                if (Content.IsHandleCreated) SendMessagePoint(Content.Handle, 0x04DE, IntPtr.Zero, ref scroll);
            }
        }
        [StructLayout(LayoutKind.Sequential)]
        private struct NativePoint { public int X, Y; }
        [DllImport("user32.dll", EntryPoint = "SendMessage", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessagePoint(IntPtr window, int message, IntPtr wParam, ref NativePoint point);
    }
}
