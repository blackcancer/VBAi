// SPDX-License-Identifier: MPL-2.0
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;
using Blackcancer.ChatUI.Presentation;

namespace Blackcancer.ChatUI.Views
{
    /// <summary>Designer-backed Markdown transcript. Only changed blocks are rendered again.</summary>
    public partial class ChatMarkdownView : ChatView
    {
        private sealed class Entry { internal ChatMarkdownBlock Block; internal Control View; }
        private List<Entry> entries = new List<Entry>();
        private ChatServices services = new ChatServices();
        private SharedChatStyle style = new SharedChatStyle();
        private Action<string> error = text => System.Diagnostics.Debug.WriteLine(text);
        private IDictionary<string, Action> tokenActions;
        private string source = "";
        private bool rebinding;
        public ChatMarkdownView()
        {
            InitializeComponent();
            blocks.SizeChanged += (sender, args) => FitBlocks();
            copySource.Click += (sender, args) => CopyMarkdown();
            BindStyle();
        }
        public string SourceText => source;
        /// <summary>Optional complete-source copy callback. The per-instance service is used by default.</summary>
        [Browsable(false)] public Action<string> CopySource { get; set; }
        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Content)] public FlowLayoutPanel ContentPanel => blocks;
        public ChatServices Services
        { get => services; set { services = value ?? new ChatServices(); foreach (var entry in entries) Configure(entry.View); } }
        public SharedChatStyle Style
        {
            get => style;
            set { style = value ?? new SharedChatStyle(); rebinding = true; try { BindStyle(); foreach (var entry in entries) { Configure(entry.View); UpdateView(entry); } } finally { rebinding = false; } }
        }
        public void BindStyle()
        { BackColor = blocks.BackColor = style.Background; ForeColor = blocks.ForeColor = style.Foreground; Font = style.BodyFont; RightToLeft = style.RightToLeft ? RightToLeft.Yes : RightToLeft.No; copySource.Text = services.Text("Copy Markdown"); }
        public void ShowMarkdown(string text, Action<string> onError, IDictionary<string, Action> actions = null)
        {
            error = onError ?? (message => System.Diagnostics.Debug.WriteLine(message)); tokenActions = actions;
            foreach (var entry in entries) Configure(entry.View);
            UpdateMarkdown(text);
        }
        public void AppendDelta(string delta) { if (!string.IsNullOrEmpty(delta)) UpdateMarkdown(source + delta); }
        public void CompleteMarkdown(string text) => UpdateMarkdown(text);
        public void UpdateMarkdown(string text)
        {
            text = text ?? "";
            if (source == text) return;
            ChatMarkdownDocument document;
            try { document = ChatMarkdownEngine.Parse(text); }
            catch (Exception ex) { error(ex.Message); return; }
            var next = new List<Entry>(document.Blocks.Count);
            var used = new HashSet<Entry>();
            var identities = new Dictionary<Tuple<ChatMarkdownBlockKind, string, string, int, int>, Queue<Entry>>();
            foreach (var entry in entries)
            {
                var key = Identity(entry.Block); Queue<Entry> queue;
                if (!identities.TryGetValue(key, out queue)) { queue = new Queue<Entry>(); identities.Add(key, queue); }
                queue.Enqueue(entry);
            }
            // Prefer exact identities across insertions/removals; keep changed blocks of the same kind in place.
            foreach (var block in document.Blocks)
            {
                Queue<Entry> candidates;
                var entry = identities.TryGetValue(Identity(block), out candidates) && candidates.Count > 0 ? candidates.Dequeue() : null;
                if (entry != null) { used.Add(entry); next.Add(entry); }
                else next.Add(new Entry { Block = block });
            }
            for (int index = 0; index < next.Count; index++)
            {
                if (next[index].View != null) continue;
                var block = next[index].Block;
                var entry = index < entries.Count && !used.Contains(entries[index]) && entries[index].Block.Kind == block.Kind ? entries[index] : null;
                if (entry == null && block.Kind == ChatMarkdownBlockKind.Code)
                    entry = entries.FirstOrDefault(e => !used.Contains(e) && e.Block.Kind == block.Kind && e.Block.Language == block.Language);
                if (entry != null) { used.Add(entry); entry.Block = block; next[index] = entry; }
            }
            var scrollHost = FindScrollHost();
            var scroll = scrollHost?.AutoScrollPosition ?? System.Drawing.Point.Empty;
            blocks.SuspendLayout(); SuspendLayout();
            try
            {
                foreach (var entry in entries.Where(e => !used.Contains(e))) { blocks.Controls.Remove(entry.View); entry.View.Dispose(); }
                for (int index = 0; index < next.Count; index++)
                {
                    var entry = next[index];
                    if (entry.View == null) { entry.View = CreateView(entry.Block); Configure(entry.View); }
                    UpdateView(entry);
                    if (entry.View.Parent != blocks) blocks.Controls.Add(entry.View);
                    if (blocks.Controls.GetChildIndex(entry.View) != index) blocks.Controls.SetChildIndex(entry.View, index);
                }
                entries = next; source = document.Source; FitBlocks();
            }
            finally
            {
                blocks.ResumeLayout(true); ResumeLayout(true);
                if (scrollHost != null) scrollHost.AutoScrollPosition = new System.Drawing.Point(-scroll.X, -scroll.Y);
            }
        }
        private Control CreateView(ChatMarkdownBlock block)
        {
            if (block.Kind == ChatMarkdownBlockKind.Code) return new ChatCodeBlockView();
            if (block.Kind == ChatMarkdownBlockKind.Table) return new ChatTableView();
            return new ChatTextView();
        }
        private static Tuple<ChatMarkdownBlockKind, string, string, int, int> Identity(ChatMarkdownBlock block)
            => Tuple.Create(block.Kind, block.Source, block.ListPrefix, block.ListLevel, block.QuoteDepth);
        private void Configure(Control view)
        {
            var text = view as ChatTextView; if (text != null) { text.Services = services; if (!ReferenceEquals(text.Style, style) || rebinding) text.Style = style; text.Configure(error, tokenActions); }
            var code = view as ChatCodeBlockView; if (code != null) { code.Services = services; if (!ReferenceEquals(code.Style, style) || rebinding) code.Style = style; }
            var table = view as ChatTableView; if (table != null) { table.Services = services; if (!ReferenceEquals(table.Style, style) || rebinding) table.Style = style; }
        }
        private void UpdateView(Entry entry)
        {
            var text = entry.View as ChatTextView; if (text != null) text.ShowBlock(entry.Block);
            var code = entry.View as ChatCodeBlockView; if (code != null) code.ShowCode(entry.Block.CodeText, entry.Block.Language, error);
            var table = entry.View as ChatTableView; if (table != null) table.ShowTable(entry.Block, error);
        }
        protected override void OnSizeChanged(EventArgs e) { base.OnSizeChanged(e); FitBlocks(); }
        private void FitBlocks()
        {
            if (blocks == null) return;
            int width = Math.Max(1, blocks.ClientSize.Width - blocks.Padding.Horizontal);
            foreach (Control view in blocks.Controls) { int next = Math.Max(1, width - view.Margin.Horizontal); if (view.Width != next) view.Width = next; }
        }
        private ScrollableControl FindScrollHost()
        { for (Control control = Parent; control != null; control = control.Parent) { var candidate = control as ScrollableControl; if (candidate != null && candidate.AutoScroll) return candidate; } return null; }
        private void CopyMarkdown()
        { try { (CopySource ?? services.CopyText)?.Invoke(source); } catch (Exception ex) { error(ex.Message); } }
    }
}
