// SPDX-License-Identifier: MPL-2.0
using System;
using System.ComponentModel;
using System.Windows.Forms;

namespace Blackcancer.ChatUI.Views
{
    /// <summary>Fixed Designer template for collapsible content; title updates do not relayout the body.</summary>
    [Designer(typeof(ChatDisclosureDesigner))]
    public partial class ChatDisclosureView : ChatView
    {
        private bool expanded;
        private string title = "Details";
        private ChatServices services = new ChatServices();
        private SharedChatStyle style = new SharedChatStyle();
        public ChatDisclosureView() { InitializeComponent(); toggle.Click += (s, e) => Expanded = !Expanded; UpdateExpansion(); BindStyle(); }
        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Content)]
        public FlowLayoutPanel ContentPanel => body;
        [Category("Appearance")]
        public string Title { get => title; set { if (title == value) return; title = value ?? ""; UpdateCaption(); } }
        [DefaultValue(false), Category("Appearance")]
        public bool Expanded { get => expanded; set { if (expanded == value) return; expanded = value; UpdateExpansion(); ExpansionChanged?.Invoke(this, EventArgs.Empty); } }
        public ChatServices Services { get => services; set { services = value ?? new ChatServices(); UpdateCaption(); } }
        public SharedChatStyle Style { get => style; set { style = value ?? new SharedChatStyle(); BindStyle(); } }
        public event EventHandler ExpansionChanged;
        public void BindStyle()
        { BackColor = body.BackColor = toggle.BackColor = style.Background; ForeColor = body.ForeColor = toggle.ForeColor = style.Foreground; Font = style.BodyFont; RightToLeft = style.RightToLeft ? RightToLeft.Yes : RightToLeft.No; }
        private void UpdateCaption() { string caption = (expanded ? "▾ " : "▸ ") + services.Text(title); if (toggle.Text != caption) toggle.Text = caption; }
        private void UpdateExpansion() { UpdateCaption(); if (body.Visible != expanded) body.Visible = expanded; }
    }
}
