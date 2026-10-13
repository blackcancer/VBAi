// SPDX-License-Identifier: MPL-2.0
using System;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;
using Blackcancer.ChatUI.Presentation;

namespace Blackcancer.ChatUI.Views
{
    /// <summary>Fixed internal card template used by the shared queue view.</summary>
    public partial class ChatPendingMessageCardView : UserControl
    {
        private Action editAction, removeAction, sendNextAction;
        public ChatPendingMessageCardView()
        {
            InitializeComponent();
            edit.Click += (sender, args) => { if (edit.Enabled) editAction?.Invoke(); };
            remove.Click += (sender, args) => { if (remove.Enabled) removeAction?.Invoke(); };
            sendNext.Click += (sender, args) => { if (sendNext.Enabled) sendNextAction?.Invoke(); };
        }
        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public ChatPendingMessageEntry Entry { get; private set; }
        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public TextBox PreviewBox => preview;

        internal void ConfigureRequests(Action onEdit, Action onRemove, Action onSendNext)
        { editAction = onEdit; removeAction = onRemove; sendNextAction = onSendNext; }
        internal void BindEntry(ChatPendingMessageEntry value, ChatServices services, SharedChatStyle style, CultureInfo culture)
        {
            Entry = value;
            if (preview.Text != value.Preview)
            {
                int start = preview.SelectionStart, length = preview.SelectionLength;
                preview.Text = value.Preview;
                start = Math.Min(start, preview.TextLength); preview.Select(start, Math.Min(length, preview.TextLength - start));
            }
            if (entryStatus.Text != value.Status) entryStatus.Text = value.Status;
            string position = value.Position.ToString(culture);
            if (positionLabel.Text != position) positionLabel.Text = position;
            var resources = new ComponentResourceManager(typeof(ChatPendingMessageCardView));
            foreach (Control control in new Control[] { edit, remove, sendNext })
                control.Text = services.Text(resources.GetString(control.Name + ".Text", culture) ?? control.Text);
            toolTips.SetToolTip(sendNext, services.Text(resources.GetString("sendNextTip", culture)
                ?? "Request this message next, after the current response. Sending remains controlled by the host."));
            ApplyPalette(this, style); ResetFont(); RightToLeft = style.RightToLeft ? RightToLeft.Yes : RightToLeft.No;
        }
        internal void SetActions(bool canEdit, bool canRemove, bool canSendNext)
        { edit.Enabled = canEdit; remove.Enabled = canRemove; sendNext.Enabled = canSendNext; }
        internal void SetCardWidth(int width)
        {
            if (Width == width && MaximumSize.Width == width) return;
            MinimumSize = new Size(width, 0); MaximumSize = new Size(width, 0); Width = width;
            int contentWidth = Math.Max(1, width - layout.Padding.Horizontal - 6);
            entryStatus.MaximumSize = new Size(contentWidth, 0);
            toolbar.MaximumSize = new Size(contentWidth, 0);
        }
        private static void ApplyPalette(Control control, SharedChatStyle style)
        {
            control.BackColor = style.Background; control.ForeColor = style.Foreground;
            foreach (Control child in control.Controls) ApplyPalette(child, style);
        }
    }
}
