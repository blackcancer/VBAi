// SPDX-License-Identifier: MPL-2.0
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using Blackcancer.ChatUI.Presentation;

namespace Blackcancer.ChatUI.Views
{
    /// <summary>Queue presentation only. Requests never remove, edit, stop or send a message in this module.</summary>
    public partial class ChatPendingMessageListView : UserControl
    {
        private readonly Dictionary<string, ChatPendingMessageCardView> cards = new Dictionary<string, ChatPendingMessageCardView>(StringComparer.Ordinal);
        private IReadOnlyList<ChatPendingMessageEntry> messages = new ReadOnlyCollection<ChatPendingMessageEntry>(new ChatPendingMessageEntry[0]);
        private Action<ChatPendingMessageEntry> editAction, removeAction, sendNextAction;
        private Action<string> reportError = message => System.Diagnostics.Debug.WriteLine(message);
        private ChatServices services = new ChatServices();
        private SharedChatStyle style = new SharedChatStyle();
        private CultureInfo resourceCulture;
        private bool busy, dispatching;
        private Font ownedFont;

        public ChatPendingMessageListView()
        {
            InitializeComponent();
            messageFlow.SizeChanged += (sender, args) => ResizeCards();
        }

        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public IReadOnlyList<ChatPendingMessageEntry> Messages => messages;
        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public bool IsBusy => busy;
        [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public string Status => statusNote.Text;

        public void Configure(Action<ChatPendingMessageEntry> onEdit = null, Action<ChatPendingMessageEntry> onRemove = null,
            Action<ChatPendingMessageEntry> onSendNext = null, ChatServices services = null, SharedChatStyle style = null,
            Action<string> onError = null, CultureInfo culture = null)
        {
            editAction = onEdit; removeAction = onRemove; sendNextAction = onSendNext;
            this.services = services ?? new ChatServices(); this.style = style ?? new SharedChatStyle();
            reportError = onError ?? (message => System.Diagnostics.Debug.WriteLine(message));
            resourceCulture = culture ?? CultureInfo.CurrentUICulture;
            var nextFont = (Font)(this.style.BodyFont ?? SystemFonts.MessageBoxFont).Clone();
            if (ownedFont != null && ownedFont.Equals(nextFont)) nextFont.Dispose();
            else { Font previous = ownedFont; Font = ownedFont = nextFont; previous?.Dispose(); }
            ApplyPalette(this); RightToLeft = this.style.RightToLeft ? RightToLeft.Yes : RightToLeft.No;
            foreach (var card in cards.Values) card.BindEntry(card.Entry, this.services, this.style, resourceCulture);
            UpdateHeading(); UpdateActions(); ResizeCards();
        }

        /// <summary>Snapshot the supplied entries and retain existing cards by ordinal ID, including their text selection.</summary>
        public void UpdateMessages(IEnumerable<ChatPendingMessageEntry> values)
        {
            if (IsDisposed) return;
            var next = (values ?? throw new ArgumentNullException(nameof(values))).ToArray();
            if (next.Any(value => value == null) || next.Select(value => value.Id).Distinct(StringComparer.Ordinal).Count() != next.Length)
                throw new ArgumentException("Pending messages require unique non-null identifiers.", nameof(values));
            next = next.OrderBy(value => value.Position).ThenBy(value => value.Id, StringComparer.Ordinal).ToArray();
            var ids = new HashSet<string>(next.Select(value => value.Id), StringComparer.Ordinal);
            Point scroll = messageFlow.AutoScrollPosition;
            messageFlow.SuspendLayout();
            try
            {
                foreach (string id in cards.Keys.Where(id => !ids.Contains(id)).ToArray())
                { var old = cards[id]; messageFlow.Controls.Remove(old); cards.Remove(id); old.Dispose(); }
                for (int index = 0; index < next.Length; index++)
                {
                    ChatPendingMessageEntry value = next[index]; ChatPendingMessageCardView card;
                    if (!cards.TryGetValue(value.Id, out card))
                    {
                        card = new ChatPendingMessageCardView(); cards.Add(value.Id, card);
                        string id = value.Id;
                        card.ConfigureRequests(() => Request(id, editAction), () => Request(id, removeAction), () => Request(id, sendNextAction));
                        messageFlow.Controls.Add(card);
                    }
                    card.BindEntry(value, services, style, resourceCulture ?? CultureInfo.CurrentUICulture);
                    if (messageFlow.Controls.GetChildIndex(card) != index) messageFlow.Controls.SetChildIndex(card, index);
                }
                messages = new ReadOnlyCollection<ChatPendingMessageEntry>(next);
                ResizeCards(); UpdateHeading(); UpdateActions();
            }
            finally { messageFlow.ResumeLayout(true); }
            messageFlow.AutoScrollPosition = new Point(-scroll.X, -scroll.Y);
        }

        public void SetBusy(bool isBusy, string status = null)
        {
            if (IsDisposed) return;
            busy = isBusy; if (status != null) statusNote.Text = status; UpdateActions();
        }

        public void SetStatus(string status)
        { if (!IsDisposed) statusNote.Text = status ?? string.Empty; }

        private void Request(string id, Action<ChatPendingMessageEntry> action)
        {
            ChatPendingMessageCardView card;
            if (IsDisposed || busy || dispatching || action == null || !cards.TryGetValue(id, out card)) return;
            ChatPendingMessageEntry entry = card.Entry;
            busy = true; dispatching = true; UpdateActions();
            try { action(entry); }
            catch (Exception error) { reportError(error.Message); }
            finally { dispatching = false; if (!IsDisposed) UpdateActions(); }
        }
        private void UpdateActions()
        {
            bool available = !busy && !dispatching;
            foreach (var card in cards.Values) card.SetActions(available && editAction != null,
                available && removeAction != null, available && sendNextAction != null);
        }
        private void UpdateHeading()
        {
            var resources = new ComponentResourceManager(typeof(ChatPendingMessageListView));
            header.Text = string.Format(resourceCulture ?? CultureInfo.CurrentUICulture,
                services.Text(resources.GetString("headerFormat", resourceCulture ?? CultureInfo.CurrentUICulture) ?? "Queued messages ({0})"), messages.Count);
            emptyNote.Text = services.Text(resources.GetString("emptyNote.Text", resourceCulture ?? CultureInfo.CurrentUICulture) ?? "No queued messages.");
            emptyNote.Visible = messages.Count == 0;
        }
        private void ResizeCards()
        {
            foreach (var card in cards.Values)
                card.SetCardWidth(Math.Max(1, messageFlow.ClientSize.Width - messageFlow.Padding.Horizontal
                    - card.Margin.Horizontal - SystemInformation.VerticalScrollBarWidth));
        }
        private void ApplyPalette(Control control)
        {
            control.BackColor = style.Background; control.ForeColor = style.Foreground;
            foreach (Control child in control.Controls) ApplyPalette(child);
        }
        private void DisposeOwnedAppearance()
        { ownedFont?.Dispose(); ownedFont = null; }
    }
}
