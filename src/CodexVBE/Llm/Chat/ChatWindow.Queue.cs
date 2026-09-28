using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
namespace CodexVBE
{
    internal sealed class QueuedChatMessage
    {
        public string Id { get; set; } = Guid.NewGuid().ToString("N");
        public string Text { get; set; }
        public VbeChatReference[] References { get; set; }
        public ChatAttachment[] Attachments { get; set; }
        public string Memory { get; set; }
    }
    internal sealed partial class ChatWindow
    {
        private string immediateMessageId;
        private List<QueuedChatMessage> PendingMessages
        {
            get
            {
                if (currentSession == null) return null;
                return currentSession.PendingMessages ?? (currentSession.PendingMessages = new List<QueuedChatMessage>());
            }
        }
        private void QueueComposerMessage()
        {
            if (currentSession == null || string.IsNullOrWhiteSpace(prompt.Text)) return;
            try { EnsureCurrentScope(); }
            catch (Exception error) { SetStatus(error.Message); return; }
            PendingMessages.Add(new QueuedChatMessage {
                Text = prompt.Text.Trim(), References = CurrentReferences(prompt.Text), Attachments = draftAttachments.ToArray(),
                Memory = attachMemory.Checked ? queuedDraftMemory ?? projectMemory : null
            });
            queuedDraftMemory = null; prompt.Clear(); selectedReferences.Clear(); draftAttachments.Clear(); attachMemory.Checked = false;
            HideReferences(); RefreshContextChips(); RefreshPendingMessages(); SaveCurrentSession();
        }
        private void RefreshPendingMessages()
        {
            if (pendingMessagesPanel == null) return;
            pendingMessagesPanel.SuspendLayout();
            try
            {
                while (pendingMessagesPanel.Controls.Count > 0) pendingMessagesPanel.Controls[0].Dispose();
                foreach (var item in PendingMessages ?? new List<QueuedChatMessage>())
                {
                    var row = new ChatQueuedMessageView();
                    row.ShowMessage(item.Text);
                    row.SendNowRequested += async (s, e) => await SendPendingNowAsync(item);
                    row.EditRequested += (s, e) => EditPendingMessage(item);
                    row.DeleteRequested += (s, e) => DeletePendingMessage(item);
                    UiTheme.Apply(row);
                    row.Width = Math.Max(200, pendingMessagesPanel.ClientSize.Width - 24);
                    pendingMessagesPanel.Controls.Add(row);
                }
                pendingMessagesPanel.Visible = pendingMessagesPanel.Controls.Count > 0;
            }
            finally { pendingMessagesPanel.ResumeLayout(true); }
        }
        private void DeletePendingMessage(QueuedChatMessage item)
        {
            if (PendingMessages?.Remove(item) != true) return;
            if (immediateMessageId == item.Id) immediateMessageId = null;
            RefreshPendingMessages(); SaveCurrentSession();
        }
        private void EditPendingMessage(QueuedChatMessage item)
        {
            if (PendingMessages?.Contains(item) != true) return;
            // Keep an existing draft intact rather than silently replacing it.
            if (!string.IsNullOrWhiteSpace(prompt.Text) || draftAttachments.Count > 0 || selectedReferences.Count > 0 || attachMemory.Checked)
            { SetStatus(UiText.Get("Send or clear the current draft before editing a queued message.")); return; }
            PendingMessages.Remove(item);
            if (immediateMessageId == item.Id) immediateMessageId = null;
            selectedReferences.AddRange(item.References ?? new VbeChatReference[0]);
            draftAttachments.AddRange(item.Attachments ?? new ChatAttachment[0]);
            // Store the captured memory with the draft, without changing project memory.
            queuedDraftMemory = item.Memory;
            attachMemory.Checked = !string.IsNullOrWhiteSpace(item.Memory);
            prompt.Text = item.Text; prompt.CaretIndex = prompt.Text.Length; prompt.Focus();
            RefreshContextChips(); RefreshPendingMessages(); SaveCurrentSession();
        }
        private async Task SendPendingNowAsync(QueuedChatMessage item)
        {
            if (PendingMessages?.Contains(item) != true) return;
            if (!busy) { await SendRequestAsync(item); return; }
            if (stopRequested) return;
            immediateMessageId = item.Id;
            await StopTurnAsync();
            if (busy && !stopRequested) immediateMessageId = null;
        }
        private async Task DispatchPendingAsync(bool completed)
        {
            if (IsDisposed || busy || PendingMessages == null) return;
            var item = immediateMessageId == null ? null : PendingMessages.FirstOrDefault(m => m.Id == immediateMessageId);
            immediateMessageId = null;
            if (item == null && completed && !stopRequested && currentSession.BudgetPaused != true) item = PendingMessages.FirstOrDefault();
            if (item != null) await SendRequestAsync(item);
        }
        private string queuedDraftMemory;
    }
}
