using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
namespace CodexVBE
{
    /// <summary>Represents a chat message waiting to be sent, with its captured references, attachments, and memory.</summary>
internal sealed class QueuedChatMessage
    {
        /// <summary>Gets or sets the id.</summary>
/// <value>The current value represented by this member.</value>
public string Id { get; set; } = Guid.NewGuid().ToString("N");
        /// <summary>Gets or sets the text.</summary>
/// <value>The current value represented by this member.</value>
public string Text { get; set; }
        /// <summary>Gets or sets the references.</summary>
/// <value>The current value represented by this member.</value>
public VbeChatReference[] References { get; set; }
        /// <summary>Gets or sets the attachments.</summary>
/// <value>The current value represented by this member.</value>
public ChatAttachment[] Attachments { get; set; }
        /// <summary>Gets or sets the memory.</summary>
/// <value>The current value represented by this member.</value>
public string Memory { get; set; }
    }
    /// <summary>Provides the chat window implementation.</summary>
internal sealed partial class ChatWindow
    {
        /// <summary>Identifies the queued message selected for dispatch immediately after the active response stops.</summary>
private string immediateMessageId;
        /// <summary>Gets the current session&apos;s messages waiting for dispatch.</summary>
/// <value>The current session&apos;s queued messages, or null when there is no active session.</value>
private List<QueuedChatMessage> PendingMessages
        {
            get
            {
                if (currentSession == null) return null;
                return currentSession.PendingMessages ?? (currentSession.PendingMessages = new List<QueuedChatMessage>());
            }
        }
        /// <summary>Moves the current composer text, references, attachments, and captured memory into the pending queue.</summary>
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
        /// <summary>Rebuilds the pending message rows and hooks up their send, edit, and delete actions.</summary>
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
        /// <summary>Removes a queued message from the current session and persists the updated queue.</summary>
/// <param name="item">Queued message to remove.</param>
private void DeletePendingMessage(QueuedChatMessage item)
        {
            if (PendingMessages?.Remove(item) != true) return;
            if (immediateMessageId == item.Id) immediateMessageId = null;
            RefreshPendingMessages(); SaveCurrentSession();
        }
        /// <summary>Moves a queued message back into the composer when the current draft is empty.</summary>
/// <param name="item">Queued message to move back into the composer.</param>
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
        /// <summary>Dispatches a selected queued message immediately, stopping the active response when needed.</summary>
/// <param name="item">Queued message to dispatch ahead of other pending messages.</param>
/// <returns>The result produced by this operation.</returns>
private async Task SendPendingNowAsync(QueuedChatMessage item)
        {
            if (PendingMessages?.Contains(item) != true) return;
            if (!busy) { await SendRequestAsync(item); return; }
            if (stopRequested) return;
            immediateMessageId = item.Id;
            await StopTurnAsync();
            if (busy && !stopRequested) immediateMessageId = null;
        }
        /// <summary>Sends the next queued message after a response completes, unless dispatch is paused or stopped.</summary>
/// <param name="completed">Indicates whether completed is enabled.</param>
/// <returns>The result produced by this operation.</returns>
private async Task DispatchPendingAsync(bool completed)
        {
            if (IsDisposed || busy || PendingMessages == null) return;
            var item = immediateMessageId == null ? null : PendingMessages.FirstOrDefault(m => m.Id == immediateMessageId);
            immediateMessageId = null;
            if (item == null && completed && !stopRequested && currentSession.BudgetPaused != true) item = PendingMessages.FirstOrDefault();
            if (item != null) await SendRequestAsync(item);
        }
        /// <summary>Keeps the selected queued message memory attached while its text is being edited in the composer.</summary>
private string queuedDraftMemory;
    }
}
