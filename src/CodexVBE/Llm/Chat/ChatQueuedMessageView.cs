using System;
using System.Windows.Forms;
namespace CodexVBE
{
    /// <summary>Designer-editable row for a pending chat message.</summary>
    public sealed partial class ChatQueuedMessageView : UserControl
    {
        /// <summary>Requests immediate dispatch after interrupting the active response.</summary>
        public event EventHandler SendNowRequested;
        /// <summary>Requests moving this message back to the composer.</summary>
        public event EventHandler EditRequested;
        /// <summary>Requests removal from the pending queue.</summary>
        public event EventHandler DeleteRequested;
        /// <summary>Creates the Designer controls and action handlers.</summary>
        public ChatQueuedMessageView()
        {
            InitializeComponent(); UiText.Apply(this, components);
            sendNow.Click += (s, e) => SendNowRequested?.Invoke(this, e);
            edit.Click += (s, e) => EditRequested?.Invoke(this, e);
            delete.Click += (s, e) => DeleteRequested?.Invoke(this, e);
        }
        /// <summary>Displays the queued message with a full-text tooltip.</summary>
        /// <param name="text">User message awaiting dispatch.</param>
        public void ShowMessage(string text) { message.Text = text; toolTips.SetToolTip(message, text); }
    }
}
