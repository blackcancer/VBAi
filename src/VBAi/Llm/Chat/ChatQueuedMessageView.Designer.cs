namespace VBAi
{

    /// <summary>Displays one queued chat message and provides actions to send it, edit it, or remove it.</summary>
    public sealed partial class ChatQueuedMessageView
    {

        /// <summary>Container that owns the disposable components created by the WinForms Designer.</summary>
        private System.ComponentModel.IContainer components;

        /// <summary>Flow layout panel that contains this transcript view&apos;s child controls.</summary>
        private System.Windows.Forms.TableLayoutPanel layout;

        /// <summary>Ellipsized preview of the queued message text.</summary>
        private System.Windows.Forms.Label message;

        /// <summary>Button that interrupts the current response and sends this queued message next.</summary>
        private ChatActionButton sendNow;

        /// <summary>Button that returns this queued message to the composer for editing.</summary>
        private ChatActionButton edit;

        /// <summary>Button that removes this message from the pending queue.</summary>
        private ChatActionButton delete;

        /// <summary>ToolTip component used to show full text for transcript controls.</summary>
        private System.Windows.Forms.ToolTip toolTips;

        /// <summary>Releases Designer-owned components.</summary>
        /// <param name="disposing">Whether managed resources should be released.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing) components?.Dispose();
            base.Dispose(disposing);
        }

        /// <summary>Creates and configures the chat queued message view controls serialized by the WinForms Designer.</summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.toolTips = new System.Windows.Forms.ToolTip(this.components);
            this.layout = new System.Windows.Forms.TableLayoutPanel();
            this.message = new System.Windows.Forms.Label();
            this.sendNow = new ChatActionButton();
            this.edit = new ChatActionButton();
            this.delete = new ChatActionButton();
            this.SuspendLayout(); this.layout.SuspendLayout();
            this.layout.ColumnCount = 4;
            this.layout.RowCount = 2;
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 26F));
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 34F));
            this.layout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.layout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Controls.Add(this.message, 0, 0);
            this.layout.SetColumnSpan(this.message, 4);
            this.layout.Controls.Add(this.sendNow, 1, 1);
            this.layout.Controls.Add(this.edit, 2, 1);
            this.layout.Controls.Add(this.delete, 3, 1);
            this.layout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layout.Name = "layout";
            this.message.AutoEllipsis = true;
            this.message.Dock = System.Windows.Forms.DockStyle.Fill;
            this.message.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.message.Name = "message";
            this.message.Text = "Queued message";
            this.sendNow.AutoSize = true;
            this.sendNow.Name = "sendNow";
            this.sendNow.Text = "Send now";
            this.edit.AutoSize = true;
            this.edit.Name = "edit";
            this.edit.Text = "Edit";
            this.delete.AutoSize = true;
            this.delete.Name = "delete";
            this.delete.Text = "Delete";
            this.toolTips.SetToolTip(this.sendNow, "Stop the current response and send this message next.");
            this.toolTips.SetToolTip(this.edit, "Move this queued message back to the composer.");
            this.toolTips.SetToolTip(this.delete, "Remove this message from the queue.");
            this.Controls.Add(this.layout);
            this.Name = "ChatQueuedMessageView";
            this.Size = new System.Drawing.Size(540, 60);
            this.Margin = new System.Windows.Forms.Padding(0, 2, 0, 2);
            this.layout.ResumeLayout(false); this.layout.PerformLayout();
            this.sendNow.Symbol = VBAi.UiSymbol.Play;
            this.sendNow.IconOnly = true;
            this.sendNow.AutoSize = false;
            this.sendNow.MinimumSize = System.Drawing.Size.Empty;
            this.sendNow.Size = new System.Drawing.Size(32, 30);
            this.edit.Symbol = VBAi.UiSymbol.Edit;
            this.edit.IconOnly = true;
            this.edit.AutoSize = false;
            this.edit.MinimumSize = System.Drawing.Size.Empty;
            this.edit.Size = new System.Drawing.Size(32, 30);
            this.delete.Symbol = VBAi.UiSymbol.Delete;
            this.delete.IconOnly = true;
            this.delete.AutoSize = false;
            this.delete.MinimumSize = System.Drawing.Size.Empty;
            this.delete.Size = new System.Drawing.Size(32, 30);
            this.ResumeLayout(false);
        }
    }
}
