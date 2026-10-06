namespace VBAi
{

    /// <summary>Displays a transcript message with its content, metadata, and available actions.</summary>
    public sealed partial class ChatMessageView
    {

        /// <summary>Container that owns the disposable components created by the WinForms Designer.</summary>
        private System.ComponentModel.IContainer components;

        /// <summary>ToolTip component used to show full text for transcript controls.</summary>
        private System.Windows.Forms.ToolTip toolTips;

        /// <summary>Flow layout panel that contains this transcript view&apos;s child controls.</summary>
        private ChatComposerPanel layout;

        /// <summary>Arranges the message author and its available actions.</summary>
        private System.Windows.Forms.TableLayoutPanel header;

        /// <summary>Displays the role or name of the message author.</summary>
        internal System.Windows.Forms.Label speaker;

        /// <summary>Provides the action that copies the message content.</summary>
        internal ChatActionButton copy;

        /// <summary>Provides the action that starts a conversation from this message.</summary>
        internal ChatActionButton fork;

        /// <summary>Contains the actions shown alongside the message author.</summary>
        internal System.Windows.Forms.FlowLayoutPanel headingActions;

        /// <summary>Displays the message text and its interactive references.</summary>
        internal ChatTextContentView message;

        /// <summary>Displays the memory snapshot attached to this message.</summary>
        internal ChatDisclosureView memory;

        /// <summary>Contains the files attached to this message.</summary>
        internal System.Windows.Forms.FlowLayoutPanel attachments;

        /// <summary>Contains the VBA references associated with this message.</summary>
        internal System.Windows.Forms.FlowLayoutPanel references;

        /// <summary>Contains the code change targets produced by this message.</summary>
        internal System.Windows.Forms.FlowLayoutPanel targets;

        /// <summary>Button that requests rollback of all changes in this conversation turn.</summary>
        internal ChatActionButton undoTurn;

        /// <summary>Button that prepares a follow-up fix request from the message and its code context.</summary>
        internal ChatActionButton fix;

        /// <summary>Releases the Designer components.</summary>
        /// <param name="disposing">Whether to release managed resources.</param>
        protected override void Dispose(bool disposing) { if (disposing) components?.Dispose(); base.Dispose(disposing); }

        /// <summary>Creates and configures the chat message view controls serialized by the WinForms Designer.</summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.toolTips = new System.Windows.Forms.ToolTip(this.components);
            this.copy = new ChatActionButton(); this.copy.Name = "copy"; this.copy.Text = "Copy"; this.copy.AutoSize = true;
            this.fork = new ChatActionButton(); this.fork.Name = "fork"; this.fork.Text = "Branch conversation"; this.fork.AutoSize = true;
            this.layout = new ChatComposerPanel();
            this.layout.ShowBorder = false;
            this.header = new System.Windows.Forms.TableLayoutPanel();
            this.speaker = new System.Windows.Forms.Label();
            this.speaker.Name = "speaker";
            this.speaker.AutoSize = true;
            this.speaker.Dock = System.Windows.Forms.DockStyle.Fill;
            this.speaker.Text = "Assistant";
            this.headingActions = new System.Windows.Forms.FlowLayoutPanel();
            this.headingActions.Name = "headingActions";
            this.headingActions.Controls.Add(this.copy); this.headingActions.Controls.Add(this.fork);
            this.headingActions.AutoSize = true;
            this.headingActions.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.headingActions.Dock = System.Windows.Forms.DockStyle.Fill;
            this.headingActions.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft;
            this.headingActions.WrapContents = false;
            this.message = new ChatTextContentView();
            this.message.Name = "message";
            this.message.Dock = System.Windows.Forms.DockStyle.Fill;
            this.message.Margin = new System.Windows.Forms.Padding(0);
            this.memory = new ChatDisclosureView();
            this.memory.Name = "memory";
            this.memory.Dock = System.Windows.Forms.DockStyle.Fill;
            this.memory.Title = "Attached document memory";
            this.attachments = new System.Windows.Forms.FlowLayoutPanel();
            this.attachments.Name = "attachments";
            this.attachments.AutoSize = true;
            this.attachments.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.attachments.Dock = System.Windows.Forms.DockStyle.Fill;
            this.attachments.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.attachments.WrapContents = false;
            this.references = new System.Windows.Forms.FlowLayoutPanel();
            this.references.Name = "references";
            this.references.AutoSize = true;
            this.references.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.references.Dock = System.Windows.Forms.DockStyle.Fill;
            this.references.FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight;
            this.references.WrapContents = true;
            this.targets = new System.Windows.Forms.FlowLayoutPanel();
            this.targets.Name = "targets";
            this.targets.AutoSize = true;
            this.targets.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.targets.Dock = System.Windows.Forms.DockStyle.Fill;
            this.targets.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.targets.WrapContents = false;
            this.undoTurn = new ChatActionButton();
            this.undoTurn.Name = "undoTurn";
            this.undoTurn.AutoSize = true;
            this.undoTurn.Text = "Undo entire turn";
            this.fix = new ChatActionButton();
            this.fix.Name = "fix";
            this.fix.AutoSize = true;
            this.fix.Text = "Prepare a fix";
            this.header.Name = "header";
            this.header.AutoSize = true;
            this.header.Dock = System.Windows.Forms.DockStyle.Fill;
            this.header.ColumnCount = 2;
            this.header.RowCount = 1;
            this.header.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.header.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.AutoSize));
            this.header.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.header.Controls.Add(this.speaker, 0, 0);
            this.header.Controls.Add(this.headingActions, 1, 0);
            this.speaker.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.SuspendLayout(); this.layout.SuspendLayout();
            this.layout.AutoSize = true;
            this.layout.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.layout.Dock = System.Windows.Forms.DockStyle.Top;
            this.layout.ColumnCount = 1;
            this.layout.RowCount = 8;
            this.layout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.layout.Controls.Add(this.header, 0, 0);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Controls.Add(this.message, 0, 1);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Controls.Add(this.memory, 0, 2);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Controls.Add(this.attachments, 0, 3);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Controls.Add(this.references, 0, 4);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Controls.Add(this.targets, 0, 5);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Controls.Add(this.undoTurn, 0, 6);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Controls.Add(this.fix, 0, 7);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Name = "layout";
            this.layout.Padding = new System.Windows.Forms.Padding(4);
            this.toolTips.SetToolTip(this.copy, "Copy the message text to the clipboard.");
            this.toolTips.SetToolTip(this.fork, "Create an independent conversation with the history up to this message.");
            this.toolTips.SetToolTip(this.undoTurn, "Undo changes from this turn after checking for conflicts.");
            this.toolTips.SetToolTip(this.fix, "Prepare a fix");
            this.Controls.Add(this.layout);
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Name = "ChatMessageView";
            this.Size = new System.Drawing.Size(500, 120);
            this.layout.ResumeLayout(false); this.layout.PerformLayout();
            this.copy.Symbol = VBAi.UiSymbol.Copy;
            this.copy.IconOnly = true;
            this.copy.AutoSize = false;
            this.copy.MinimumSize = System.Drawing.Size.Empty;
            this.copy.Size = new System.Drawing.Size(32, 30);
            this.fork.Symbol = VBAi.UiSymbol.Add;
            this.fork.IconOnly = true;
            this.fork.AutoSize = false;
            this.fork.Size = new System.Drawing.Size(32, 30);
            this.undoTurn.Symbol = VBAi.UiSymbol.Undo;
            this.undoTurn.IconOnly = true;
            this.undoTurn.AutoSize = false;
            this.undoTurn.Size = new System.Drawing.Size(32, 30);
            this.fix.Symbol = VBAi.UiSymbol.Check;
            this.fix.IconOnly = true;
            this.fix.AutoSize = false;
            this.fix.Size = new System.Drawing.Size(32, 30);
            this.ResumeLayout(false); this.PerformLayout();
        }
    }
}
