namespace VBAi
{

    /// <summary>Displays a proposed code change with its diff and available undo actions.</summary>
    public sealed partial class ChatChangeCardView
    {

        /// <summary>Container that owns the disposable components created by the WinForms Designer.</summary>
        private System.ComponentModel.IContainer components;

        /// <summary>ToolTip component used to show full text for transcript controls.</summary>
        private System.Windows.Forms.ToolTip toolTips;

        /// <summary>Flow layout panel that contains this transcript view&apos;s child controls.</summary>
        private ChatComposerPanel layout;

        /// <summary>Displays the VBA component containing the proposed change.</summary>
        internal ChatActionButton module;

        /// <summary>Displays the number of changed lines or blocks.</summary>
        internal System.Windows.Forms.Label count;

        /// <summary>Expandable section containing the proposed diff and its undo controls.</summary>
        internal ChatDisclosureView section;

        /// <summary>Displays the code diff for the proposed change.</summary>
        internal CodeDiffView diff;

        /// <summary>Provides the action that reverts the proposed code change.</summary>
        internal ChatActionButton undo;

        /// <summary>Contains the individual change blocks that can be undone.</summary>
        internal ChatActionButton blocks;

        /// <summary>Provides the action that undoes the complete assistant turn.</summary>
        internal ChatActionButton undoTurn;

        /// <summary>Context menu for actions on an individual code change block.</summary>
        internal System.Windows.Forms.ContextMenuStrip blockMenu;

        /// <summary>Contains the actions available for this change card.</summary>
        internal System.Windows.Forms.FlowLayoutPanel actions;

        /// <summary>Displays the current status of the proposed code change.</summary>
        internal System.Windows.Forms.Label state;

        /// <summary>Releases the Designer components.</summary>
        /// <param name="disposing">Whether to release managed resources.</param>
        protected override void Dispose(bool disposing) { if (disposing) components?.Dispose(); base.Dispose(disposing); }

        /// <summary>Creates and configures the chat change card view controls serialized by the WinForms Designer.</summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.toolTips = new System.Windows.Forms.ToolTip(this.components);
            this.blockMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.blockMenu.Name = "blockMenu";
            this.undo = new ChatActionButton(); this.undo.Name = "undo"; this.undo.Text = "Undo change"; this.undo.AutoSize = true;
            this.blocks = new ChatActionButton(); this.blocks.Name = "blocks"; this.blocks.Text = "Undo a block…"; this.blocks.AutoSize = true;
            this.undoTurn = new ChatActionButton(); this.undoTurn.Name = "undoTurn"; this.undoTurn.Text = "Undo turn"; this.undoTurn.AutoSize = true;
            this.layout = new ChatComposerPanel();
            this.module = new ChatActionButton();
            this.module.Name = "module";
            this.module.AutoSize = true;
            this.module.Text = "Open the modified module";
            this.count = new System.Windows.Forms.Label();
            this.count.Name = "count";
            this.count.AutoSize = true;
            this.count.Dock = System.Windows.Forms.DockStyle.Fill;
            this.count.Text = "Changes";
            this.diff = new CodeDiffView(); this.diff.Name = "diff"; this.diff.UnifiedDiff = true; this.diff.Height = 300; this.diff.MinimumSize = new System.Drawing.Size(0, 200); this.diff.Dock = System.Windows.Forms.DockStyle.None;
            this.section = new ChatDisclosureView();
            this.section.ContentPanel.Controls.Add(this.diff);
            this.section.Name = "section";
            this.section.Dock = System.Windows.Forms.DockStyle.Fill;
            this.section.Title = "Change diff";
            this.actions = new System.Windows.Forms.FlowLayoutPanel();
            this.actions.Name = "actions";
            this.actions.Controls.Add(this.undo); this.actions.Controls.Add(this.blocks); this.actions.Controls.Add(this.undoTurn);
            this.actions.AutoSize = true;
            this.actions.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.actions.Dock = System.Windows.Forms.DockStyle.Fill;
            this.actions.FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight;
            this.actions.WrapContents = true;
            this.state = new System.Windows.Forms.Label();
            this.state.Name = "state";
            this.state.AutoSize = true;
            this.state.Dock = System.Windows.Forms.DockStyle.Fill;
            this.state.Text = "Applied";
            this.SuspendLayout(); this.layout.SuspendLayout();
            this.layout.AutoSize = true;
            this.layout.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.layout.Dock = System.Windows.Forms.DockStyle.Top;
            this.layout.ColumnCount = 1;
            this.layout.RowCount = 5;
            this.layout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.layout.Controls.Add(this.module, 0, 0);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Controls.Add(this.count, 0, 1);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Controls.Add(this.section, 0, 2);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Controls.Add(this.actions, 0, 3);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Controls.Add(this.state, 0, 4);
            this.layout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.layout.Name = "layout";
            this.layout.Padding = new System.Windows.Forms.Padding(4);
            this.toolTips.SetToolTip(this.undoTurn, "Undo changes from this turn after checking for conflicts.");
            this.toolTips.SetToolTip(this.module, "Open the modified module in the VBE.");
            this.toolTips.SetToolTip(this.undo, "Restore the code before this change after checking for conflicts.");
            this.toolTips.SetToolTip(this.blocks, "Choose a block to undo while keeping other changes.");
            this.Controls.Add(this.layout);
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Name = "ChatChangeCardView";
            this.Size = new System.Drawing.Size(500, 120);
            this.layout.ResumeLayout(false); this.layout.PerformLayout();
            this.undo.Symbol = VBAi.UiSymbol.Undo;
            this.undo.IconOnly = true;
            this.undo.AutoSize = false;
            this.undo.MinimumSize = System.Drawing.Size.Empty;
            this.undo.Size = new System.Drawing.Size(32, 30);
            this.blocks.Symbol = VBAi.UiSymbol.Undo;
            this.blocks.IconOnly = true;
            this.blocks.AutoSize = false;
            this.blocks.Size = new System.Drawing.Size(32, 30);
            this.undoTurn.Symbol = VBAi.UiSymbol.Undo;
            this.undoTurn.IconOnly = true;
            this.undoTurn.AutoSize = false;
            this.undoTurn.Size = new System.Drawing.Size(32, 30);
            this.module.Symbol = VBAi.UiSymbol.Code;
            this.module.IconOnly = false;
            this.module.AutoSize = true;
            this.module.Size = new System.Drawing.Size(32, 30);
            this.ResumeLayout(false); this.PerformLayout();
        }
    }
}
