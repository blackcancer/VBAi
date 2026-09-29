namespace VBAi
{
    /// <summary>Designer-generated controls for reviewing and committing project changes.</summary>
    public sealed partial class GitChangesView
    {
        /// <summary>Caption for the commit message input.</summary>
        internal System.Windows.Forms.Label messageLabel;
        /// <summary>Commit message input.</summary>
        internal VBAi.UiTextBox commitMessage;
        /// <summary>Commits selected changes.</summary>
        internal VBAi.ThemedButton commit;
        /// <summary>Split layout for changed files and their code diff.</summary>
        internal System.Windows.Forms.SplitContainer changeSplit;
        /// <summary>Side-by-side or unified diff for the selected file.</summary>
        internal VBAi.CodeDiffView diff;
        /// <summary>Changed files available for staging and review.</summary>
        internal VBAi.UiCheckedListBox changes;
        /// <summary>Actions for opening and restoring the selected module.</summary>
        internal System.Windows.Forms.FlowLayoutPanel reviewActions;
        /// <summary>Opens the selected VBA module.</summary>
        internal VBAi.ThemedButton openModule;
        /// <summary>Restores the selected module content.</summary>
        internal VBAi.ThemedButton restoreModule;
        /// <summary>Layout for the code-diff editor.</summary>
        internal System.Windows.Forms.TableLayoutPanel editorLayout;
        /// <summary>Container that owns Designer components.</summary>
        private System.ComponentModel.IContainer components;
        /// <summary>Chat-style commit composer, editable in the Designer.</summary>
        private VBAi.ChatComposerPanel commitComposer;
        /// <summary>Tooltips associated with change actions.</summary>
        private System.Windows.Forms.ToolTip toolTips;
        /// <summary>Libère les composants de la vue avant son contrôle natif.</summary>
        /// <param name="disposing">Indique si les ressources managées doivent être libérées.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing) components?.Dispose();
            base.Dispose(disposing);
        }

        /// <summary>Creates and arranges change review and commit controls.</summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.toolTips = new System.Windows.Forms.ToolTip(this.components);
            this.editorLayout = new System.Windows.Forms.TableLayoutPanel();
            this.commitComposer = new VBAi.ChatComposerPanel();
            this.commitComposer.SuspendLayout();
            this.messageLabel = new System.Windows.Forms.Label();
            this.commitMessage = new VBAi.UiTextBox();
            this.commit = new VBAi.ThemedButton();
            this.changeSplit = new System.Windows.Forms.SplitContainer();
            this.diff = new VBAi.CodeDiffView();
            this.changes = new VBAi.UiCheckedListBox();
            this.reviewActions = new System.Windows.Forms.FlowLayoutPanel();
            this.openModule = new VBAi.ThemedButton();
            this.restoreModule = new VBAi.ThemedButton();
            ((System.ComponentModel.ISupportInitialize)(this.changeSplit)).BeginInit();
            this.diff.SuspendLayout();
            this.reviewActions.SuspendLayout();
            this.editorLayout.SuspendLayout();
            this.changeSplit.Panel1.SuspendLayout();
            this.changeSplit.Panel2.SuspendLayout();
            this.changeSplit.SuspendLayout();
            this.SuspendLayout();
            this.messageLabel.Text = "Commit message";
            this.messageLabel.AutoSize = true;
            this.messageLabel.Name = "messageLabel";
            this.commitMessage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.commitMessage.Multiline = true;
            this.commitMessage.Name = "commitMessage";
            this.commitMessage.TabIndex = 3;
            this.commit.Text = "Commit selected";
            this.commit.AutoSize = true;
            this.commit.Enabled = false;
            this.commit.Name = "commit";
            this.toolTips.SetToolTip(this.commit, "Commit checked modules and their form resources. Push publishes them afterwards.");
            this.changes.Dock = System.Windows.Forms.DockStyle.Fill;
            this.changes.HorizontalScrollbar = true;
            this.changes.IntegralHeight = false;
            this.changes.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.changes.Name = "changes";
            this.changeSplit.Dock = System.Windows.Forms.DockStyle.Fill;
            this.changeSplit.Size = new System.Drawing.Size(860, 260);
            this.changeSplit.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this.changeSplit.SplitterDistance = 200;
            this.changeSplit.Name = "changeSplit";
            this.changeSplit.Panel1.Controls.Add(this.changes);
            this.changeSplit.Panel2.Controls.Add(this.diff);
            this.diff.Dock = System.Windows.Forms.DockStyle.Fill;
            this.diff.Name = "diff";
            this.reviewActions.Name = "reviewActions";
            this.openModule.Name = "openModule";
            this.restoreModule.Name = "restoreModule";
            this.reviewActions.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.reviewActions.AutoSize = true;
            this.openModule.Text = "Open in the VBE";
            this.openModule.AutoSize = true;
            this.restoreModule.Text = "Restore selected module";
            this.restoreModule.AutoSize = true;
            this.reviewActions.Controls.Add(this.openModule);
            this.reviewActions.Controls.Add(this.restoreModule);
            this.toolTips.SetToolTip(this.restoreModule, "Restore only the selected module from the reviewed revision, with a checkpoint first.");
            this.changes.CheckOnClick = true;
            this.commit.Margin = new System.Windows.Forms.Padding(3);
            this.editorLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.editorLayout.ColumnCount = 1;
            this.editorLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.editorLayout.RowCount = 5;
            this.editorLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.editorLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 0F));
            this.editorLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.editorLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.editorLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.editorLayout.Controls.Add(this.commitComposer, 0, 0);
            this.commitComposer.Name = "commitComposer";
            this.commitComposer.Dock = System.Windows.Forms.DockStyle.Fill;
            this.commitComposer.AutoSize = true;
            this.commitComposer.Padding = new System.Windows.Forms.Padding(10, 6, 10, 6);
            this.commitComposer.ColumnCount = 2;
            this.commitComposer.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.commitComposer.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 38F));
            this.commitComposer.RowCount = 2;
            this.commitComposer.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.commitComposer.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 54F));
            this.commitComposer.Controls.Add(this.messageLabel, 0, 0);
            this.commitComposer.Controls.Add(this.commitMessage, 0, 1);
            this.commitComposer.Controls.Add(this.commit, 1, 1);
            this.commitMessage.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.commit.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right;


            this.editorLayout.Controls.Add(this.changeSplit, 0, 3);
            this.reviewActions.Dock = System.Windows.Forms.DockStyle.Fill;
            this.editorLayout.Controls.Add(this.reviewActions, 0, 4);
            this.Controls.Add(this.editorLayout);
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Name = "GitChangesView";
            this.Size = new System.Drawing.Size(860, 500);
            this.messageLabel.Location = new System.Drawing.Point(3, 0);
            this.messageLabel.Size = new System.Drawing.Size(99, 21);
            this.messageLabel.TabIndex = 0;
            this.commitMessage.Location = new System.Drawing.Point(3, 24);
            this.commitMessage.Size = new System.Drawing.Size(854, 48);
            this.commit.Location = new System.Drawing.Point(3, 78);
            this.commit.Size = new System.Drawing.Size(106, 28);
            this.commit.TabIndex = 4;
            this.changeSplit.Location = new System.Drawing.Point(3, 112);
            this.changeSplit.TabIndex = 5;
            this.diff.Location = new System.Drawing.Point(0, 0);
            this.diff.Size = new System.Drawing.Size(650, 345);
            this.diff.TabIndex = 0;
            this.changes.Location = new System.Drawing.Point(0, 0);
            this.changes.Size = new System.Drawing.Size(200, 345);
            this.changes.TabIndex = 0;
            this.reviewActions.Location = new System.Drawing.Point(3, 463);
            this.reviewActions.Size = new System.Drawing.Size(854, 34);
            this.reviewActions.TabIndex = 6;
            this.openModule.Location = new System.Drawing.Point(3, 3);
            this.openModule.Size = new System.Drawing.Size(103, 28);
            this.openModule.TabIndex = 0;
            this.restoreModule.Location = new System.Drawing.Point(112, 3);
            this.restoreModule.Size = new System.Drawing.Size(149, 28);
            this.restoreModule.TabIndex = 1;
            this.editorLayout.Name = "editorLayout";
            this.editorLayout.Location = new System.Drawing.Point(0, 0);
            this.editorLayout.Size = new System.Drawing.Size(860, 500);
            this.editorLayout.TabIndex = 0;
            this.changeSplit.Panel1.ResumeLayout(false);
            this.changeSplit.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.changeSplit)).EndInit();
            this.diff.ResumeLayout(false);
            this.diff.PerformLayout();
            this.reviewActions.ResumeLayout(false);
            this.reviewActions.PerformLayout();
            this.editorLayout.ResumeLayout(false);
            this.editorLayout.PerformLayout();
            this.changeSplit.ResumeLayout(false);
            this.changeSplit.PerformLayout();
            this.commit.Symbol = VBAi.UiSymbol.Check;
            this.commit.IconOnly = true;
            this.commit.AutoSize = false;
            this.commit.MinimumSize = System.Drawing.Size.Empty;
            this.commit.Size = new System.Drawing.Size(32, 30);
            this.commit.Primary = true;
            this.openModule.Symbol = VBAi.UiSymbol.Code;
            this.openModule.IconOnly = true;
            this.openModule.AutoSize = false;
            this.openModule.MinimumSize = System.Drawing.Size.Empty;
            this.openModule.Size = new System.Drawing.Size(32, 30);
            this.restoreModule.Symbol = VBAi.UiSymbol.Undo;
            this.restoreModule.IconOnly = true;
            this.restoreModule.AutoSize = false;
            this.restoreModule.MinimumSize = System.Drawing.Size.Empty;
            this.restoreModule.Size = new System.Drawing.Size(32, 30);
            this.commitComposer.ResumeLayout(false);
            this.commitComposer.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
