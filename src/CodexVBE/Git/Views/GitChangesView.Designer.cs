namespace CodexVBE
{
    public sealed partial class GitChangesView
    {
        internal System.Windows.Forms.Label messageLabel;
        internal System.Windows.Forms.TextBox commitMessage;
        internal System.Windows.Forms.Button commit;
        internal System.Windows.Forms.SplitContainer changeSplit;
        internal CodexVBE.CodeDiffView diff;
        internal System.Windows.Forms.CheckedListBox changes;
        internal System.Windows.Forms.FlowLayoutPanel reviewActions;
        internal System.Windows.Forms.Button openModule;
        internal System.Windows.Forms.Button restoreModule;
        internal System.Windows.Forms.TableLayoutPanel editorLayout;
        private System.ComponentModel.IContainer components;
        private System.Windows.Forms.ToolTip toolTips;
        protected override void Dispose(bool disposing)
        {
            if (disposing) components?.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.toolTips = new System.Windows.Forms.ToolTip(this.components);
            this.editorLayout = new System.Windows.Forms.TableLayoutPanel();
            this.messageLabel = new System.Windows.Forms.Label();
            this.commitMessage = new System.Windows.Forms.TextBox();
            this.commit = new CodexVBE.ThemedButton();
            this.changeSplit = new System.Windows.Forms.SplitContainer();
            this.diff = new CodexVBE.CodeDiffView();
            this.changes = new System.Windows.Forms.CheckedListBox();
            this.reviewActions = new System.Windows.Forms.FlowLayoutPanel();
            this.openModule = new CodexVBE.ThemedButton();
            this.restoreModule = new CodexVBE.ThemedButton();
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
            this.commit.Margin = new System.Windows.Forms.Padding(3, 3, 15, 3);
            this.editorLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.editorLayout.ColumnCount = 1;
            this.editorLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.editorLayout.RowCount = 5;
            this.editorLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.editorLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 54F));
            this.editorLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.editorLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.editorLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.editorLayout.Controls.Add(this.messageLabel, 0, 0);
            this.editorLayout.Controls.Add(this.commitMessage, 0, 1);
            this.editorLayout.Controls.Add(this.commit, 0, 2);
            this.editorLayout.Controls.Add(this.changeSplit, 0, 3);
            this.reviewActions.Dock = System.Windows.Forms.DockStyle.Fill;
            this.editorLayout.Controls.Add(this.reviewActions, 0, 4);
            this.Controls.Add(this.editorLayout);
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Name = "GitChangesView";
            this.Size = new System.Drawing.Size(860, 500);
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
