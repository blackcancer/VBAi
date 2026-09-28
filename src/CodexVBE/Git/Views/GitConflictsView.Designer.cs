namespace CodexVBE
{
    public sealed partial class GitConflictsView
    {
        internal System.Windows.Forms.FlowLayoutPanel conflictActions;
        internal System.Windows.Forms.ListBox conflictList;
        internal System.Windows.Forms.TextBox resolutionText;
        internal System.Windows.Forms.Button mergeOurs;
        internal System.Windows.Forms.Button mergeTheirs;
        internal System.Windows.Forms.Button mergeText;
        internal System.Windows.Forms.Button mergeComplete;
        internal System.Windows.Forms.Button mergeAbort;
        internal System.Windows.Forms.DataGridView conflictDiff;
        internal System.Windows.Forms.DataGridViewTextBoxColumn conflictOurs;
        internal System.Windows.Forms.DataGridViewTextBoxColumn conflictTheirs;
        internal System.Windows.Forms.TextBox baseContent;
        internal System.Windows.Forms.TableLayoutPanel conflictLayout;
        internal System.Windows.Forms.Label ancestorLabel;
        internal System.Windows.Forms.Label resultLabel;
        private System.ComponentModel.IContainer components;
        private System.Windows.Forms.ToolTip toolTips;
        /// <summary>Libère les composants de la vue avant son contrôle natif.</summary>
        /// <param name="disposing">Indique si les ressources managées doivent être libérées.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing) components?.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.toolTips = new System.Windows.Forms.ToolTip(this.components);
            this.conflictActions = new System.Windows.Forms.FlowLayoutPanel();
            this.conflictList = new System.Windows.Forms.ListBox();
            this.resolutionText = new System.Windows.Forms.TextBox();
            this.mergeOurs = new CodexVBE.ThemedButton();
            this.mergeTheirs = new CodexVBE.ThemedButton();
            this.mergeText = new CodexVBE.ThemedButton();
            this.mergeComplete = new CodexVBE.ThemedButton();
            this.mergeAbort = new CodexVBE.ThemedButton();
            this.conflictDiff = new System.Windows.Forms.DataGridView();
            this.conflictOurs = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.conflictTheirs = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.baseContent = new System.Windows.Forms.TextBox();
            this.conflictLayout = new System.Windows.Forms.TableLayoutPanel();
            this.ancestorLabel = new System.Windows.Forms.Label();
            this.resultLabel = new System.Windows.Forms.Label();
            this.SuspendLayout();
            this.conflictActions.Name = "conflictActions";
            this.mergeOurs.Name = "mergeOurs";
            this.mergeTheirs.Name = "mergeTheirs";
            this.mergeText.Name = "mergeText";
            this.mergeComplete.Name = "mergeComplete";
            this.mergeAbort.Name = "mergeAbort";
            this.conflictOurs.Name = "conflictOurs";
            this.conflictTheirs.Name = "conflictTheirs";
            this.conflictList.Dock = System.Windows.Forms.DockStyle.Top;
            this.conflictList.Height = 55;
            this.conflictList.Name = "conflictList";
            this.conflictDiff.Dock = System.Windows.Forms.DockStyle.Top;
            this.conflictDiff.Height = 110;
            this.conflictDiff.Name = "conflictDiff";
            this.conflictDiff.ReadOnly = true;
            this.conflictDiff.AllowUserToAddRows = false;
            this.conflictDiff.AllowUserToDeleteRows = false;
            this.conflictDiff.RowHeadersVisible = false;
            this.conflictDiff.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.conflictDiff.BackgroundColor = System.Drawing.Color.White;
            this.conflictDiff.Font = new System.Drawing.Font("Consolas", 9F);
            this.conflictDiff.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] { this.conflictOurs, this.conflictTheirs });
            this.conflictOurs.HeaderText = "Current branch";
            this.conflictOurs.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.conflictTheirs.HeaderText = "Incoming branch";
            this.conflictTheirs.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.resolutionText.Dock = System.Windows.Forms.DockStyle.Fill;
            this.resolutionText.Multiline = true;
            this.resolutionText.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.resolutionText.WordWrap = false;
            this.resolutionText.Font = new System.Drawing.Font("Consolas", 9F);
            this.resolutionText.Name = "resolutionText";
            this.resolutionText.AccessibleName = "Full resolution content for the selected VBA file";
            this.conflictActions.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.conflictActions.AutoSize = true;
            this.conflictActions.Controls.Add(this.mergeOurs);
            this.conflictActions.Controls.Add(this.mergeTheirs);
            this.conflictActions.Controls.Add(this.mergeText);
            this.conflictActions.Controls.Add(this.mergeComplete);
            this.conflictActions.Controls.Add(this.mergeAbort);
            this.mergeOurs.Text = "Keep local";
            this.mergeOurs.AutoSize = true;
            this.mergeTheirs.Text = "Keep incoming";
            this.mergeTheirs.AutoSize = true;
            this.mergeText.Text = "Use this text";
            this.mergeText.AutoSize = true;
            this.mergeComplete.Text = "Complete merge";
            this.mergeComplete.AutoSize = true;
            this.mergeAbort.Text = "Abort merge";
            this.mergeAbort.AutoSize = true;
            this.toolTips.SetToolTip(this.mergeOurs, "Choose the entire current-branch file for the selected conflict.");
            this.toolTips.SetToolTip(this.mergeTheirs, "Choose the entire incoming-branch file for the selected conflict.");
            this.toolTips.SetToolTip(this.mergeText, "Replace the entire VBA text file with the text edited above.");
            this.toolTips.SetToolTip(this.mergeComplete, "Validate all files, create a merge commit, then import VBA with a backup.");
            this.toolTips.SetToolTip(this.mergeAbort, "Discard only the prepared merge. VBA and commits remain intact.");
            this.baseContent.Name = "baseContent";
            this.baseContent.Dock = System.Windows.Forms.DockStyle.Top;
            this.baseContent.Height = 100;
            this.baseContent.Multiline = true;
            this.baseContent.ReadOnly = true;
            this.baseContent.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.baseContent.Font = new System.Drawing.Font("Consolas", 9F);
            this.baseContent.AccessibleName = "Common ancestor";
            this.toolTips.SetToolTip(this.baseContent, "Common ancestor");
            this.conflictLayout.Name = "conflictLayout";
            this.conflictLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.conflictLayout.ColumnCount = 1;
            this.conflictLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.conflictLayout.RowCount = 7;
            this.conflictLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 48F));
            this.conflictLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.conflictLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.conflictLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 35F));
            this.conflictLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.conflictLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 40F));
            this.conflictLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            this.ancestorLabel.Name = "ancestorLabel";
            this.ancestorLabel.Text = "Common ancestor";
            this.ancestorLabel.AutoSize = true;
            this.resultLabel.Name = "resultLabel";
            this.resultLabel.Text = "Full resolution content for the selected VBA file";
            this.resultLabel.AutoSize = true;
            this.conflictList.Dock = System.Windows.Forms.DockStyle.Fill;
            this.baseContent.Dock = System.Windows.Forms.DockStyle.Fill;
            this.conflictDiff.Dock = System.Windows.Forms.DockStyle.Fill;
            this.conflictActions.Dock = System.Windows.Forms.DockStyle.Fill;
            this.conflictLayout.Controls.Add(this.conflictList, 0, 0);
            this.conflictLayout.Controls.Add(this.ancestorLabel, 0, 1);
            this.conflictLayout.Controls.Add(this.baseContent, 0, 2);
            this.conflictLayout.Controls.Add(this.conflictDiff, 0, 3);
            this.conflictLayout.Controls.Add(this.resultLabel, 0, 4);
            this.conflictLayout.Controls.Add(this.resolutionText, 0, 5);
            this.conflictLayout.Controls.Add(this.conflictActions, 0, 6);
            this.Controls.Add(this.conflictLayout);
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Name = "GitConflictsView";
            this.Size = new System.Drawing.Size(860, 500);
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
