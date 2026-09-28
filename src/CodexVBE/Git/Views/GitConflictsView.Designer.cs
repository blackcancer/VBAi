namespace CodexVBE
{
    /// <summary>Designer-generated conflict list, resolution editor, and side-by-side comparison.</summary>
    public sealed partial class GitConflictsView
    {
        /// <summary>Conflict resolution actions.</summary>
        internal System.Windows.Forms.FlowLayoutPanel conflictActions;
        /// <summary>Files with unresolved conflicts.</summary>
        internal System.Windows.Forms.ListBox conflictList;
        /// <summary>Manual conflict resolution text.</summary>
        internal System.Windows.Forms.TextBox resolutionText;
        /// <summary>Uses the current branch's version.</summary>
        internal CodexVBE.ThemedButton mergeOurs;
        /// <summary>Uses the incoming branch's version.</summary>
        internal CodexVBE.ThemedButton mergeTheirs;
        /// <summary>Applies the manual resolution text.</summary>
        internal CodexVBE.ThemedButton mergeText;
        /// <summary>Completes the merge after all conflicts are resolved.</summary>
        internal CodexVBE.ThemedButton mergeComplete;
        /// <summary>Aborts the current merge.</summary>
        internal CodexVBE.ThemedButton mergeAbort;
        /// <summary>Comparison grid for the conflicting versions.</summary>
        internal System.Windows.Forms.DataGridView conflictDiff;
        /// <summary>Current branch's conflict content.</summary>
        internal System.Windows.Forms.DataGridViewTextBoxColumn conflictOurs;
        /// <summary>Incoming branch's conflict content.</summary>
        internal System.Windows.Forms.DataGridViewTextBoxColumn conflictTheirs;
        /// <summary>Common ancestor version of the conflict.</summary>
        internal System.Windows.Forms.TextBox baseContent;
        /// <summary>Layout for ancestor and resolved-result content.</summary>
        internal System.Windows.Forms.TableLayoutPanel conflictLayout;
        /// <summary>Caption for the common ancestor content.</summary>
        internal System.Windows.Forms.Label ancestorLabel;
        /// <summary>Caption for the resolved content.</summary>
        internal System.Windows.Forms.Label resultLabel;
        /// <summary>Container that owns Designer components.</summary>
        private System.ComponentModel.IContainer components;
        /// <summary>Tooltips associated with conflict actions.</summary>
        private System.Windows.Forms.ToolTip toolTips;
        /// <summary>Libère les composants de la vue avant son contrôle natif.</summary>
        /// <param name="disposing">Indique si les ressources managées doivent être libérées.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing) components?.Dispose();
            base.Dispose(disposing);
        }

        /// <summary>Creates and arranges conflict controls and their comparison grid.</summary>
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
            ((System.ComponentModel.ISupportInitialize)(this.conflictDiff)).BeginInit();
            this.conflictActions.SuspendLayout();
            this.conflictDiff.SuspendLayout();
            this.conflictLayout.SuspendLayout();
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
            this.conflictActions.Location = new System.Drawing.Point(3, 462);
            this.conflictActions.Size = new System.Drawing.Size(854, 35);
            this.conflictActions.TabIndex = 6;
            this.conflictList.Location = new System.Drawing.Point(3, 3);
            this.conflictList.Size = new System.Drawing.Size(854, 42);
            this.conflictList.TabIndex = 0;
            this.resolutionText.Location = new System.Drawing.Point(3, 314);
            this.resolutionText.Size = new System.Drawing.Size(854, 142);
            this.resolutionText.TabIndex = 5;
            this.mergeOurs.Location = new System.Drawing.Point(3, 3);
            this.mergeOurs.Size = new System.Drawing.Size(75, 28);
            this.mergeOurs.TabIndex = 0;
            this.mergeTheirs.Location = new System.Drawing.Point(84, 3);
            this.mergeTheirs.Size = new System.Drawing.Size(96, 28);
            this.mergeTheirs.TabIndex = 1;
            this.mergeText.Location = new System.Drawing.Point(186, 3);
            this.mergeText.Size = new System.Drawing.Size(81, 28);
            this.mergeText.TabIndex = 2;
            this.mergeComplete.Location = new System.Drawing.Point(273, 3);
            this.mergeComplete.Size = new System.Drawing.Size(106, 28);
            this.mergeComplete.TabIndex = 3;
            this.mergeAbort.Location = new System.Drawing.Point(385, 3);
            this.mergeAbort.Size = new System.Drawing.Size(84, 28);
            this.mergeAbort.TabIndex = 4;
            this.conflictDiff.Location = new System.Drawing.Point(3, 164);
            this.conflictDiff.Size = new System.Drawing.Size(854, 123);
            this.conflictDiff.TabIndex = 3;
            this.baseContent.Location = new System.Drawing.Point(3, 72);
            this.baseContent.Size = new System.Drawing.Size(854, 86);
            this.baseContent.TabIndex = 2;
            this.conflictLayout.Location = new System.Drawing.Point(0, 0);
            this.conflictLayout.Size = new System.Drawing.Size(860, 500);
            this.conflictLayout.TabIndex = 0;
            this.ancestorLabel.Location = new System.Drawing.Point(3, 48);
            this.ancestorLabel.Size = new System.Drawing.Size(105, 21);
            this.ancestorLabel.TabIndex = 1;
            this.resultLabel.Location = new System.Drawing.Point(3, 290);
            this.resultLabel.Size = new System.Drawing.Size(261, 21);
            this.resultLabel.TabIndex = 4;
            ((System.ComponentModel.ISupportInitialize)(this.conflictDiff)).EndInit();
            this.conflictActions.ResumeLayout(false);
            this.conflictActions.PerformLayout();
            this.conflictDiff.ResumeLayout(false);
            this.conflictDiff.PerformLayout();
            this.conflictLayout.ResumeLayout(false);
            this.conflictLayout.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
