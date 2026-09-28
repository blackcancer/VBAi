namespace CodexVBE
{
    public sealed partial class GitHubPullComposeView
    {
        internal System.Windows.Forms.TableLayoutPanel composeLayout;
        internal CodexVBE.ThemedComboBox targetBranch;
        internal System.Windows.Forms.Label sourceLabel;
        internal System.Windows.Forms.TextBox pullTitle;
        internal System.Windows.Forms.TextBox pullBody;
        internal System.Windows.Forms.CheckBox draft;
        internal CodexVBE.ThemedButton createPull;
        internal System.Windows.Forms.Label targetLabel;
        internal System.Windows.Forms.Label titleLabel;
        internal System.Windows.Forms.Label bodyLabel;
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
            this.composeLayout = new System.Windows.Forms.TableLayoutPanel();
            this.targetBranch = new CodexVBE.ThemedComboBox();
            this.sourceLabel = new System.Windows.Forms.Label();
            this.pullTitle = new System.Windows.Forms.TextBox();
            this.pullBody = new System.Windows.Forms.TextBox();
            this.draft = new System.Windows.Forms.CheckBox();
            this.createPull = new CodexVBE.ThemedButton();
            this.targetLabel = new System.Windows.Forms.Label();
            this.titleLabel = new System.Windows.Forms.Label();
            this.bodyLabel = new System.Windows.Forms.Label();
            this.composeLayout.SuspendLayout();
            this.SuspendLayout();
            this.composeLayout.Name = "composeLayout";
            this.targetBranch.Name = "targetBranch";
            this.sourceLabel.Name = "sourceLabel";
            this.pullTitle.Name = "pullTitle";
            this.pullBody.Name = "pullBody";
            this.draft.Name = "draft";
            this.createPull.Name = "createPull";
            this.composeLayout.ColumnCount = 1;
            this.composeLayout.RowCount = 6;
            this.composeLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.composeLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize, 0F));
            this.composeLayout.Controls.Add(this.sourceLabel, 0, 0);
            this.sourceLabel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.composeLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.composeLayout.Controls.Add(this.targetBranch, 0, 1);
            this.targetBranch.Dock = System.Windows.Forms.DockStyle.Fill;
            this.composeLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.composeLayout.Controls.Add(this.pullTitle, 0, 2);
            this.pullTitle.Dock = System.Windows.Forms.DockStyle.Fill;
            this.composeLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.composeLayout.Controls.Add(this.pullBody, 0, 3);
            this.pullBody.Dock = System.Windows.Forms.DockStyle.Fill;
            this.composeLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize, 0F));
            this.composeLayout.Controls.Add(this.draft, 0, 4);
            this.draft.Dock = System.Windows.Forms.DockStyle.Fill;
            this.composeLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize, 0F));
            this.composeLayout.Controls.Add(this.createPull, 0, 5);
            this.createPull.Dock = System.Windows.Forms.DockStyle.Fill;
            this.targetBranch.AccessibleName = "Target branch";
            this.pullTitle.AccessibleName = "Pull request title";
            this.pullBody.AccessibleName = "Pull request description";
            this.draft.Text = "Draft";
            this.draft.Checked = true;
            this.draft.AutoSize = true;
            this.createPull.Text = "Create pull request";
            this.createPull.AutoSize = true;
            this.pullBody.Multiline = true;
            this.pullBody.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.pullBody.Dock = System.Windows.Forms.DockStyle.Fill;
            this.targetLabel.Text = "Target branch";
            this.titleLabel.Text = "Pull request title";
            this.bodyLabel.Text = "Pull request description";
            this.targetLabel.AutoSize = true;
            this.titleLabel.AutoSize = true;
            this.bodyLabel.AutoSize = true;
            this.composeLayout.ColumnCount = 2;
            this.composeLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 150F));
            this.composeLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.composeLayout.SetColumn(this.targetBranch, 1);
            this.composeLayout.SetColumn(this.pullTitle, 1);
            this.composeLayout.SetColumn(this.pullBody, 1);
            this.composeLayout.SetColumnSpan(this.sourceLabel, 2);
            this.composeLayout.SetColumnSpan(this.createPull, 2);
            this.composeLayout.Controls.Add(this.targetLabel, 0, 1);
            this.composeLayout.Controls.Add(this.titleLabel, 0, 2);
            this.composeLayout.Controls.Add(this.bodyLabel, 0, 3);
            this.Controls.Add(this.composeLayout);
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Name = "GitHubPullComposeView";
            this.Size = new System.Drawing.Size(860, 500);
            this.composeLayout.Location = new System.Drawing.Point(0, 0);
            this.composeLayout.Size = new System.Drawing.Size(860, 500);
            this.composeLayout.TabIndex = 0;
            this.targetBranch.Location = new System.Drawing.Point(153, 26);
            this.targetBranch.Size = new System.Drawing.Size(704, 23);
            this.targetBranch.TabIndex = 1;
            this.sourceLabel.Location = new System.Drawing.Point(3, 0);
            this.sourceLabel.Size = new System.Drawing.Size(854, 23);
            this.sourceLabel.TabIndex = 0;
            this.pullTitle.Location = new System.Drawing.Point(153, 56);
            this.pullTitle.Size = new System.Drawing.Size(704, 23);
            this.pullTitle.TabIndex = 2;
            this.pullBody.Location = new System.Drawing.Point(153, 86);
            this.pullBody.Size = new System.Drawing.Size(704, 349);
            this.pullBody.TabIndex = 3;
            this.draft.Location = new System.Drawing.Point(3, 441);
            this.draft.Size = new System.Drawing.Size(144, 22);
            this.draft.TabIndex = 4;
            this.createPull.Location = new System.Drawing.Point(3, 469);
            this.createPull.Size = new System.Drawing.Size(854, 28);
            this.createPull.TabIndex = 5;
            this.targetLabel.Name = "targetLabel";
            this.targetLabel.Location = new System.Drawing.Point(3, 23);
            this.targetLabel.Size = new System.Drawing.Size(80, 21);
            this.targetLabel.TabIndex = 6;
            this.titleLabel.Name = "titleLabel";
            this.titleLabel.Location = new System.Drawing.Point(3, 53);
            this.titleLabel.Size = new System.Drawing.Size(93, 21);
            this.titleLabel.TabIndex = 7;
            this.bodyLabel.Name = "bodyLabel";
            this.bodyLabel.Location = new System.Drawing.Point(3, 83);
            this.bodyLabel.Size = new System.Drawing.Size(133, 21);
            this.bodyLabel.TabIndex = 8;
            this.composeLayout.ResumeLayout(false);
            this.composeLayout.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
