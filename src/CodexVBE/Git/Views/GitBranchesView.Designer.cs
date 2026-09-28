namespace CodexVBE
{
    public sealed partial class GitBranchesView
    {
        internal System.Windows.Forms.FlowLayoutPanel branchActions;
        internal System.Windows.Forms.ListBox branchList;
        internal System.Windows.Forms.ComboBox branchName;
        internal System.Windows.Forms.Button branchRemote;
        internal System.Windows.Forms.Button branchCreate;
        internal System.Windows.Forms.Button branchTrack;
        internal System.Windows.Forms.Button branchSwitch;
        internal System.Windows.Forms.Button mergeBegin;
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
            this.branchActions = new System.Windows.Forms.FlowLayoutPanel();
            this.branchList = new System.Windows.Forms.ListBox();
            this.branchName = new CodexVBE.ThemedComboBox();
            this.branchRemote = new CodexVBE.ThemedButton();
            this.branchCreate = new CodexVBE.ThemedButton();
            this.branchTrack = new CodexVBE.ThemedButton();
            this.branchSwitch = new CodexVBE.ThemedButton();
            this.mergeBegin = new CodexVBE.ThemedButton();
            this.SuspendLayout();
            this.branchActions.Name = "branchActions";
            this.branchCreate.Name = "branchCreate";
            this.branchRemote.Name = "branchRemote";
            this.branchTrack.Name = "branchTrack";
            this.branchSwitch.Name = "branchSwitch";
            this.mergeBegin.Name = "mergeBegin";
            this.branchList.Dock = System.Windows.Forms.DockStyle.Fill;
            this.branchList.Name = "branchList";
            this.branchActions.Dock = System.Windows.Forms.DockStyle.Top;
            this.branchActions.AutoSize = true;
            this.branchActions.Controls.Add(this.branchName);
            this.branchActions.Controls.Add(this.branchCreate);
            this.branchActions.Controls.Add(this.branchRemote);
            this.branchActions.Controls.Add(this.branchTrack);
            this.branchActions.Controls.Add(this.branchSwitch);
            this.branchActions.Controls.Add(this.mergeBegin);
            this.branchName.Width = 170;
            this.branchName.Name = "branchName";
            this.branchName.AccessibleName = "Branch name to create or track";
            this.branchCreate.Text = "Create";
            this.branchCreate.AutoSize = true;
            this.branchRemote.Text = "Remote branches";
            this.branchRemote.AutoSize = true;
            this.toolTips.SetToolTip(this.branchRemote, "List GitHub repository branches in the name selector without importing VBA.");
            this.branchTrack.Text = "Track remote";
            this.branchTrack.AutoSize = true;
            this.branchSwitch.Text = "Switch";
            this.branchSwitch.AutoSize = true;
            this.mergeBegin.Text = "Merge";
            this.mergeBegin.AutoSize = true;
            this.toolTips.SetToolTip(this.branchCreate, "Create the named branch from the current commit without publishing it.");
            this.toolTips.SetToolTip(this.branchTrack, "Fetch the named remote branch without importing its VBA.");
            this.toolTips.SetToolTip(this.branchSwitch, "Import the selected branch with a checkpoint first. Uncommitted changes are refused.");
            this.toolTips.SetToolTip(this.mergeBegin, "Prepare a merge from the selected branch into the current branch.");
            this.Controls.Add(this.branchList);
            this.Controls.Add(this.branchActions);
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Name = "GitBranchesView";
            this.Size = new System.Drawing.Size(860, 500);
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
