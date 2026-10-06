namespace VBAi
{

    /// <summary>Designer-generated controls for listing, tracking, creating, and merging branches.</summary>
    public sealed partial class GitBranchesView
    {

        /// <summary>Actions for branch and merge operations.</summary>
        internal System.Windows.Forms.FlowLayoutPanel branchActions;

        /// <summary>Local and remote branches.</summary>
        internal VBAi.UiListBox branchList;

        /// <summary>Branch name input or selection.</summary>
        internal VBAi.ThemedComboBox branchName;

        /// <summary>Refreshes available remote branches.</summary>
        internal VBAi.ThemedButton branchRemote;

        /// <summary>Creates the named branch.</summary>
        internal VBAi.ThemedButton branchCreate;

        /// <summary>Tracks the selected remote branch.</summary>
        internal VBAi.ThemedButton branchTrack;

        /// <summary>Switches to the selected branch.</summary>
        internal VBAi.ThemedButton branchSwitch;

        /// <summary>Begins a merge of the selected branch.</summary>
        internal VBAi.ThemedButton mergeBegin;

        /// <summary>Container that owns Designer components.</summary>
        private System.ComponentModel.IContainer components;

        /// <summary>Tooltips associated with branch actions.</summary>
        private System.Windows.Forms.ToolTip toolTips;

        /// <summary>Libère les composants de la vue avant son contrôle natif.</summary>
        /// <param name="disposing">Indique si les ressources managées doivent être libérées.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing) components?.Dispose();
            base.Dispose(disposing);
        }

        /// <summary>Creates and arranges branch management controls.</summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.toolTips = new System.Windows.Forms.ToolTip(this.components);
            this.branchActions = new System.Windows.Forms.FlowLayoutPanel();
            this.branchList = new VBAi.UiListBox();
            this.branchName = new VBAi.ThemedComboBox();
            this.branchRemote = new VBAi.ThemedButton();
            this.branchCreate = new VBAi.ThemedButton();
            this.branchTrack = new VBAi.ThemedButton();
            this.branchSwitch = new VBAi.ThemedButton();
            this.mergeBegin = new VBAi.ThemedButton();
            this.branchActions.SuspendLayout();
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
            this.branchActions.Location = new System.Drawing.Point(0, 0);
            this.branchActions.Size = new System.Drawing.Size(860, 34);
            this.branchActions.TabIndex = 1;
            this.branchList.Location = new System.Drawing.Point(0, 34);
            this.branchList.Size = new System.Drawing.Size(860, 466);
            this.branchList.TabIndex = 0;
            this.branchName.Location = new System.Drawing.Point(3, 3);
            this.branchName.Size = new System.Drawing.Size(170, 21);
            this.branchName.TabIndex = 0;
            this.branchRemote.Location = new System.Drawing.Point(260, 3);
            this.branchRemote.Size = new System.Drawing.Size(110, 28);
            this.branchRemote.TabIndex = 2;
            this.branchCreate.Location = new System.Drawing.Point(179, 3);
            this.branchCreate.Size = new System.Drawing.Size(75, 28);
            this.branchCreate.TabIndex = 1;
            this.branchTrack.Location = new System.Drawing.Point(376, 3);
            this.branchTrack.Size = new System.Drawing.Size(86, 28);
            this.branchTrack.TabIndex = 3;
            this.branchSwitch.Location = new System.Drawing.Point(468, 3);
            this.branchSwitch.Size = new System.Drawing.Size(75, 28);
            this.branchSwitch.TabIndex = 4;
            this.mergeBegin.Location = new System.Drawing.Point(549, 3);
            this.mergeBegin.Size = new System.Drawing.Size(75, 28);
            this.mergeBegin.TabIndex = 5;
            this.branchActions.ResumeLayout(false);
            this.branchActions.PerformLayout();
            this.branchRemote.Symbol = VBAi.UiSymbol.Refresh;
            this.branchRemote.IconOnly = true;
            this.branchRemote.AutoSize = false;
            this.branchRemote.MinimumSize = System.Drawing.Size.Empty;
            this.branchRemote.Size = new System.Drawing.Size(32, 30);
            this.branchCreate.Symbol = VBAi.UiSymbol.Add;
            this.branchCreate.IconOnly = true;
            this.branchCreate.AutoSize = false;
            this.branchCreate.MinimumSize = System.Drawing.Size.Empty;
            this.branchCreate.Size = new System.Drawing.Size(32, 30);
            this.branchTrack.Symbol = VBAi.UiSymbol.Download;
            this.branchTrack.IconOnly = true;
            this.branchTrack.AutoSize = false;
            this.branchTrack.MinimumSize = System.Drawing.Size.Empty;
            this.branchTrack.Size = new System.Drawing.Size(32, 30);
            this.branchSwitch.Symbol = VBAi.UiSymbol.Next;
            this.branchSwitch.IconOnly = true;
            this.branchSwitch.AutoSize = false;
            this.branchSwitch.MinimumSize = System.Drawing.Size.Empty;
            this.branchSwitch.Size = new System.Drawing.Size(32, 30);
            this.mergeBegin.Symbol = VBAi.UiSymbol.Code;
            this.mergeBegin.IconOnly = true;
            this.mergeBegin.AutoSize = false;
            this.mergeBegin.MinimumSize = System.Drawing.Size.Empty;
            this.mergeBegin.Size = new System.Drawing.Size(32, 30);
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
