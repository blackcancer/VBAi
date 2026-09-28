namespace CodexVBE
{
    /// <summary>Designer-generated controls for creating and restoring local project checkpoints.</summary>
    public sealed partial class GitCheckpointsView
    {
        /// <summary>Actions for checkpoint operations.</summary>
        internal System.Windows.Forms.FlowLayoutPanel checkpointActions;
        /// <summary>Saved local checkpoints.</summary>
        internal System.Windows.Forms.ListBox checkpointList;
        /// <summary>New checkpoint name input.</summary>
        internal System.Windows.Forms.TextBox checkpointName;
        /// <summary>Creates a checkpoint from the current project state.</summary>
        internal CodexVBE.ThemedButton checkpointCreate;
        /// <summary>Restores the selected checkpoint.</summary>
        internal CodexVBE.ThemedButton checkpointRestore;
        /// <summary>Container that owns Designer components.</summary>
        private System.ComponentModel.IContainer components;
        /// <summary>Tooltips associated with checkpoint actions.</summary>
        private System.Windows.Forms.ToolTip toolTips;
        /// <summary>Libère les composants de la vue avant son contrôle natif.</summary>
        /// <param name="disposing">Indique si les ressources managées doivent être libérées.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing) components?.Dispose();
            base.Dispose(disposing);
        }

        /// <summary>Creates and arranges checkpoint controls.</summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.toolTips = new System.Windows.Forms.ToolTip(this.components);
            this.checkpointActions = new System.Windows.Forms.FlowLayoutPanel();
            this.checkpointList = new System.Windows.Forms.ListBox();
            this.checkpointName = new System.Windows.Forms.TextBox();
            this.checkpointCreate = new CodexVBE.ThemedButton();
            this.checkpointRestore = new CodexVBE.ThemedButton();
            this.checkpointActions.SuspendLayout();
            this.SuspendLayout();
            this.checkpointActions.Name = "checkpointActions";
            this.checkpointCreate.Name = "checkpointCreate";
            this.checkpointRestore.Name = "checkpointRestore";
            this.checkpointList.Dock = System.Windows.Forms.DockStyle.Fill;
            this.checkpointList.Name = "checkpointList";
            this.checkpointList.HorizontalScrollbar = true;
            this.checkpointActions.Dock = System.Windows.Forms.DockStyle.Top;
            this.checkpointActions.AutoSize = true;
            this.checkpointActions.Controls.Add(this.checkpointName);
            this.checkpointActions.Controls.Add(this.checkpointCreate);
            this.checkpointActions.Controls.Add(this.checkpointRestore);
            this.checkpointName.Width = 240;
            this.checkpointName.Name = "checkpointName";
            this.checkpointName.AccessibleName = "Checkpoint name";
            this.checkpointCreate.Text = "Create checkpoint";
            this.checkpointCreate.AutoSize = true;
            this.checkpointRestore.Text = "Restore selection";
            this.checkpointRestore.AutoSize = true;
            this.toolTips.SetToolTip(this.checkpointCreate, "Save all current VBA in a private checkpoint without committing to the published branch.");
            this.toolTips.SetToolTip(this.checkpointRestore, "Restore the selected checkpoint after backing up the current state. Does not rewrite GitHub.");
            this.Controls.Add(this.checkpointList);
            this.Controls.Add(this.checkpointActions);
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Name = "GitCheckpointsView";
            this.Size = new System.Drawing.Size(860, 500);
            this.checkpointActions.Location = new System.Drawing.Point(0, 0);
            this.checkpointActions.Size = new System.Drawing.Size(860, 34);
            this.checkpointActions.TabIndex = 1;
            this.checkpointList.Location = new System.Drawing.Point(0, 34);
            this.checkpointList.Size = new System.Drawing.Size(860, 466);
            this.checkpointList.TabIndex = 0;
            this.checkpointName.Location = new System.Drawing.Point(3, 3);
            this.checkpointName.Size = new System.Drawing.Size(240, 23);
            this.checkpointName.TabIndex = 0;
            this.checkpointCreate.Location = new System.Drawing.Point(249, 3);
            this.checkpointCreate.Size = new System.Drawing.Size(113, 28);
            this.checkpointCreate.TabIndex = 1;
            this.checkpointRestore.Location = new System.Drawing.Point(368, 3);
            this.checkpointRestore.Size = new System.Drawing.Size(107, 28);
            this.checkpointRestore.TabIndex = 2;
            this.checkpointActions.ResumeLayout(false);
            this.checkpointActions.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
