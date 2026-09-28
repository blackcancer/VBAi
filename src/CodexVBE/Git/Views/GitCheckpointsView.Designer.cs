namespace CodexVBE
{
    public sealed partial class GitCheckpointsView
    {
        internal System.Windows.Forms.FlowLayoutPanel checkpointActions;
        internal System.Windows.Forms.ListBox checkpointList;
        internal System.Windows.Forms.TextBox checkpointName;
        internal System.Windows.Forms.Button checkpointCreate;
        internal System.Windows.Forms.Button checkpointRestore;
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
            this.checkpointActions = new System.Windows.Forms.FlowLayoutPanel();
            this.checkpointList = new System.Windows.Forms.ListBox();
            this.checkpointName = new System.Windows.Forms.TextBox();
            this.checkpointCreate = new CodexVBE.ThemedButton();
            this.checkpointRestore = new CodexVBE.ThemedButton();
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
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
