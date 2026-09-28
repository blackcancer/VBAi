namespace CodexVBE
{
    public sealed partial class GitHubPullFilesView
    {
        internal System.Windows.Forms.ListBox files;
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
            this.files = new System.Windows.Forms.ListBox();
            this.SuspendLayout();
            this.files.Name = "files";
            this.files.HorizontalScrollbar = true;
            this.files.Dock = System.Windows.Forms.DockStyle.Fill;
            this.Controls.Add(this.files);
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Name = "GitHubPullFilesView";
            this.Size = new System.Drawing.Size(860, 500);
            this.files.Location = new System.Drawing.Point(0, 0);
            this.files.Size = new System.Drawing.Size(860, 500);
            this.files.TabIndex = 0;
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
