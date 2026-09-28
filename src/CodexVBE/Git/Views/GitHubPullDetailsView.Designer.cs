namespace CodexVBE
{
    public sealed partial class GitHubPullDetailsView
    {
        internal System.Windows.Forms.TextBox pullDetails;
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
            this.pullDetails = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            this.pullDetails.Name = "pullDetails";
            this.pullDetails.Multiline = true;
            this.pullDetails.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.pullDetails.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pullDetails.ReadOnly = true;
            this.Controls.Add(this.pullDetails);
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Name = "GitHubPullDetailsView";
            this.Size = new System.Drawing.Size(860, 500);
            this.pullDetails.Location = new System.Drawing.Point(0, 0);
            this.pullDetails.Size = new System.Drawing.Size(860, 500);
            this.pullDetails.TabIndex = 0;
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
