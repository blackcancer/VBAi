namespace CodexVBE
{
    public sealed partial class GitHubPullChecksView
    {
        internal System.Windows.Forms.TextBox checks;
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
            this.checks = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            this.checks.Name = "checks";
            this.checks.Multiline = true;
            this.checks.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.checks.Dock = System.Windows.Forms.DockStyle.Fill;
            this.checks.ReadOnly = true;
            this.Controls.Add(this.checks);
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Name = "GitHubPullChecksView";
            this.Size = new System.Drawing.Size(860, 500);
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
