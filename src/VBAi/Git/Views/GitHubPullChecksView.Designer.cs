namespace VBAi
{
    /// <summary>Vue en lecture seule des vérifications et statuts d’une pull request.</summary>
    public sealed partial class GitHubPullChecksView
    {
        /// <summary>Affiche le résumé des exécutions de vérification et leur état.</summary>
        internal VBAi.UiTextBox checks;
        /// <summary>Conteneur des composants managés de la vue.</summary>
        private System.ComponentModel.IContainer components;
        /// <summary>Fournit les info-bulles des contrôles.</summary>
        private System.Windows.Forms.ToolTip toolTips;
        /// <summary>Libère les composants de la vue avant son contrôle natif.</summary>
        /// <param name="disposing">Indique si les ressources managées doivent être libérées.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing) components?.Dispose();
            base.Dispose(disposing);
        }

        /// <summary>Crée la zone de texte en lecture seule des vérifications.</summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.toolTips = new System.Windows.Forms.ToolTip(this.components);
            this.checks = new VBAi.UiTextBox();
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
            this.checks.Location = new System.Drawing.Point(0, 0);
            this.checks.Size = new System.Drawing.Size(860, 500);
            this.checks.TabIndex = 0;
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
