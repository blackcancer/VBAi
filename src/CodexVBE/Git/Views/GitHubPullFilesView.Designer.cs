namespace CodexVBE
{
    /// <summary>Vue listant les fichiers modifiés par la pull request sélectionnée.</summary>
    public sealed partial class GitHubPullFilesView
    {
        /// <summary>Liste les chemins des fichiers modifiés et leur état de changement.</summary>
        internal System.Windows.Forms.ListBox files;
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

        /// <summary>Crée la liste des fichiers avec défilement horizontal.</summary>
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
