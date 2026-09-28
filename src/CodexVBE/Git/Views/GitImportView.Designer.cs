namespace CodexVBE
{
    /// <summary>Vue en lecture seule du résumé d’importation d’un dépôt Git.</summary>
public sealed partial class GitImportView
    {
        /// <summary>Affiche le résultat et les détails de l’importation.</summary>
internal System.Windows.Forms.TextBox importSummary;
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

        /// <summary>Crée la zone multiligne en lecture seule du résumé d’importation.</summary>
private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.toolTips = new System.Windows.Forms.ToolTip(this.components);
            this.importSummary = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            this.importSummary.Name = "importSummary";
            this.importSummary.Dock = System.Windows.Forms.DockStyle.Fill;
            this.importSummary.Multiline = true;
            this.importSummary.ReadOnly = true;
            this.importSummary.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.Controls.Add(this.importSummary);
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.Name = "GitImportView";
            this.Size = new System.Drawing.Size(860, 500);
            this.importSummary.Location = new System.Drawing.Point(0, 0);
            this.importSummary.Size = new System.Drawing.Size(860, 500);
            this.importSummary.TabIndex = 0;
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
