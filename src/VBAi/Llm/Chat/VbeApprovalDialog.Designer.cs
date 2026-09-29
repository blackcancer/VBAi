namespace VBAi
{
    /// <summary>Déclare les contrôles et leur disposition pour la boîte de dialogue d’approbation.</summary>
    internal sealed partial class VbeApprovalDialog
    {
        /// <summary>Champ en lecture seule qui affiche le résumé de l’édition.</summary>
        private VBAi.UiTextBox details;
        /// <summary>Barre inférieure qui contient les boutons d’approbation et de refus.</summary>
        private System.Windows.Forms.FlowLayoutPanel actions;
        /// <summary>Bouton dont le résultat de dialogue autorise l’édition.</summary>
        private VBAi.UiActionButton approve;
        /// <summary>Bouton dont le résultat de dialogue refuse l’édition.</summary>
        private VBAi.UiActionButton reject;

        /// <summary>Crée et configure les contrôles, leurs résultats de dialogue et les dimensions de la fenêtre.</summary>
        private void InitializeComponent()
        {
            this.details = new VBAi.UiTextBox();
            this.actions = new System.Windows.Forms.FlowLayoutPanel();
            this.approve = new VBAi.UiActionButton();
            this.reject = new VBAi.UiActionButton();
            this.actions.SuspendLayout();
            this.SuspendLayout();
            //
            // details
            //
            this.details.Dock = System.Windows.Forms.DockStyle.Fill;
            this.details.Multiline = true;
            this.details.Name = "details";
            this.details.ReadOnly = true;
            this.details.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.details.TabIndex = 0;
            this.details.WordWrap = false;
            //
            // actions
            //
            this.actions.Controls.Add(this.approve);
            this.actions.Controls.Add(this.reject);
            this.actions.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.actions.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft;
            this.actions.Height = 44;
            this.actions.Name = "actions";
            this.actions.TabIndex = 1;
            //
            // approve
            //
            this.approve.DialogResult = System.Windows.Forms.DialogResult.Yes;
            this.approve.Name = "approve";
            this.approve.Size = new System.Drawing.Size(100, 30);
            this.approve.TabIndex = 0;
            this.approve.Text = "Allow";
            //
            // reject
            //
            this.reject.DialogResult = System.Windows.Forms.DialogResult.No;
            this.reject.Name = "reject";
            this.reject.Size = new System.Drawing.Size(100, 30);
            this.reject.TabIndex = 1;
            this.reject.Text = "Deny";
            //
            // VbeApprovalDialog
            //
            this.CancelButton = this.reject;
            this.ClientSize = new System.Drawing.Size(740, 530);
            this.Controls.Add(this.details);
            this.Controls.Add(this.actions);
            this.MaximizeBox = true;
            this.MinimizeBox = false;
            this.Name = "VbeApprovalDialog";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "VBAi — approve edit";
            this.details.Location = new System.Drawing.Point(0, 0);
            this.details.Size = new System.Drawing.Size(740, 486);
            this.actions.Location = new System.Drawing.Point(0, 486);
            this.actions.Size = new System.Drawing.Size(740, 44);
            this.approve.Location = new System.Drawing.Point(637, 3);
            this.reject.Location = new System.Drawing.Point(531, 3);
            this.actions.ResumeLayout(false);
            this.actions.PerformLayout();
            this.approve.Symbol = VBAi.UiSymbol.Check;
            this.reject.Symbol = VBAi.UiSymbol.Close;
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
