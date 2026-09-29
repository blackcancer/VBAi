using System.Windows.Forms;

namespace VBAi
{
    /// <summary>Boîte de dialogue modale qui présente une modification VBA et permet de l’autoriser ou de la refuser.</summary>
    internal sealed partial class VbeApprovalDialog : Form
    {
        /// <summary>Initialise la boîte de dialogue et applique les libellés localisés.</summary>
        public VbeApprovalDialog()
        {
            InitializeComponent();
            UiText.Apply(this, null);
        }

        /// <summary>Initialise la boîte de dialogue avec le résumé de la modification à examiner.</summary>
        /// <param name="summary">Texte affiché dans le champ de détails.</param>
        public VbeApprovalDialog(string summary) : this()
        {
            details.Text = summary;
        }
    }
}
