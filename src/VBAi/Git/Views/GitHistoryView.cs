using System.Windows.Forms;

namespace VBAi
{

    /// <summary>Vue de l’historique Git.</summary>
    public sealed partial class GitHistoryView : UserControl
    {

        /// <summary>Initialise les contrôles et les textes localisés de la vue.</summary>
        public GitHistoryView()
        {
            InitializeComponent();
            UiText.Apply(this, components);
        }
    }
}
