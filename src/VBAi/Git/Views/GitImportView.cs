using System.Windows.Forms;

namespace VBAi
{

    /// <summary>Vue de préparation des imports VBA.</summary>
    public sealed partial class GitImportView : UserControl
    {

        /// <summary>Initialise les contrôles et les textes localisés de la vue.</summary>
        public GitImportView()
        {
            InitializeComponent();
            UiText.Apply(this, components);
        }
    }
}
