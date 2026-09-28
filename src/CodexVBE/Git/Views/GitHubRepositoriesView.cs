using System.Windows.Forms;

namespace CodexVBE
{
    /// <summary>Vue de sélection et de création de dépôts GitHub.</summary>
    public sealed partial class GitHubRepositoriesView : UserControl
    {
        /// <summary>Initialise les contrôles et les textes localisés de la vue.</summary>
        public GitHubRepositoriesView()
        {
            InitializeComponent();
            UiText.Apply(this, components);
        }
    }
}
