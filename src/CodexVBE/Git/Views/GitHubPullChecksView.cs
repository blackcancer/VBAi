using System.Windows.Forms;

namespace CodexVBE
{
    /// <summary>Vue des vérifications d’une pull request.</summary>
    public sealed partial class GitHubPullChecksView : UserControl
    {
        /// <summary>Initialise les contrôles et les textes localisés de la vue.</summary>
        public GitHubPullChecksView()
        {
            InitializeComponent();
            UiText.Apply(this, components);
        }
    }
}
