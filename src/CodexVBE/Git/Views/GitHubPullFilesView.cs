using System.Windows.Forms;

namespace CodexVBE
{
    /// <summary>Vue des fichiers d’une pull request.</summary>
    public sealed partial class GitHubPullFilesView : UserControl
    {
        /// <summary>Initialise les contrôles et les textes localisés de la vue.</summary>
        public GitHubPullFilesView()
        {
            InitializeComponent();
            UiText.Apply(this, components);
        }
    }
}
