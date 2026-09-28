using System.Windows.Forms;

namespace CodexVBE
{
    /// <summary>Vue des commentaires d’une pull request.</summary>
    public sealed partial class GitHubPullCommentsView : UserControl
    {
        /// <summary>Initialise les contrôles et les textes localisés de la vue.</summary>
        public GitHubPullCommentsView()
        {
            InitializeComponent();
            UiText.Apply(this, components);
        }
    }
}
