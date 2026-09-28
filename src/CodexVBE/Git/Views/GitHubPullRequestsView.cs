using System.Windows.Forms;

namespace CodexVBE
{
    /// <summary>Vue des pull requests GitHub.</summary>
    public sealed partial class GitHubPullRequestsView : UserControl
    {
        /// <summary>Initialise les contrôles et les textes localisés de la vue.</summary>
        public GitHubPullRequestsView()
        {
            InitializeComponent();
            UiText.Apply(this, components);
        }
    }
}
