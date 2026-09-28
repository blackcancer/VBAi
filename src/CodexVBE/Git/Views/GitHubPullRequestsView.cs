using System.Windows.Forms;

namespace CodexVBE
{
    public sealed partial class GitHubPullRequestsView : UserControl
    {
        public GitHubPullRequestsView()
        {
            InitializeComponent();
            UiText.Apply(this, components);
        }
    }
}
