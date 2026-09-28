using System.Windows.Forms;

namespace CodexVBE
{
    public sealed partial class GitHubPullChecksView : UserControl
    {
        public GitHubPullChecksView()
        {
            InitializeComponent();
            UiText.Apply(this, components);
        }
    }
}
