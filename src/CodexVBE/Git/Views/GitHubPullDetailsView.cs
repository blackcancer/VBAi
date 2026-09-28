using System.Windows.Forms;

namespace CodexVBE
{
    public sealed partial class GitHubPullDetailsView : UserControl
    {
        public GitHubPullDetailsView()
        {
            InitializeComponent();
            UiText.Apply(this, components);
        }
    }
}
