using System.Windows.Forms;

namespace CodexVBE
{
    public sealed partial class GitHubPullCommentsView : UserControl
    {
        public GitHubPullCommentsView()
        {
            InitializeComponent();
            UiText.Apply(this, components);
        }
    }
}
