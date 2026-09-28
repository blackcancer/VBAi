using System.Windows.Forms;

namespace CodexVBE
{
    public sealed partial class GitHubPullFilesView : UserControl
    {
        public GitHubPullFilesView()
        {
            InitializeComponent();
            UiText.Apply(this, components);
        }
    }
}
