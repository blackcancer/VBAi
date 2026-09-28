using System.Windows.Forms;

namespace CodexVBE
{
    public sealed partial class GitHubPullComposeView : UserControl
    {
        public GitHubPullComposeView()
        {
            InitializeComponent();
            UiText.Apply(this, components);
        }
    }
}
