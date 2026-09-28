using System.Windows.Forms;

namespace CodexVBE
{
    public sealed partial class GitHubRepositoriesView : UserControl
    {
        public GitHubRepositoriesView()
        {
            InitializeComponent();
            UiText.Apply(this, components);
        }
    }
}
