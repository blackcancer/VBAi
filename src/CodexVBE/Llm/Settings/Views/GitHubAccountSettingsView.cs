using System.Windows.Forms;

namespace CodexVBE
{
    public sealed partial class GitHubAccountSettingsView : UserControl
    {
        public GitHubAccountSettingsView()
        {
            InitializeComponent();
            UiText.Apply(this, components);
        }
    }
}
