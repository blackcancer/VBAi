using System.Windows.Forms;

namespace CodexVBE
{
    public sealed partial class GitBranchesView : UserControl
    {
        public GitBranchesView()
        {
            InitializeComponent();
            UiText.Apply(this, components);
        }
    }
}
