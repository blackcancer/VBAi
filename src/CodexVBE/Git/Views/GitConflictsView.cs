using System.Windows.Forms;

namespace CodexVBE
{
    public sealed partial class GitConflictsView : UserControl
    {
        public GitConflictsView()
        {
            InitializeComponent();
            UiText.Apply(this, components);
        }
    }
}
