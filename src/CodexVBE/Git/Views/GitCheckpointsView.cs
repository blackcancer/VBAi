using System.Windows.Forms;

namespace CodexVBE
{
    public sealed partial class GitCheckpointsView : UserControl
    {
        public GitCheckpointsView()
        {
            InitializeComponent();
            UiText.Apply(this, components);
        }
    }
}
