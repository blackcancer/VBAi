using System.Windows.Forms;

namespace CodexVBE
{
    public sealed partial class GitChangesView : UserControl
    {
        public GitChangesView()
        {
            InitializeComponent();
            UiText.Apply(this, components);
        }
    }
}
