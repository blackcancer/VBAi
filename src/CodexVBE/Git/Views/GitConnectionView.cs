using System.Windows.Forms;

namespace CodexVBE
{
    public sealed partial class GitConnectionView : UserControl
    {
        public GitConnectionView()
        {
            InitializeComponent();
            UiText.Apply(this, components);
        }
    }
}
