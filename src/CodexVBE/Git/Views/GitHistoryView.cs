using System.Windows.Forms;

namespace CodexVBE
{
    public sealed partial class GitHistoryView : UserControl
    {
        public GitHistoryView()
        {
            InitializeComponent();
            UiText.Apply(this, components);
        }
    }
}
