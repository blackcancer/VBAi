using System.Windows.Forms;

namespace CodexVBE
{
    public sealed partial class GitImportView : UserControl
    {
        public GitImportView()
        {
            InitializeComponent();
            UiText.Apply(this, components);
        }
    }
}
