using System.Windows.Forms;

namespace CodexVBE
{
    internal sealed partial class VbeApprovalDialog : Form
    {
        public VbeApprovalDialog()
        {
            InitializeComponent();
            UiText.Apply(this, null);
        }

        public VbeApprovalDialog(string summary) : this()
        {
            details.Text = summary;
        }
    }
}
