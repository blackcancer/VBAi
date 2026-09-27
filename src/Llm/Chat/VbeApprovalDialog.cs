using System.Windows.Forms;

namespace CodexVBE
{
    internal sealed partial class VbeApprovalDialog : Form
    {
        public VbeApprovalDialog()
        {
            InitializeComponent();
        }

        public VbeApprovalDialog(string summary) : this()
        {
            details.Text = summary;
        }
    }
}
