using System;
using System.Windows.Forms;

namespace VBAi
{
    /// <summary>Explicit review of the single project-local VBA test support edit.</summary>
    internal sealed partial class TestSupportReviewDialog : Form
    {
        public TestSupportReviewDialog()
        {
            InitializeComponent();
            Icon = VbeWindowIcons.Icon("assistant");
            UiText.Apply(this, components);
        }

        internal static bool Confirm(IWin32Window owner, string project, string before, string after)
        {
            using (var dialog = new TestSupportReviewDialog())
            {
                dialog.projectLabel.Text = UiText.Get("Project") + ": " + project;
                dialog.diff.ShowDiff(before ?? "", after ?? "");
                return dialog.ShowDialog(owner) == DialogResult.OK;
            }
        }
    }
}
