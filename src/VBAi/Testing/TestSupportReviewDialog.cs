using System;
using System.Windows.Forms;

namespace VBAi
{

    /// <summary>Explicit review of the single project-local VBA test support edit.</summary>
    internal sealed partial class TestSupportReviewDialog : Form
    {

        /// <summary>Initializes a TestSupportReviewDialog instance with the supplied state.</summary>
        public TestSupportReviewDialog()
        {
            InitializeComponent();
            Icon = VbeWindowIcons.Icon("assistant");
            UiText.Apply(this, components);
        }

        /// <summary>Handles confirm for test support review dialog.</summary>
        /// <param name="owner">i win32 window that supplies the owner for this operation.</param>
        /// <param name="project">Text that supplies the project value. Use the format required by the calling operation.</param>
        /// <param name="before">Text that supplies the before value. Use the format required by the calling operation.</param>
        /// <param name="after">Text that supplies the after value. Use the format required by the calling operation.</param>
        /// <returns>Boolean indicating the result of the check for confirm on test support review dialog.</returns>
        internal static bool Confirm(IWin32Window owner, string project, string before, string after) => Confirm(owner, project, before, after, null);

        /// <summary>Handles confirm for test support review dialog.</summary>
        /// <param name="owner">i win32 window that supplies the owner for this operation.</param>
        /// <param name="project">Text that supplies the project value. Use the format required by the calling operation.</param>
        /// <param name="before">Text that supplies the before value. Use the format required by the calling operation.</param>
        /// <param name="after">Text that supplies the after value. Use the format required by the calling operation.</param>
        /// <param name="show">func&lt;test support review dialog, i win32 window, dialog result&gt; that supplies the show for this operation.</param>
        /// <returns>Boolean indicating the result of the check for confirm on test support review dialog.</returns>
        internal static bool Confirm(IWin32Window owner, string project, string before, string after, Func<TestSupportReviewDialog, IWin32Window, DialogResult> show)
        {
            using (var dialog = new TestSupportReviewDialog())
            {
                dialog.projectLabel.Text = UiText.Get("Project") + ": " + project;
                dialog.diff.ShowDiff(before ?? "", after ?? "");
                return (show == null ? dialog.ShowDialog(owner) : show(dialog, owner)) == DialogResult.OK;
            }
        }
    }
}
