using System;
using System.Windows.Forms;

namespace VBAi
{

    /// <summary>Explicit review of the single project-local VBA test support edit.</summary>
    internal sealed partial class TestSupportReviewDialog : Form
    {

        /// <summary>Builds the support-source review dialog and its fixed comparison layout.</summary>
        public TestSupportReviewDialog()
        {
            InitializeComponent();
            Icon = VbeWindowIcons.Icon("assistant");
            UiText.Apply(this, components);
        }

        /// <summary>Shows the project-local support-module diff and requires an explicit OK response.</summary>
        /// <param name="owner">Window that owns the modal review dialog.</param>
        /// <param name="project">Project name displayed above the diff.</param>
        /// <param name="before">Existing module source, or empty text when no module exists.</param>
        /// <param name="after">Proposed support-module source.</param>
        /// <returns>True only when the dialog returns <see cref="DialogResult.OK"/>.</returns>
        internal static bool Confirm(IWin32Window owner, string project, string before, string after) => Confirm(owner, project, before, after, null);

        /// <summary>Displays the review using the native modal UI or an injected display callback.</summary>
        /// <param name="owner">Window that owns the modal review dialog.</param>
        /// <param name="project">Project name displayed above the diff.</param>
        /// <param name="before">Existing module source, or empty text when no module exists.</param>
        /// <param name="after">Proposed support-module source.</param>
        /// <param name="show">Optional presentation seam; receives the prepared dialog and owner.</param>
        /// <returns>True only when the selected presentation returns <see cref="DialogResult.OK"/>.</returns>
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
