using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace VBAi
{

    /// <summary>Review-only controls, independent of live VBA project services.</summary>
    internal sealed partial class TestSupportReviewDialog
    {

        /// <summary>Owns child components disposed when the review dialog closes.</summary>
        private IContainer components;

        /// <summary>Arranges the project label, safety explanation, source diff, and actions.</summary>
        private TableLayoutPanel layout;

        /// <summary>Names the project whose support module source is being reviewed.</summary>
        private Label projectLabel;

        /// <summary>Explains that a source backup is preserved and the project is revalidated before applying.</summary>
        private Label explanation;

        /// <summary>Displays the proposed VBA support-module source change for explicit review.</summary>
        private CodeDiffView diff;

        /// <summary>Hosts explicit Apply and Cancel decisions.</summary>
        private FlowLayoutPanel actions;

        /// <summary>Accepts the reviewed support-module change.</summary>
        private UiActionButton apply;

        /// <summary>Closes without applying the proposed source change.</summary>
        private UiActionButton cancel;

        /// <summary>Builds review-only controls without resolving or mutating a live VBA project.</summary>
        private void InitializeComponent()
        {
            components = new Container();
            layout = new TableLayoutPanel();
            layout.Name = "layout";
            projectLabel = new Label();
            projectLabel.Name = "projectLabel";
            explanation = new Label();
            explanation.Name = "explanation";
            diff = new CodeDiffView();
            diff.Name = "diff";
            actions = new FlowLayoutPanel();
            actions.Name = "actions";
            apply = new UiActionButton();
            apply.Name = "apply";
            cancel = new UiActionButton();
            cancel.Name = "cancel";
            SuspendLayout();
            layout.SuspendLayout();
            actions.SuspendLayout();
            AutoScaleDimensions = new SizeF(96F, 96F);
            AutoScaleMode = AutoScaleMode.Dpi;
            Font = new Font("Segoe UI", 9F);
            ClientSize = new Size(960, 620);
            MinimumSize = new Size(680, 420);
            Name = "TestSupportReviewDialog";
            Text = "Review VBA test support installation";
            StartPosition = FormStartPosition.CenterParent;
            ShowInTaskbar = false;
            MinimizeBox = false;
            layout.Dock = DockStyle.Fill;
            layout.Padding = new Padding(12);
            layout.ColumnCount = 1;
            layout.RowCount = 4;
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            projectLabel.AutoSize = true;
            projectLabel.Dock = DockStyle.Fill;
            projectLabel.Font = new Font("Segoe UI", 10.5F, FontStyle.Bold);
            projectLabel.Text = "Project";
            projectLabel.Padding = new Padding(0, 0, 0, 8);
            explanation.AutoSize = true;
            explanation.Dock = DockStyle.Fill;
            explanation.Text = "Review the VBAiTestSupport module below. A source backup is preserved before the edit. Your project is revalidated before applying this change.";
            explanation.Padding = new Padding(0, 0, 0, 10);
            diff.Dock = DockStyle.Fill;
            actions.AutoSize = true;
            actions.Dock = DockStyle.Fill;
            actions.FlowDirection = FlowDirection.RightToLeft;
            actions.Padding = new Padding(0, 8, 0, 0);
            cancel.Text = "Cancel";
            cancel.AutoSize = true;
            cancel.Symbol = UiSymbol.Close;
            cancel.DialogResult = DialogResult.Cancel;
            apply.Text = "Apply test support";
            apply.AutoSize = true;
            apply.Primary = true;
            apply.Symbol = UiSymbol.Check;
            apply.DialogResult = DialogResult.OK;
            actions.Controls.Add(cancel);
            actions.Controls.Add(apply);
            layout.Controls.Add(projectLabel, 0, 0);
            layout.Controls.Add(explanation, 0, 1);
            layout.Controls.Add(diff, 0, 2);
            layout.Controls.Add(actions, 0, 3);
            Controls.Add(layout);
            AcceptButton = apply;
            CancelButton = cancel;
            actions.ResumeLayout(false);
            actions.PerformLayout();
            layout.ResumeLayout(false);
            layout.PerformLayout();
            ResumeLayout(false);
        }

        /// <summary>Releases the component container when the dialog is disposed.</summary>
        /// <param name="disposing">True when managed components should also be disposed.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing) components?.Dispose();
            base.Dispose(disposing);
        }
    }
}
