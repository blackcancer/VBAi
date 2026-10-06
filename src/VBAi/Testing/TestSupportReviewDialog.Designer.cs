using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace VBAi
{

    /// <summary>Review-only controls, independent of live VBA project services.</summary>
    internal sealed partial class TestSupportReviewDialog
    {

        /// <summary>Maintains the components state for test support review dialog.</summary>
        private IContainer components;

        /// <summary>Maintains the layout state for test support review dialog.</summary>
        private TableLayoutPanel layout;

        /// <summary>Maintains the project label state for test support review dialog.</summary>
        private Label projectLabel;

        /// <summary>Maintains the explanation state for test support review dialog.</summary>
        private Label explanation;

        /// <summary>Maintains the diff state for test support review dialog.</summary>
        private CodeDiffView diff;

        /// <summary>Maintains the actions state for test support review dialog.</summary>
        private FlowLayoutPanel actions;

        /// <summary>Maintains the apply state for test support review dialog.</summary>
        private UiActionButton apply;

        /// <summary>Tracks the cancel state of test support review dialog.</summary>
        private UiActionButton cancel;

        /// <summary>Handles initialize component for test support review dialog.</summary>
        private void InitializeComponent()
        {
            components = new Container();
            layout = new TableLayoutPanel();
            projectLabel = new Label();
            explanation = new Label();
            diff = new CodeDiffView();
            actions = new FlowLayoutPanel();
            apply = new UiActionButton();
            cancel = new UiActionButton();
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

        /// <summary>Disposes  for test support review dialog.</summary>
        /// <param name="disposing">Indicates whether disposing is enabled.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing) components?.Dispose();
            base.Dispose(disposing);
        }
    }
}
