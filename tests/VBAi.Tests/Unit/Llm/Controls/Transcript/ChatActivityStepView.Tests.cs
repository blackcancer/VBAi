using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Drawing;
using System.Windows.Forms;
using VBAi.Tests.Infrastructure;
namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ChatActivityStepViewContractTests
    {
        [STATestMethod]
        public void ConstructorCreatesNamedThemedControls() { using (var theme = new ThemeScope()) using (var view = new ChatActivityStepView()) { Assert.AreEqual("ChatActivityStepView", view.Name); Assert.IsTrue(view.Controls.Count > 0); Assert.AreEqual(UiTheme.Background, view.BackColor); Assert.AreEqual(UiTheme.Foreground, view.ForeColor); } }

        /// <summary>Parent reflows must not recursively measure an unchanged completed activity on every streamed frame.</summary>
        [STATestMethod]
        public void StableConstrainedRowMeasurementsDoNotRemeasureDescendants()
        {
            using (var row = new ChatActivityStepView())
            using (var probe = new MeasuredContent())
            {
                row.section.Expanded = true;
                row.section.body.Controls.Add(probe);
                var narrow = new Size(340, 0);
                var wide = new Size(780, 0);
                var narrowSize = row.GetPreferredSize(narrow);
                var wideSize = row.GetPreferredSize(wide);
                probe.Reads = 0;
                for (int frame = 0; frame < 50; frame++)
                {
                    Assert.AreEqual(narrowSize, row.GetPreferredSize(narrow));
                    Assert.AreEqual(wideSize, row.GetPreferredSize(wide));
                }
                Assert.AreEqual(0, probe.Reads, "An unchanged sibling must retain its width-specific measurement.");
            }
        }

        /// <summary>Cached rows must continue to honor all changing content, disclosure, font and geometry inputs.</summary>
        [DataRow("text")]
        [DataRow("font")]
        [DataRow("padding")]
        [DataRow("margin")]
        [DataRow("visible")]
        [DataRow("size")]
        [DataRow("minimum")]
        [DataRow("maximum")]
        [DataRow("disclosure")]
        [DataRow("title")]
        [DataRow("state")]
        [DataRow("width")]
        [DataRow("add")]
        [DataRow("remove")]
        [STATestMethod]
        public void ChangedRowInputsMatchAFreshMeasurement(string change)
        {
            using (var row = new ChatActivityStepView())
            using (var probe = new MeasuredContent())
            using (var font = new Font("Segoe UI", 16))
            using (var added = new Label { AutoSize = true, Text = "Added descendant" })
            using (var reference = new ChatActivityStepView())
            using (var referenceProbe = new MeasuredContent())
            using (var referenceAdded = new Label { AutoSize = true, Text = "Added descendant" })
            {
                row.section.Expanded = true;
                row.section.body.Controls.Add(probe);
                reference.section.Expanded = true;
                reference.section.body.Controls.Add(referenceProbe);
                var constraint = new Size(340, 0);
                row.GetPreferredSize(constraint);
                row.SuspendLayout(); row.section.body.SuspendLayout();
                reference.SuspendLayout(); reference.section.body.SuspendLayout();
                try
                {
                    ApplyMeasurementChange(change, row, probe, font, added);
                    ApplyMeasurementChange(change, reference, referenceProbe, font, referenceAdded);
                    if (change == "width") constraint = new Size(780, 0);
                    var observed = row.GetPreferredSize(constraint);
                    // The uncached reference has the same pending native layouts.
                    // WinForms can legitimately change preferred sizes on ResumeLayout.
                    Assert.AreEqual(reference.GetPreferredSize(constraint), observed, "pending " + change);
                    row.section.body.ResumeLayout(true); row.ResumeLayout(true);
                    reference.section.body.ResumeLayout(true); reference.ResumeLayout(true);
                    Assert.AreEqual(reference.GetPreferredSize(constraint), row.GetPreferredSize(constraint), "settled " + change);
                }
                finally
                {
                    row.section.body.ResumeLayout(true); row.ResumeLayout(true);
                    reference.section.body.ResumeLayout(true); reference.ResumeLayout(true);
                }
            }
        }

        private static void ApplyMeasurementChange(string change, ChatActivityStepView row, MeasuredContent probe, Font font, Label added)
        {
            switch (change)
            {
                case "text": probe.Text = new string('x', 120); break;
                case "font": row.Font = font; break;
                case "padding": probe.Padding = new Padding(20); break;
                case "margin": probe.Margin = new Padding(18); break;
                case "visible": probe.Visible = false; break;
                case "size": probe.AutoSize = false; probe.Size = new Size(120, 180); break;
                case "minimum": probe.MinimumSize = new Size(0, 180); break;
                case "maximum": probe.MaximumSize = new Size(70, 16); break;
                case "disclosure": row.section.Expanded = false; break;
                case "title": row.section.Title = "Updated action title"; break;
                case "state": row.state.Text = "Failed · 1.0 s"; break;
                case "add": row.section.body.Controls.Add(added); break;
                case "remove": row.section.body.Controls.Remove(probe); break;
            }
        }

        /// <summary>A removed descendant must no longer invalidate or retain the activity that formerly owned it.</summary>
        [STATestMethod]
        public void RemovedDescendantsCannotInvalidateStableRowMeasurements()
        {
            using (var row = new ChatActivityStepView())
            using (var remaining = new MeasuredContent())
            using (var removed = new Label { AutoSize = true, Text = "Removed" })
            {
                row.section.Expanded = true;
                row.section.body.Controls.Add(remaining);
                row.section.body.Controls.Add(removed);
                row.section.body.Controls.Remove(removed);
                var constraint = new Size(340, 0);
                row.GetPreferredSize(constraint); remaining.Reads = 0;
                removed.Text = "No longer belongs to the activity";
                row.GetPreferredSize(constraint);
                Assert.AreEqual(0, remaining.Reads);
            }
        }

        /// <summary>Exposes actual descendant measurement work and a height that responds to content and font changes.</summary>
        private sealed class MeasuredContent : Control
        {
            internal int Reads;
            internal MeasuredContent() { AutoSize = true; Text = "Initial"; }
            public override Size GetPreferredSize(Size proposedSize)
            {
                Reads++;
                return new Size(80 + Padding.Horizontal, 20 + Text.Length + (int)Font.Size + Padding.Vertical);
            }
        }
    }
}
