using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Infrastructure;
namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ChatActivityStepViewDesignerContractTests
    {
        [DataRow(9f)]
        [DataRow(14f)]
        [STATestMethod]
        public void StatusColumnFitsLocalizedTerminalStateAndDurationWithoutWrapping(float fontSize)
        {
            using (var theme = new ThemeScope())
            using (var culture = new LocalizationScope("fr-FR"))
            using (var form = new System.Windows.Forms.Form { Left = -10000, Top = -10000, ShowInTaskbar = false, ClientSize = new System.Drawing.Size(660, 300) })
            using (var font = new System.Drawing.Font("Segoe UI", fontSize))
            using (var view = new ChatActivityStepView { Dock = System.Windows.Forms.DockStyle.Top, Font = font })
            {
                form.Controls.Add(view);
                view.state.Text = UiText.Get("Cancelled") + " · 123,4 s";
                form.Show();
                System.Windows.Forms.Application.DoEvents();
                view.PerformLayout();
                var layout = UiInvoke.Field<System.Windows.Forms.TableLayoutPanel>(view, "layout");
                Assert.AreEqual(System.Windows.Forms.SizeType.AutoSize, layout.ColumnStyles[1].SizeType);
                Assert.IsFalse(view.state.UseCompatibleTextRendering, "The status uses the same GDI text renderer as its measurement oracle.");
                var natural = view.state.GetPreferredSize(System.Drawing.Size.Empty);
                var measured = System.Windows.Forms.TextRenderer.MeasureText(view.state.Text, view.state.Font);
                Assert.IsTrue(view.state.Width >= measured.Width, "The status must reserve the translated single-line width.");
                Assert.IsTrue(view.state.GetPreferredSize(new System.Drawing.Size(view.state.Width, 0)).Height <= measured.Height + 2, "A terminal word and duration must not split across lines.");
                Assert.IsTrue(view.state.Width >= natural.Width, "The status must retain its renderer-native unconstrained width.");
                Assert.AreEqual(natural.Height, view.state.Height, "The status occupies exactly one natural text line.");
                Assert.IsTrue(view.section.Width > 300, "The activity caption must retain usable width.");
            }
        }

        [STATestMethod]
        public void DesignerLayoutAndManagedDisposalAreComplete() { TranscriptFixture.LayoutAndLifecycle<ChatActivityStepView>(); }
    }
}