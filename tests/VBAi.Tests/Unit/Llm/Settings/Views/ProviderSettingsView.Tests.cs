namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System.Drawing;
    using System.Windows.Forms;
    using VBAi;

    [TestClass, TestCategory("Unit")]
    public sealed class ProviderSettingsViewTests
    {
        /// <summary>Uses native tab navigation through the real detached view without settings/authentication services.</summary>
        [WinFormsTestMethod]
        public void TabNavigationScrollsOffscreenProviderFieldsIntoViewInBothDirections()
        {
            using (var form = new Form { ClientSize = new Size(720, 230), ShowInTaskbar = false })
            using (var view = new ProviderSettingsView { Dock = DockStyle.Fill })
            {
                form.Controls.Add(view);
                form.Show(); Application.DoEvents();
                Assert.IsTrue(view.AutoScroll);
                Assert.IsTrue(view.VerticalScroll.Visible, "The compact viewport must exercise offscreen controls.");
                Assert.IsTrue(view.provider.Focus());
                Application.DoEvents();
                Control current = view.provider;
                var forward = new Control[] { view.codexLogin, view.codexRefresh, view.openAiEndpoint,
                    view.ollamaEndpoint, view.openAiKey, view.clearKey, view.approvalPicker,
                    view.manualModels, view.customName, view.azureEntra, view.ollamaTemperature, view.ollamaTopP };
                foreach (var expected in forward)
                {
                    Assert.IsTrue(form.SelectNextControl(current, true, true, true, false));
                    Application.DoEvents();
                    Assert.IsTrue(expected.Focused, "Unexpected forward tab target: " + expected.Name);
                    AssertFieldVisible(view, expected);
                    current = expected;
                }
                Assert.IsTrue(view.AutoScrollPosition.Y < 0, "Focusing the last field must move the scroll viewport.");
                for (int i = forward.Length - 2; i >= -1; i--)
                {
                    var expected = i < 0 ? (Control)view.provider : forward[i];
                    Assert.IsTrue(form.SelectNextControl(current, false, true, true, false));
                    Application.DoEvents();
                    Assert.IsTrue(expected.Focused, "Unexpected backward tab target: " + expected.Name);
                    AssertFieldVisible(view, expected);
                    current = expected;
                }
                Assert.IsTrue(view.provider.Focused);
            }
        }

        private static void AssertFieldVisible(ProviderSettingsView view, Control field)
        {
            var bounds = view.RectangleToClient(field.RectangleToScreen(field.ClientRectangle));
            Assert.IsTrue(view.ClientRectangle.Contains(bounds),
                "Focused field must be fully visible: " + field.Name + "; field=" + bounds + "; viewport=" + view.ClientRectangle);
        }
    }
}
