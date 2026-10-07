using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Reflection;
using System.Windows.Forms;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class SettingsExtendedLayoutTests
    {
        [STATestMethod]
        public void ProviderConfigurationKeepsAdditionalRowsVisible()
        {
            using (var window = new LlmSettingsWindow(new LlmSettings { ProviderName = "OpenAI API" }))
            {
                var grid = (TableLayoutPanel)typeof(LlmSettingsWindow).GetField("grid", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(window);
                using (var label = new Label { Text = "Extension", Visible = false })
                {
                    grid.Controls.Add(label, 0, grid.RowCount);
                    LlmBoundaryScope.Call(window, "UpdateRows");
                    window.Show(); Assert.IsTrue(label.Visible);
                }
            }
        }
    }
}
