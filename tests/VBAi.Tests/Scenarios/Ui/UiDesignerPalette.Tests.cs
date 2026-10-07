using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.ComponentModel.Design;
using System.Drawing;
using System.Windows.Forms;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Unit
{
    [TestClass]
    public sealed class UiDesignerPaletteTests
    {
        [STATestMethod]
        public void UnsitedTabsFollowTheirActualParentBeforeTheDesignerHasAssignedSites()
        {
            using (var theme = new ThemeScope())
            using (var parent = new Panel { BackColor = Color.White, ForeColor = Color.Black })
            using (var tabs = new ThemedTabControl())
            {
                ThemeScope.SetChoice(ThemeChoice.Dark);
                parent.Controls.Add(tabs);
                Assert.IsNull(tabs.Site);
                Assert.AreEqual(Color.White, UiTheme.BackgroundFor(tabs));
                Assert.AreEqual(Color.Black, UiTheme.ForegroundFor(tabs));
                parent.BackColor = Color.FromArgb(22, 26, 33);
                parent.ForeColor = Color.White;
                Assert.AreEqual(parent.BackColor, UiTheme.BackgroundFor(tabs));
                Assert.AreEqual(Color.White, UiTheme.ForegroundFor(tabs));
            }
        }

        [STATestMethod]
        public void DarkRuntimePreferenceDoesNotDarkenLightDesignerTabsOrFields()
        {
            using (var theme = new ThemeScope())
            {
                ThemeScope.SetChoice(ThemeChoice.Dark);
                foreach (var type in new[] { typeof(LlmSettingsWindow), typeof(GitWindow), typeof(ModernEditorWindow) })
                    using (var surface = new DesignSurface(type))
                    {
                        var host = (IDesignerHost)surface.GetService(typeof(IDesignerHost));
                        var form = (Form)host.RootComponent;
                        form.BackColor = Color.WhiteSmoke;
                        form.ForeColor = Color.Black;
                        Assert.IsTrue(UiTheme.IsDesignPreview(form));
                        UiTheme.Apply(form);
                        Assert.AreEqual(Color.WhiteSmoke, form.BackColor);
                        var tabs = UiInvoke.Field<ThemedTabControl>(form, type == typeof(LlmSettingsWindow) ? "settingsTabs" : "tabs");
                        Assert.IsTrue(UiTheme.IsDesignPreview(tabs));
                        Assert.IsTrue(UiTheme.BackgroundFor(tabs).GetBrightness() > .5f);
                        using (var bitmap = new Bitmap(500, 60))
                        using (var graphics = Graphics.FromImage(bitmap))
                        {
                            UiInvoke.Call(typeof(ThemedTabControl), "OnPaint", tabs, new PaintEventArgs(graphics, new Rectangle(0, 0, 500, 60)));
                            Assert.IsTrue(bitmap.GetPixel(499, 59).GetBrightness() > .5f);
                        }
                        Assert.IsTrue(UiTheme.Dark, "The Designer must not change the user's runtime preference.");
                    }
            }
        }
    }
}
