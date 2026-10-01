using System;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class TestExplorerWindowDesignerTests
    {
        [STATestMethod]
        public void DesignerUsesGlobalCommandsAndInputsWithoutLiveServices()
        {
            using (var window = new TestExplorerWindow())
            {
                Assert.IsNull(Field<object>(window, "service"));
                Assert.IsInstanceOfType(Field<object>(window, "refresh"), typeof(UiActionButton));
                Assert.IsInstanceOfType(Field<object>(window, "search"), typeof(UiTextBox));
                Assert.IsInstanceOfType(Field<object>(window, "projectList"), typeof(UiComboBox));
                Assert.IsInstanceOfType(Field<object>(window, "resultTabs"), typeof(ThemedTabControl));
                Assert.AreEqual(AutoScaleMode.Dpi, window.AutoScaleMode);
                Assert.AreEqual(2000, Field<System.Windows.Forms.Timer>(window, "freshnessTimer").Interval);
                Assert.IsFalse(Field<System.Windows.Forms.Timer>(window, "freshnessTimer").Enabled);
                Assert.IsFalse(Field<Button>(window, "runScope").Enabled);
            }
        }

        [STATestMethod]
        public void EnglishAndFrenchControlsUseSeparateLocalizedCaptionsAndAccessibility()
        {
            var culture = UiText.Culture;
            var property = typeof(UiText).GetProperty("Culture", BindingFlags.Static | BindingFlags.NonPublic);
            try
            {
                foreach (var language in new[] { "en-US", "fr-FR" })
                {
                    property.SetValue(null, CultureInfo.GetCultureInfo(language));
                    using (var window = new TestExplorerWindow())
                    {
                        Assert.AreEqual(language == "fr-FR" ? "Actualiser les tests" : "Refresh tests", Field<Button>(window, "refresh").Text);
                        Assert.AreEqual(language == "fr-FR" ? "Exécuter les tests sélectionnés" : "Run selected tests", Field<Button>(window, "runSelected").Text);
                        Assert.AreEqual(language == "fr-FR" ? "Exécuter la sélection avec couverture…" : "Run selected with coverage...", Field<Button>(window, "runSelectedCoverage").Text);
                        Assert.AreEqual(language == "fr-FR" ? "Hiérarchie des tests" : "Test hierarchy", Field<TreeView>(window, "testTree").AccessibleName);
                        Assert.AreEqual(language == "fr-FR" ? "Explorateur de tests VBAi…" : "VBAi test explorer…", UiText.Get("VBAi test explorer…"));
                        Assert.AreEqual(UiTheme.Foreground, Field<Button>(window, "runSelected").ForeColor);
                        Assert.IsNotNull(window.Icon);
                    }
                }
            }
            finally { property.SetValue(null, culture); }
        }

        [STATestMethod]
        public void HumanReportUsesFrenchLabelsAndNativeDecimalFormat()
        {
            var language = UiText.Culture;
            var numeric = CultureInfo.CurrentCulture;
            var property = typeof(UiText).GetProperty("Culture", BindingFlags.Static | BindingFlags.NonPublic);
            try
            {
                property.SetValue(null, CultureInfo.GetCultureInfo("fr-FR"));
                CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("fr-FR");
                var run = new VbaTestRun { Id = "run", Revision = "revision" };
                run.Results.Add(new VbaTestResult { Test = new VbaTestDescriptor { Id = "test", Module = "Tests", Procedure = "Adds" },
                    Outcome = VbaTestOutcome.Passed, Duration = TimeSpan.FromTicks(125000) });
                string report = VbaTestReports.Human(run);
                StringAssert.Contains(report, "Résultats des tests VBA");
                StringAssert.Contains(report, "Taux de réussite : 100%");
                StringAssert.Contains(report, "12,5 ms");
                StringAssert.Contains(VbaTestReports.Compact(run), "\"ms\":12.5");
            }
            finally { property.SetValue(null, language); CultureInfo.CurrentCulture = numeric; }
        }

        [STATestMethod]
        public void GlobalPaletteReappliesForBothLightAndDarkChoices()
        {
            var choice = UiTheme.Choice;
            var contrast = UiTheme.HighContrast;
            var property = typeof(UiTheme).GetProperty("Choice", BindingFlags.Static | BindingFlags.NonPublic);
            try
            {
                UiTheme.HighContrast = () => false;
                using (var window = new TestExplorerWindow())
                    foreach (var theme in new[] { ThemeChoice.Light, ThemeChoice.Dark })
                    {
                        property.SetValue(null, theme);
                        UiTheme.Apply(window);
                        Assert.AreEqual(UiTheme.Background, window.BackColor);
                        Assert.AreEqual(UiTheme.Foreground, window.ForeColor);
                        Assert.AreEqual(UiTheme.Surface, Field<TextBox>(window, "search").BackColor);
                        Assert.AreEqual(UiTheme.Foreground, Field<TreeView>(window, "testTree").ForeColor);
                        Assert.AreEqual(UiTheme.Foreground, Field<Button>(window, "runSelected").ForeColor);
                    }
            }
            finally { property.SetValue(null, choice); UiTheme.HighContrast = contrast; }
        }

        [STATestMethod]
        public void DesignerDisposalSupportsFinalizerStyleAndAlreadyAbsentComponents()
        {
            var window = new TestExplorerWindow();
            typeof(TestExplorerWindow).GetMethod("Dispose", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(window, new object[] { false });
            window.Dispose();
            var noComponents = new TestExplorerWindow();
            var field = typeof(TestExplorerWindow).GetField("components", BindingFlags.Instance | BindingFlags.NonPublic);
            ((System.ComponentModel.IContainer)field.GetValue(noComponents)).Dispose();
            field.SetValue(noComponents, null);
            noComponents.Dispose();
            Assert.IsTrue(noComponents.IsDisposed);
        }
        private static T Field<T>(object target, string name) => (T)target.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(target);
    }
}
