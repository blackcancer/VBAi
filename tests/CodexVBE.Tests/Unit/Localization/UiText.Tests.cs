namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Globalization;
    using System.Windows.Forms;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class UiLocalizationTests
    {
        [TestMethod]
        public void VbeMenuLanguageOverridesWindowsWithoutChangingTheHostCulture()
        {
            var previous = System.Threading.Thread.CurrentThread.CurrentCulture;
            var previousUi = System.Threading.Thread.CurrentThread.CurrentUICulture;
            try
            {
                Assert.AreEqual("fr-FR", UiText.Detect(Host("&Outils"), CultureInfo.GetCultureInfo("en-US")).Name);
                Assert.AreEqual("en-US", UiText.Detect(Host("&View"), CultureInfo.GetCultureInfo("fr-CA")).Name);
                Assert.AreEqual("de-DE", UiText.Detect(Host("Ansicht"), CultureInfo.GetCultureInfo("fr-CA")).Name);
                Assert.AreEqual("fr-FR", UiText.Detect(null, CultureInfo.GetCultureInfo("fr-CA")).Name);
                Assert.AreEqual("en-US", UiText.Supported(CultureInfo.GetCultureInfo("fi-FI")).Name);
                UiText.Initialize(Host("&Outils"));
                Assert.AreEqual("fr-FR", UiText.Culture.Name);
                Assert.AreEqual(previous, System.Threading.Thread.CurrentThread.CurrentCulture);
                Assert.AreEqual(previousUi, System.Threading.Thread.CurrentThread.CurrentUICulture);
                Assert.AreEqual("Enregistrer", UiText.Get("Save"));
                Assert.AreEqual("VOUS", UiText.Get("YOU"));
                Assert.AreEqual("RÉFLEXION", UiText.Speaker("Réflexion").ToUpperInvariant());
                Assert.IsNull(UiText.Get(null));
                Assert.AreEqual("Untranslated fixture", UiText.Get("Untranslated fixture"));
            }
            finally
            {
                UiText.Initialize(null);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void FixedControlsMenusAndTooltipsTranslateRecursively()
        {
            try
            {
                UiText.Initialize(Host("Outils"));
                using (var components = new Container())
                using (var form = new Form
                {
                    Text = "VBAi — approve edit"
                }

                )
                using (var tips = new ToolTip())
                {
                    var button = new Button
                    {
                        Text = "Allow",
                        AccessibleName = "Allow"
                    };
                    var combo = new ComboBox();
                    combo.Items.Add("Automatic");
                    form.Controls.Add(button);
                    form.Controls.Add(combo);
                    tips.SetToolTip(button, "Deny");
                    var menu = new ContextMenuStrip(components);
                    menu.Items.Add(new ToolStripMenuItem("Settings…"));
                    UiText.Apply(form, components, tips);
                    Assert.AreEqual(UiText.Get("VBAi — approve edit"), form.Text);
                    Assert.AreEqual("Autoriser", button.Text);
                    Assert.AreEqual("Autoriser", button.AccessibleName);
                    Assert.AreEqual("Refuser", tips.GetToolTip(button));
                    Assert.AreEqual(UiText.Get("Automatic"), combo.Items[0]);
                    Assert.AreEqual("Paramètres…", menu.Items[0].Text);
                }
            }
            finally
            {
                UiText.Initialize(null);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void ArabicMirrorsTheFormButPreservesTechnicalFields()
        {
            try
            {
                UiText.Initialize(Host("عرض"));
                using (var form = new Form())
                {
                    var url = new TextBox
                    {
                        Name = "remote",
                        Text = "https://github.com/org/repo.git"
                    };
                    var code = new TextBox
                    {
                        Name = "resolutionText",
                        Text = "Sub Test()"
                    };
                    var message = new TextBox
                    {
                        Name = "commitMessage"
                    };
                    form.Controls.AddRange(new Control[] { url, code, message });
                    UiText.Apply(form, null);
                    Assert.IsTrue(form.RightToLeftLayout);
                    Assert.AreEqual(RightToLeft.Yes, form.RightToLeft);
                    Assert.AreEqual(RightToLeft.No, url.RightToLeft);
                    Assert.AreEqual(RightToLeft.No, code.RightToLeft);
                    Assert.AreEqual(RightToLeft.Yes, message.RightToLeft);
                    Assert.AreEqual("https://github.com/org/repo.git", url.Text);
                    Assert.AreEqual("Sub Test()", code.Text);
                }
            }
            finally
            {
                UiText.Initialize(null);
            }
        }
    }
}
