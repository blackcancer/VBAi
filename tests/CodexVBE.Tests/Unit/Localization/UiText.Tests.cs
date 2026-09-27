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
            var originalUiTextCulture = UiText.Culture;
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
                CodexVBE.Tests.Infrastructure.LocalizationScope.Set(originalUiTextCulture.Name);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void FixedControlsMenusAndTooltipsTranslateRecursively()
        {
            var originalUiTextCulture = UiText.Culture;
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
                CodexVBE.Tests.Infrastructure.LocalizationScope.Set(originalUiTextCulture.Name);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void ArabicMirrorsTheFormButPreservesTechnicalFields()
        {
            var originalUiTextCulture = UiText.Culture;
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
                CodexVBE.Tests.Infrastructure.LocalizationScope.Set(originalUiTextCulture.Name);
            }
        }
    }
}
namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Globalization;
    using System.Reflection;
    using System.Resources;
    using System.Windows.Forms;
    using CodexVBE;
    using CodexVBE.Tests.Infrastructure;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    [TestClass]
    public sealed class UiTextCoverageTests
    {
        [TestMethod, TestCategory("Unit")]
        public void DetectionSkipsNonMenuBarsAndHandlesEmptyAndUnavailableHosts()
        {
            using(var scope=new LocalizationScope())
            {
                var host=new UiLocalizationTests.FakeVbe {CommandBars=new[]{new UiLocalizationTests.FakeBar {Type=0,Controls=new[]{new UiLocalizationTests.FakeControl {Caption="Outils"}}},new UiLocalizationTests.FakeBar {Type=1,Controls=new UiLocalizationTests.FakeControl[0]}}};
                Assert.AreEqual("en-US",UiText.Detect(host,CultureInfo.GetCultureInfo("fr-FR")).Name);
                host.CommandBars=new[]{new UiLocalizationTests.FakeBar {Type=0}};
                Assert.AreEqual("de-DE",UiText.Detect(host,CultureInfo.GetCultureInfo("de-DE")).Name);
                Assert.AreEqual("en-US",UiText.Detect(null,null).Name);
                UiText.Initialize(host); Assert.AreEqual(UiText.Supported(CultureInfo.CurrentUICulture).Name,UiText.Culture.Name);
            }
        }
        [TestMethod, TestCategory("Unit")]
        public void ResourceFallbackAndAllSpeakerTokensPreserveUnknownText()
        {
            using(var scope=new LocalizationScope("fr-FR"))
            {
                Assert.IsNull(UiText.Get(null)); Assert.AreEqual(UiText.Get("Document conversations").ToUpperInvariant(),UiText.Get("DOCUMENT CONVERSATIONS")); Assert.AreEqual(UiText.Get("You").ToUpperInvariant(),UiText.Get("YOU"));
                var catalogue=(Dictionary<string,ResourceManager>)typeof(UiText).GetField("Catalogues",BindingFlags.Static|BindingFlags.NonPublic).GetValue(null);
                var original=catalogue[UiText.Culture.Name];
                try { catalogue[UiText.Culture.Name]=new EmptyResourceManager(); Assert.AreEqual("Save",UiText.Get("Save")); Assert.AreEqual("unknown fixture",UiText.Get("unknown fixture")); }
                finally {catalogue[UiText.Culture.Name]=original;}
                var tokens=new[]{"Vous","Réflexion","Vérification","Outil","Erreur","Intervention","Assistant"};
                var keys=new[]{"You","Reasoning","Verification","Tool","Error","Turn","Assistant"};
                for(int i=0;i<tokens.Length;i++) Assert.AreEqual(UiText.Get(keys[i]),UiText.Speaker(tokens[i]));
                Assert.AreEqual("other",UiText.Speaker("other")); Assert.IsNull(UiText.Speaker(null));
                foreach(var language in UiLanguages.All) {LocalizationScope.Set(language.CultureName); Assert.IsFalse(string.IsNullOrEmpty(UiText.Get("Save")));}
            }
        }
        [STATestMethod, TestCategory("Unit")]
        public void ControlsTooltipsNestedMenusAndTechnicalFieldsTranslateInBothDirections()
        {
            using(var theme=new ThemeScope())
            using(var scope=new LocalizationScope("fr-FR"))
            using(var components=new Container())
            using(var form=new Form {Text="Save"})
            using(var externalTips=new ToolTip())
            {
                var tips=new ToolTip(components); var menu=new ContextMenuStrip(components);
                var parent=new ToolStripMenuItem("Save") {ToolTipText="Save",AccessibleName="Save"}; parent.DropDownItems.Add(new ToolStripMenuItem("Cancel")); menu.Items.Add(parent); menu.Items.Add(new ToolStripLabel("Save"));
                components.Add(new Component());
                var button=new Button {Text="Save",AccessibleName="Save",AccessibleDescription="Cancel"};
                var combo=new ComboBox(); var marker=new object(); combo.Items.AddRange(new object[]{"Save",marker});
                var grid=new DataGridView(); grid.Columns.Add("code","Save");
                form.Controls.AddRange(new Control[]{button,combo,grid}); tips.SetToolTip(button,"Save"); externalTips.SetToolTip(button,"Cancel");
                UiText.Apply(form,components,externalTips);
                Assert.AreEqual("Enregistrer",button.Text); Assert.AreEqual("Annuler",button.AccessibleDescription);
                Assert.AreEqual("Enregistrer",tips.GetToolTip(button)); Assert.AreEqual("Annuler",externalTips.GetToolTip(button));
                Assert.AreEqual("Enregistrer",combo.Items[0]); Assert.AreSame(marker,combo.Items[1]); Assert.AreEqual("Enregistrer",grid.Columns[0].HeaderText);
                Assert.AreEqual("Enregistrer",parent.Text); Assert.AreEqual("Enregistrer",parent.ToolTipText); Assert.AreEqual("Enregistrer",parent.AccessibleName); Assert.AreEqual("Annuler",parent.DropDownItems[0].Text); Assert.AreEqual("Enregistrer",menu.Items[1].Text);
                Assert.AreEqual(RightToLeft.No,form.RightToLeft); Assert.IsFalse(form.RightToLeftLayout);
                LocalizationScope.Set("ar-SA");
                using(var rtl=new Form())
                {
                    foreach(var name in new[]{"remote","branch","branchName","branchList","baseContent","openAiEndpoint","ollamaEndpoint","openAiKey","manualModels","resolutionText","contextPreview","details"}) rtl.Controls.Add(new TextBox {Name=name});
                    var ordinary=new Label {Name="message",Text="Save"}; rtl.Controls.Add(ordinary); rtl.Controls.Add(new ComboBox()); rtl.Controls.Add(new DataGridView()); UiText.Apply(rtl,null);
                    Assert.AreEqual(RightToLeft.Yes,rtl.RightToLeft); Assert.IsTrue(rtl.RightToLeftLayout); Assert.AreEqual(RightToLeft.Yes,ordinary.RightToLeft);
                    foreach(Control child in rtl.Controls) if(child!=ordinary) Assert.AreEqual(RightToLeft.No,child.RightToLeft);
                }
                var context=LicenseManager.CurrentContext;
                try {LicenseManager.CurrentContext=new DesignContext(); using(var design=new Label {Text="Save"}) {UiText.Apply(design,null); Assert.AreEqual("Save",design.Text);}}
                finally {LicenseManager.CurrentContext=context;}
            }
        }
    }
}
