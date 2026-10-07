namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System.ComponentModel;
    using System.Globalization;
    using System.Windows.Forms;
    using VBAi;

    /// <summary>Vérifie la traduction des contrôles et la détection de culture de l’interface.</summary>
    public sealed partial class UiLocalizationTests
    {
        [TestMethod]
        public void ModelCapabilityLabelsAndExplicitUnknownStatesExistInEveryCatalogue()
        {
            using (var scope = new VBAi.Tests.Infrastructure.LocalizationScope())
                foreach (var language in UiLanguages.All)
                {
                    VBAi.Tests.Infrastructure.LocalizationScope.Set(language.CultureName);
                    var catalogue = new System.Resources.ResourceManager("VBAi.Localization.UiStrings" + language.ResourceSuffix, typeof(UiText).Assembly);
                    foreach (var key in new[] { "Tool calling", "Reasoning", "Vision", "Supported", "Not supported", "Unknown", "Text only", "This model does not support tool calling." })
                    {
                        string translated = catalogue.GetString(key, CultureInfo.InvariantCulture);
                        Assert.IsFalse(string.IsNullOrWhiteSpace(translated), language.CultureName + " " + key);
                        Assert.AreEqual(translated, UiText.Get(key));
                    }
                    Assert.AreNotEqual(UiText.Get("Supported"), UiText.Get("Not supported"));
                    Assert.AreNotEqual(UiText.Get("Unknown"), UiText.Get("Not supported"));
                }
        }

        [TestMethod]
        public void ActivityCaptionsHaveTranslationsAndCancellationRemainsDistinctFromInterruption()
        {
            string[] keys = { "The LLM endpoint must use HTTPS (or HTTP on localhost), without embedded credentials or a fragment.",
                "The configured chat endpoint has no recognized model catalog path.", "The request does not belong to its model connection.",
                "Reading VBA code", "Updating VBA code", "Working in the editor", "Inspecting VBA project", "Using a tool", "Searching the web", "Viewing an image", "Running a command", "Coordinating agents", "Appearance recovery failed",
                "The host does not expose the native Compile command. Compilation was not verified.",
                "The native Compile command is disabled. The project may already be compiled or the host may restrict it; compilation was not verified." };
            using (var scope = new VBAi.Tests.Infrastructure.LocalizationScope())
                foreach (var language in UiLanguages.All)
                {
                    VBAi.Tests.Infrastructure.LocalizationScope.Set(language.CultureName);
                    var catalogue = new System.Resources.ResourceManager("VBAi.Localization.UiStrings" + language.ResourceSuffix, typeof(UiText).Assembly);
                    foreach (string key in keys)
                    {
                        string translated = catalogue.GetString(key, CultureInfo.InvariantCulture);
                        Assert.IsFalse(string.IsNullOrWhiteSpace(translated), language.CultureName);
                        Assert.AreEqual(translated, UiText.Get(key));
                        if (language.CultureName != "en-US") Assert.AreNotEqual(key, translated, language.CultureName);
                    }
                    Assert.AreNotEqual(UiText.Get("Cancelled"), UiText.Get("Interrupted"), language.CultureName);
                }
        }

        [TestMethod]
        public void ConversationDeletionHasLocalizedConfirmationAndPreservesItsTitlePlaceholder()
        {
            string[] keys = {
                "Delete conversation", "Permanently delete the selected conversation from local history.",
                "Delete conversation \"{0}\" from local history? This cannot be undone.",
                "Conversation deleted from local history", "Conversation not deleted: ",
                "This conversation changed in another host. Reopen it before deleting."
            };
            using (var scope = new VBAi.Tests.Infrastructure.LocalizationScope())
                foreach (var language in UiLanguages.All)
                {
                    VBAi.Tests.Infrastructure.LocalizationScope.Set(language.CultureName);
                    var catalogue = new System.Resources.ResourceManager("VBAi.Localization.UiStrings" + language.ResourceSuffix, typeof(UiText).Assembly);
                    foreach (string key in keys)
                    {
                        string translated = catalogue.GetString(key, CultureInfo.InvariantCulture);
                        Assert.IsFalse(string.IsNullOrWhiteSpace(translated), language.CultureName);
                        Assert.AreEqual(translated, UiText.Get(key));
                        if (language.CultureName != "en-US") Assert.AreNotEqual(key, translated, language.CultureName);
                    }
                    StringAssert.Contains(string.Format(UiText.Get(keys[2]), "Fixture title"), "Fixture title");
                    using (var window = new ChatWindow())
                    {
                        var delete = (Button)typeof(ChatWindow).GetField("deleteSession", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).GetValue(window);
                        Assert.AreEqual(UiText.Get(keys[0]), delete.Text);
                        Assert.AreEqual(UiText.Get(keys[0]), delete.AccessibleName);
                        Assert.IsTrue(delete.AutoSize, "Localized captions must retain their preferred width.");
                        Assert.AreEqual(0, delete.Parent.Controls.GetChildIndex(delete), "Deletion must be the first history action.");
                        Assert.IsFalse(delete.Enabled, "The inert Designer window has no saved session to delete.");
                    }
                }
        }

        /// <summary>Requires genuine localized recovery warnings while preserving their exact English wording in en-US.</summary>
        [TestMethod]
        public void RecoveryWarningsHaveExplicitTranslationsAndKeepEnglishDiagnosticKeys()
        {
            string[] warnings = {
                "VBA import failed and its resulting state could not be recorded. Recovery remains pending; inspect the retained backup and live project before restoring.",
                "Inspection and restoration of the native Code Colors category both failed. The original category selection is unverified."
            };
            using (var scope = new VBAi.Tests.Infrastructure.LocalizationScope())
            {
                foreach (var language in UiLanguages.All)
                {
                    VBAi.Tests.Infrastructure.LocalizationScope.Set(language.CultureName);
                    var catalogue = new System.Resources.ResourceManager("VBAi.Localization.UiStrings" + language.ResourceSuffix, typeof(UiText).Assembly);
                    foreach (string warning in warnings)
                    {
                        string embedded = catalogue.GetString(warning, CultureInfo.InvariantCulture);
                        Assert.IsFalse(string.IsNullOrWhiteSpace(embedded), language.CultureName);
                        Assert.AreEqual(embedded, UiText.Get(warning), language.CultureName);
                        if (language.CultureName == "en-US") Assert.AreEqual(warning, embedded);
                        else Assert.AreNotEqual(warning, embedded, "A populated English fallback is not a translation: " + language.CultureName);
                    }
                }
            }
        }

        /// <summary>Donne priorité aux menus du VBE sans modifier la culture du thread hôte.</summary>
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
                VBAi.Tests.Infrastructure.LocalizationScope.Set(originalUiTextCulture.Name);
            }
        }

        /// <summary>Traduit récursivement les contrôles fixes, menus, noms accessibles et infobulles.</summary>
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
                VBAi.Tests.Infrastructure.LocalizationScope.Set(originalUiTextCulture.Name);
            }
        }

        /// <summary>Applique le sens de lecture arabe tout en conservant les champs techniques de gauche à droite.</summary>
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
                VBAi.Tests.Infrastructure.LocalizationScope.Set(originalUiTextCulture.Name);
            }
        }
    }
}
namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Globalization;
    using System.Reflection;
    using System.Resources;
    using System.Windows.Forms;
    using VBAi;
    using VBAi.Tests.Infrastructure;
    /// <summary>Vérifie les chemins de repli, détection et application de la localisation.</summary>
    [TestClass]
    public sealed class UiTextCoverageTests
    {
        /// <summary>Ignore les barres qui ne sont pas des menus et traite les hôtes vides ou indisponibles.</summary>
        [TestMethod, TestCategory("Unit")]
        public void DetectionSkipsNonMenuBarsAndHandlesEmptyAndUnavailableHosts()
        {
            using (var scope = new LocalizationScope())
            {
                var host = new UiLocalizationTests.FakeVbe { CommandBars = new[] { new UiLocalizationTests.FakeBar { Type = 0, Controls = new[] { new UiLocalizationTests.FakeControl { Caption = "Outils" } } }, new UiLocalizationTests.FakeBar { Type = 1, Controls = new UiLocalizationTests.FakeControl[0] } } };
                Assert.AreEqual("en-US", UiText.Detect(host, CultureInfo.GetCultureInfo("fr-FR")).Name);
                host.CommandBars = new[] { new UiLocalizationTests.FakeBar { Type = 0 } };
                Assert.AreEqual("de-DE", UiText.Detect(host, CultureInfo.GetCultureInfo("de-DE")).Name);
                Assert.AreEqual("en-US", UiText.Detect(null, null).Name);
                UiText.Initialize(host); Assert.AreEqual(UiText.Supported(CultureInfo.CurrentUICulture).Name, UiText.Culture.Name);
            }
        }
        /// <summary>Vérifie les ressources de repli et les jetons de locuteur sans altérer le texte inconnu.</summary>
        [TestMethod, TestCategory("Unit")]
        public void ResourceFallbackAndAllSpeakerTokensPreserveUnknownText()
        {
            using (var scope = new LocalizationScope("fr-FR"))
            {
                Assert.IsNull(UiText.Get(null)); Assert.AreEqual(UiText.Get("Document conversations").ToUpperInvariant(), UiText.Get("DOCUMENT CONVERSATIONS")); Assert.AreEqual(UiText.Get("You").ToUpperInvariant(), UiText.Get("YOU"));
                var catalogue = (Dictionary<string, ResourceManager>)typeof(UiText).GetField("Catalogues", BindingFlags.Static | BindingFlags.NonPublic).GetValue(null);
                var original = catalogue[UiText.Culture.Name];
                try { catalogue[UiText.Culture.Name] = new EmptyResourceManager(); Assert.AreEqual("Save", UiText.Get("Save")); Assert.AreEqual("unknown fixture", UiText.Get("unknown fixture")); }
                finally { catalogue[UiText.Culture.Name] = original; }
                var tokens = new[] { "Vous", "Réflexion", "Vérification", "Outil", "Erreur", "Intervention", "Assistant" };
                var keys = new[] { "You", "Reasoning", "Verification", "Tool", "Error", "Turn", "Assistant" };
                for (int i = 0; i < tokens.Length; i++) Assert.AreEqual(UiText.Get(keys[i]), UiText.Speaker(tokens[i]));
                Assert.AreEqual("other", UiText.Speaker("other")); Assert.IsNull(UiText.Speaker(null));
                foreach (var language in UiLanguages.All) { LocalizationScope.Set(language.CultureName); Assert.IsFalse(string.IsNullOrEmpty(UiText.Get("Save"))); }
            }
        }
        /// <summary>Traduit menus imbriqués et infobulles dans les deux sens sans modifier les champs techniques.</summary>
        [STATestMethod, TestCategory("Unit")]
        public void ControlsTooltipsNestedMenusAndTechnicalFieldsTranslateInBothDirections()
        {
            using (var theme = new ThemeScope())
            using (var scope = new LocalizationScope("fr-FR"))
            using (var components = new Container())
            using (var form = new Form { Text = "Save" })
            using (var externalTips = new ToolTip())
            {
                var tips = new ToolTip(components); var menu = new ContextMenuStrip(components);
                var parent = new ToolStripMenuItem("Save") { ToolTipText = "Save", AccessibleName = "Save" }; parent.DropDownItems.Add(new ToolStripMenuItem("Cancel")); menu.Items.Add(parent); menu.Items.Add(new ToolStripLabel("Save"));
                components.Add(new Component());
                var button = new Button { Text = "Save", AccessibleName = "Save", AccessibleDescription = "Cancel" };
                var combo = new ComboBox(); var marker = new object(); combo.Items.AddRange(new object[] { "Save", marker });
                var grid = new DataGridView(); grid.Columns.Add("code", "Save");
                form.Controls.AddRange(new Control[] { button, combo, grid }); tips.SetToolTip(button, "Save"); externalTips.SetToolTip(button, "Cancel");
                UiText.Apply(form, components, externalTips);
                Assert.AreEqual("Enregistrer", button.Text); Assert.AreEqual("Annuler", button.AccessibleDescription);
                Assert.AreEqual("Enregistrer", tips.GetToolTip(button)); Assert.AreEqual("Annuler", externalTips.GetToolTip(button));
                Assert.AreEqual("Enregistrer", combo.Items[0]); Assert.AreSame(marker, combo.Items[1]); Assert.AreEqual("Enregistrer", grid.Columns[0].HeaderText);
                Assert.AreEqual("Enregistrer", parent.Text); Assert.AreEqual("Enregistrer", parent.ToolTipText); Assert.AreEqual("Enregistrer", parent.AccessibleName); Assert.AreEqual("Annuler", parent.DropDownItems[0].Text); Assert.AreEqual("Enregistrer", menu.Items[1].Text);
                Assert.AreEqual(RightToLeft.No, form.RightToLeft); Assert.IsFalse(form.RightToLeftLayout);
                LocalizationScope.Set("ar-SA");
                using (var rtl = new Form())
                {
                    foreach (var name in new[] { "remote", "branch", "branchName", "branchList", "baseContent", "openAiEndpoint", "ollamaEndpoint", "openAiKey", "manualModels", "resolutionText", "contextPreview", "details" }) rtl.Controls.Add(new TextBox { Name = name });
                    var ordinary = new Label { Name = "message", Text = "Save" }; rtl.Controls.Add(ordinary); rtl.Controls.Add(new ComboBox()); rtl.Controls.Add(new DataGridView()); UiText.Apply(rtl, null);
                    Assert.AreEqual(RightToLeft.Yes, rtl.RightToLeft); Assert.IsTrue(rtl.RightToLeftLayout); Assert.AreEqual(RightToLeft.Yes, ordinary.RightToLeft);
                    foreach (Control child in rtl.Controls) if (child != ordinary) Assert.AreEqual(RightToLeft.No, child.RightToLeft);
                }
                var context = LicenseManager.CurrentContext;
                try { LicenseManager.CurrentContext = new DesignContext(); using (var design = new Label { Text = "Save" }) { UiText.Apply(design, null); Assert.AreEqual("Save", design.Text); } }
                finally { LicenseManager.CurrentContext = context; }
            }
        }
    }
}
