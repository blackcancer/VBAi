namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System.Globalization;
    using VBAi;

    /// <summary>Vérifie l’exhaustivité des catalogues de langue intégrés et la détection des menus.</summary>
    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class UiLocalizationTests
    {
        /// <summary>Vérifie les traductions, les noms de culture pris en charge et les libellés de menus.</summary>
        [TestMethod]
        public void AllLanguagesHaveCompleteEmbeddedCataloguesAndRecognizableMenus()
        {
            var baseline = new System.Resources.ResourceManager("VBAi.Localization.UiStrings", typeof(UiText).Assembly).GetResourceSet(CultureInfo.InvariantCulture, true, true);
            try
            {
                foreach (var language in UiLanguages.All)
                {
                    var host = new FakeVbe
                    {
                        CommandBars = new[]
                        {
                            new FakeBar
                            {
                                Type = 1,
                                Controls = new[]
                                {
                                    new FakeControl
                                    {
                                        Caption = language.View[0] + "(&V)"
                                    },
                                    new FakeControl
                                    {
                                        Caption = language.Tools[0] + "(&T)"
                                    }
                                }
                            }
                        }
                    };
                    Assert.AreEqual(language.CultureName, UiText.Detect(host, CultureInfo.GetCultureInfo("en-US")).Name);
                    Assert.IsTrue(UiLanguages.IsMenu(language.View[0] + "(&V)", true));
                    Assert.IsTrue(UiLanguages.IsMenu(language.Tools[0] + "(&T)", false));
                    UiText.Initialize(host);
                    var catalogue = new System.Resources.ResourceManager("VBAi.Localization.UiStrings" + language.ResourceSuffix, typeof(UiText).Assembly).GetResourceSet(CultureInfo.InvariantCulture, true, true);
                    foreach (System.Collections.DictionaryEntry entry in baseline)
                    {
                        var value = catalogue.GetString((string)entry.Key);
                        Assert.IsFalse(string.IsNullOrWhiteSpace(value), language.CultureName + " / " + entry.Key);
                        Assert.AreEqual(value, UiText.Get((string)entry.Key));
                    }
                }

                Assert.AreEqual("zh-TW", UiText.Supported(CultureInfo.GetCultureInfo("zh-HK")).Name);
                Assert.AreEqual("zh-TW", UiText.Supported(CultureInfo.GetCultureInfo("zh-Hant")).Name);
                Assert.AreEqual("zh-CN", UiText.Supported(CultureInfo.GetCultureInfo("zh-SG")).Name);
                Assert.AreEqual("pt-BR", UiText.Supported(CultureInfo.GetCultureInfo("pt-PT")).Name);
                Assert.AreEqual("es-ES", UiText.Supported(CultureInfo.GetCultureInfo("es-MX")).Name);
                Assert.AreEqual("en-US", UiText.Detect(Host("Tools"), CultureInfo.GetCultureInfo("de-DE")).Name);
            }
            finally
            {
                UiText.Initialize(null);
            }
        }
    }
}
