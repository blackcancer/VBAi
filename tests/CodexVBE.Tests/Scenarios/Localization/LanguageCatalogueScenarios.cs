namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Globalization;
    using System.Windows.Forms;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class UiLocalizationTests
    {
        [TestMethod]
        public void AllLanguagesHaveCompleteEmbeddedCataloguesAndRecognizableMenus()
        {
            var baseline = new System.Resources.ResourceManager("CodexVBE.Localization.UiStrings", typeof(UiText).Assembly).GetResourceSet(CultureInfo.InvariantCulture, true, true);
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
                    var catalogue = new System.Resources.ResourceManager("CodexVBE.Localization.UiStrings" + language.ResourceSuffix, typeof(UiText).Assembly).GetResourceSet(CultureInfo.InvariantCulture, true, true);
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
