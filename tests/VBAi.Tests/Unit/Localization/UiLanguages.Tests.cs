using System.Globalization;
using System.Linq;
using VBAi;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace VBAi.Tests.Unit
{
    /// <summary>Vérifie la sélection de culture et la détection des menus localisés.</summary>
    [TestClass]
    public sealed class UiLanguagesTests
    {
        /// <summary>Associe les cultures du catalogue et les alias chinois aux langues prises en charge.</summary>
        [TestMethod, TestCategory("Unit")]
        public void CultureSelectionSupportsEveryCatalogueAndAllTraditionalChineseAliases()
        {
            foreach(var language in UiLanguages.All)
            {
                Assert.AreSame(language,UiLanguages.For(CultureInfo.GetCultureInfo(language.CultureName)));
                Assert.IsTrue(language.View.Length>0); Assert.IsTrue(language.Tools.Length>0);
            }
            foreach(var name in new[]{"zh-Hant","zh-TW","zh-HK","zh-MO"}) Assert.AreEqual("zh-TW",UiLanguages.For(CultureInfo.GetCultureInfo(name)).CultureName,name);
            foreach(var name in new[]{"zh-Hans","zh-CN","zh-SG"}) Assert.AreEqual("zh-CN",UiLanguages.For(CultureInfo.GetCultureInfo(name)).CultureName,name);
            Assert.AreEqual("en-US",UiLanguages.For(null).CultureName);
            Assert.AreEqual("en-US",UiLanguages.For(CultureInfo.InvariantCulture).CultureName);
            Assert.AreEqual("en-US",UiLanguages.For(CultureInfo.GetCultureInfo("fi-FI")).CultureName);
            Assert.AreEqual("fr-FR",UiLanguages.For(CultureInfo.GetCultureInfo("fr-CA")).CultureName);
        }
        /// <summary>Normalise les légendes de menu et détecte la langue à partir des menus de chaque catalogue.</summary>
        [TestMethod, TestCategory("Unit")]
        public void MenuDetectionNormalizesMnemonicsDirectionMarksAndEveryCatalogue()
        {
            Assert.AreEqual("",UiLanguages.NormalizeMenu(null));
            Assert.AreEqual("view",UiLanguages.NormalizeMenu(" \u200e&View\u200f "));
            Assert.AreEqual("表示",UiLanguages.NormalizeMenu("表示(&V)"));
            Assert.AreEqual("檢視",UiLanguages.NormalizeMenu("檢視（V）"));
            foreach(var language in UiLanguages.All)
            {
                foreach(var name in language.View) {Assert.IsTrue(UiLanguages.IsMenu(name,true)); Assert.IsFalse(UiLanguages.IsMenu(name,false));}
                foreach(var name in language.Tools) {Assert.IsTrue(UiLanguages.IsMenu(name,false)); Assert.IsFalse(UiLanguages.IsMenu(name,true));}
                Assert.AreEqual(language.CultureName,UiLanguages.FromMenus(language.View.Concat(language.Tools),CultureInfo.GetCultureInfo(language.CultureName)).Name);
            }
            Assert.IsFalse(UiLanguages.IsMenu("unknown",true)); Assert.IsFalse(UiLanguages.IsMenu(null,false));
            Assert.AreEqual("en-US",UiLanguages.FromMenus(new string[0],CultureInfo.GetCultureInfo("fr-FR")).Name);
            Assert.AreEqual("zh-TW",UiLanguages.FromMenus(new[]{"工具"},CultureInfo.GetCultureInfo("zh-HK")).Name);
            Assert.AreEqual("zh-CN",UiLanguages.FromMenus(new[]{"工具"},CultureInfo.GetCultureInfo("fr-FR")).Name);
            Assert.AreEqual("en-US",UiLanguages.FromMenus(new[]{"view","outils"},CultureInfo.GetCultureInfo("fr-FR")).Name);
        }
    }
}
