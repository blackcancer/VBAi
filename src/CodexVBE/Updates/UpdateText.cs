using System.Globalization;
using System.Resources;

namespace CodexVBE
{
    /// <summary>Textes du programme externe, qui ne charge pas la bibliothèque COM.</summary>
    internal static class UpdateText
    {
        internal static string Culture = CultureInfo.CurrentUICulture.Name;
        internal static string Get(string key)
        {
            var language = UiLanguages.For(CultureInfo.GetCultureInfo(Culture));
            var resources = new ResourceManager("CodexVBE.Localization.UiStrings" + language.ResourceSuffix, typeof(UpdateText).Assembly);
            return resources.GetString(key, CultureInfo.InvariantCulture) ?? key;
        }
    }
}
