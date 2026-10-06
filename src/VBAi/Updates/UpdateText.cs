using System.Globalization;
using System.Resources;

namespace VBAi
{

    /// <summary>Textes du programme externe, qui ne charge pas la bibliothèque COM.</summary>
    internal static class UpdateText
    {

        /// <summary>Culture name selected by the external updater process.</summary>
        internal static string Culture = CultureInfo.CurrentUICulture.Name;

        /// <summary>Looks up a localized updater resource using the current updater culture.</summary>
        /// <param name="key">Resource key.</param><returns>Localized text, or the key when no resource exists.</returns>
        internal static string Get(string key)
        {
            var language = UiLanguages.For(CultureInfo.GetCultureInfo(Culture));
            var resources = new ResourceManager("VBAi.Localization.UiStrings" + language.ResourceSuffix, typeof(UpdateText).Assembly);
            return resources.GetString(key, CultureInfo.InvariantCulture) ?? key;
        }
    }
}
