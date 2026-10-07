using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace VBAi
{

    /// <summary>Décrit les noms de menus et le suffixe de ressources d’une langue prise en charge.</summary>
    internal sealed class UiLanguage
    {

        /// <summary>Nom de culture .NET et suffixe du catalogue de ressources.</summary>
        internal readonly string CultureName, ResourceSuffix;

        /// <summary>Variantes reconnues des menus Affichage et Outils.</summary>
        internal readonly string[] View, Tools;

        /// <summary>Initialise une langue et découpe ses variantes de menu séparées par une barre verticale.</summary>
        /// <param name="culture">Nom de culture .NET.</param>
        /// <param name="resource">Suffixe de ressources, ou chaîne vide pour la langue anglaise.</param>
        /// <param name="view">Libellés du menu Affichage séparés par « | ».</param>
        /// <param name="tools">Libellés du menu Outils séparés par « | ».</param>
        internal UiLanguage(string culture, string resource, string view, string tools)
        {
            CultureName = culture; ResourceSuffix = resource;
            View = view.Split('|'); Tools = tools.Split('|');
        }
    }

    /// <summary>Associe les langues prises en charge aux libellés de menus du VBE.</summary>
    internal static class UiLanguages
    {

        /// <summary>Catalogue des cultures et variantes de menus reconnues.</summary>
        internal static readonly UiLanguage[] All = {
            new UiLanguage("en-US", "", "view", "tools"),
            new UiLanguage("fr-FR", "French", "affichage", "outils"),
            new UiLanguage("es-ES", "Spanish", "ver", "herramientas"),
            new UiLanguage("de-DE", "German", "ansicht", "extras"),
            new UiLanguage("pt-BR", "Portuguese", "exibir|ver", "ferramentas"),
            new UiLanguage("it-IT", "Italian", "visualizza", "strumenti"),
            new UiLanguage("ja-JP", "Japanese", "表示", "ツール"),
            new UiLanguage("ko-KR", "Korean", "보기", "도구"),
            new UiLanguage("zh-CN", "ChineseSimplified", "视图", "工具"),
            new UiLanguage("zh-TW", "ChineseTraditional", "檢視|查看", "工具"),
            new UiLanguage("ru-RU", "Russian", "вид", "сервис|инструменты"),
            new UiLanguage("ar-SA", "Arabic", "عرض", "أدوات|ادوات"),
            new UiLanguage("hi-IN", "Hindi", "दृश्य|देखें", "उपकरण|टूल्स")
        };

        /// <summary>Sélectionne la langue prise en charge correspondant à la culture, avec un traitement explicite du chinois traditionnel.</summary>
        /// <param name="culture">Culture à associer ; null utilise le choix anglais par défaut.</param>
        /// <returns>Langue reconnue correspondante, ou langue anglaise.</returns>
        internal static UiLanguage For(CultureInfo culture)
        {
            string name = culture?.Name ?? "";
            string language = culture?.TwoLetterISOLanguageName ?? "en";
            if (language == "zh")
                return All.First(x => x.CultureName ==
                    (name.IndexOf("Hant", StringComparison.OrdinalIgnoreCase) >= 0 ||
                     name.EndsWith("-TW", StringComparison.OrdinalIgnoreCase) ||
                     name.EndsWith("-HK", StringComparison.OrdinalIgnoreCase) ||
                     name.EndsWith("-MO", StringComparison.OrdinalIgnoreCase) ? "zh-TW" : "zh-CN"));
            return All.FirstOrDefault(x => x.CultureName.StartsWith(language + "-", StringComparison.OrdinalIgnoreCase)) ?? All[0];
        }

        /// <summary>Retire le raccourci clavier et les marqueurs directionnels avant de comparer une légende de menu.</summary>
        /// <param name="caption">Légende de menu native, éventuellement null.</param>
        /// <returns>Légende normalisée en minuscules invariant.</returns>
        internal static string NormalizeMenu(string caption)
        {
            // CJK menus use a suffix such as 表示(&V); Latin menus put & inside the caption.
            return Regex.Replace(caption ?? "", @"\s*[\(（]&?[A-Za-z][\)）]\s*$", "")
                .Replace("&", "").Replace("\u200e", "").Replace("\u200f", "").Trim().ToLowerInvariant();
        }

        /// <summary>Vérifie si la légende normalisée correspond à un menu reconnu dans le catalogue.</summary>
        /// <param name="caption">Légende native à comparer.</param>
        /// <param name="view">Vrai pour le menu Affichage ; faux pour le menu Outils.</param>
        /// <returns>Vrai si au moins une langue contient la variante.</returns>
        internal static bool IsMenu(string caption, bool view)
        {
            string name = NormalizeMenu(caption);
            return All.Any(x => (view ? x.View : x.Tools).Contains(name));
        }

        /// <summary>Déduit la culture depuis les légendes de menus et utilise la culture de repli pour départager les égalités.</summary>
        /// <param name="captions">Légendes des menus principaux du VBE.</param>
        /// <param name="fallback">Culture qui départage des résultats de même score.</param>
        /// <returns>Culture correspondante ou culture anglaise si aucun menu ne correspond.</returns>
        internal static CultureInfo FromMenus(IEnumerable<string> captions, CultureInfo fallback)
        {
            var names = captions.Select(NormalizeMenu).ToArray();
            var scores = All.Select(x => new
            {
                Language = x,
                Score = names.Count(n => x.View.Contains(n)) * 2 + names.Count(n => x.Tools.Contains(n))
            }).ToArray();
            int max = scores.Max(x => x.Score);
            if (max == 0) return CultureInfo.GetCultureInfo("en-US");
            var best = scores.Where(x => x.Score == max).Select(x => x.Language).ToArray();
            var selected = best.FirstOrDefault(x => x == For(fallback)) ?? best[0];
            return CultureInfo.GetCultureInfo(selected.CultureName);
        }
    }
}
