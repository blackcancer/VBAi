using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace CodexVBE
{
    internal sealed class UiLanguage
    {
        internal readonly string CultureName, ResourceSuffix;
        internal readonly string[] View, Tools;
        internal UiLanguage(string culture, string resource, string view, string tools)
        {
            CultureName = culture; ResourceSuffix = resource;
            View = view.Split('|'); Tools = tools.Split('|');
        }
    }

    internal static class UiLanguages
    {
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

        internal static string NormalizeMenu(string caption)
        {
            // CJK menus use a suffix such as 表示(&V); Latin menus put & inside the caption.
            return Regex.Replace(caption ?? "", @"\s*[\(（]&?[A-Za-z][\)）]\s*$", "")
                .Replace("&", "").Replace("\u200e", "").Replace("\u200f", "").Trim().ToLowerInvariant();
        }

        internal static bool IsMenu(string caption, bool view)
        {
            string name = NormalizeMenu(caption);
            return All.Any(x => (view ? x.View : x.Tools).Contains(name));
        }

        internal static CultureInfo FromMenus(IEnumerable<string> captions, CultureInfo fallback)
        {
            var names = captions.Select(NormalizeMenu).ToArray();
            var scores = All.Select(x => new { Language = x,
                Score = names.Count(n => x.View.Contains(n)) * 2 + names.Count(n => x.Tools.Contains(n)) }).ToArray();
            int max = scores.Max(x => x.Score);
            if (max == 0) return CultureInfo.GetCultureInfo("en-US");
            var best = scores.Where(x => x.Score == max).Select(x => x.Language).ToArray();
            var selected = best.FirstOrDefault(x => x == For(fallback)) ?? best[0];
            return CultureInfo.GetCultureInfo(selected.CultureName);
        }
    }
}
