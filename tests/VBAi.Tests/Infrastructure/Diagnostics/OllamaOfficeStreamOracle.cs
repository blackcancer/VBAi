using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace VBAi.Tests.Integration
{
    /// <summary>Excludes echoed user prompts and settings labels from native streaming acceptance.</summary>
    internal static class OllamaOfficeStreamOracle
    {
        internal static bool IsNumberedResponse(string text) => text != null &&
            Regex.IsMatch(text, @"(?m)^\s*1[.)]\s+\S") && text.Length > 20;

        /// <summary>Accepts a rendered list in one text view or three ordered short items in separate views.</summary>
        internal static bool HasNumberedResponse(IEnumerable<string> visibleTexts)
        {
            if (visibleTexts == null) return false;
            var texts = visibleTexts.ToArray();
            if (texts.Any(IsNumberedResponse)) return true;

            // Exact whole-view matches exclude the quoted examples inside the echoed prompt.
            int next = 1;
            foreach (string text in texts)
            {
                if (text?.Trim() == next + ". Object " + next && ++next == 4) return true;
            }
            return false;
        }

        internal static bool IsReadyResponse(string text) => text?.Trim() == "UI_READY_42";
    }
}
