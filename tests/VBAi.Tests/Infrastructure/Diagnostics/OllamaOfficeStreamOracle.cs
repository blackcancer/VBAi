using System.Text.RegularExpressions;

namespace VBAi.Tests.Integration
{
    /// <summary>Excludes echoed user prompts and settings labels from native streaming acceptance.</summary>
    internal static class OllamaOfficeStreamOracle
    {
        internal static bool IsNumberedResponse(string text) => text != null &&
            Regex.IsMatch(text, @"(?m)^\s*1[.)]\s+\S") && text.Length > 20;
        internal static bool IsReadyResponse(string text) => text?.Trim() == "UI_READY_42";
    }
}
