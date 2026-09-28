using System;

namespace CodexVBE
{
    /// <summary>External effects used by the native Designer-backed Markdown renderer.</summary>
    internal static class ChatMarkdown
    {
        internal static Action<string> CopyText = System.Windows.Clipboard.SetText;
        internal static Action<string> OpenLink = SafeLinks.Open;
    }
}
