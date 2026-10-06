using System;

namespace VBAi
{

    /// <summary>External effects used by the native Designer-backed Markdown renderer.</summary>
    internal static class ChatMarkdown
    {

        /// <summary>Action utilisée pour copier un bloc de code dans le presse-papiers.</summary>
        internal static Action<string> CopyText = System.Windows.Clipboard.SetText;

        /// <summary>Action utilisée pour ouvrir un lien HTTP(S) validé.</summary>
        internal static Action<string> OpenLink = SafeLinks.Open;
    }
}
