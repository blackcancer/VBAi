using System;
using System.Diagnostics;
namespace VBAi
{

    /// <summary>Valide les liens externes et les ouvre uniquement avec HTTP ou HTTPS.</summary>
    internal static class SafeLinks
    {

        /// <summary>Accepte uniquement une URI absolue HTTP ou HTTPS sans informations utilisateur.</summary>
        /// <param name="text">Texte interprété comme URI.</param>
        /// <returns>Vrai si le schéma et les informations utilisateur respectent la règle.</returns>
        internal static bool Allowed(string text) { return Uri.TryCreate(text, UriKind.Absolute, out var uri) && (uri.Scheme == "https" || uri.Scheme == "http") && string.IsNullOrEmpty(uri.UserInfo); }

        /// <summary>Valide le lien puis le remet au lanceur de processus Windows.</summary>
        /// <param name="text">URI absolue à ouvrir.</param>
        internal static void Open(string text) { Open(text, Process.Start); }

        /// <summary>Refuse les protocoles non HTTP(S) puis utilise le lanceur fourni.</summary>
        /// <param name="text">URI absolue à ouvrir.</param>
        /// <param name="start">Fonction qui lance le navigateur ou le gestionnaire associé.</param>
        /// <exception cref="ArgumentException">Le lien n’est pas une URI absolue HTTP(S) sans informations utilisateur.</exception>
        internal static void Open(string text, Func<ProcessStartInfo, Process> start)
        {
            if (!Allowed(text)) throw new ArgumentException(UiText.Get("Only HTTP and HTTPS links can be opened."));
            start(new ProcessStartInfo(text) { UseShellExecute = true });
        }
    }
}
