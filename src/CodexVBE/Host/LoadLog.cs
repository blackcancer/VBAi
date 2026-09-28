using System;
using System.IO;

namespace CodexVBE
{
    /// <summary>Écrit des diagnostics de chargement dans un fichier temporaire.</summary>
    internal static class LoadLog
    {
        /// <summary>Chemin complet du journal de chargement du complément.</summary>
        internal static readonly string PathName = Path.Combine(Path.GetTempPath(), "CodexVBE-load.log");

        /// <summary>Ajoute le message daté au journal et ignore les refus d’accès et erreurs d’écriture.</summary>
        /// <param name="message">Message de diagnostic à ajouter.</param>
        internal static void Write(string message)
        {
            try
            {
                File.AppendAllText(PathName, DateTime.Now.ToString("o") + " " + message + Environment.NewLine);
            }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }
}
