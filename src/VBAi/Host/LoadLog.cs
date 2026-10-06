using System;
using System.IO;

namespace VBAi
{

    /// <summary>Écrit des diagnostics de chargement dans un fichier temporaire.</summary>
    internal static class LoadLog
    {

        /// <summary>Chemin complet du journal de chargement du complément.</summary>
        internal static readonly string PathName = Path.Combine(Path.GetTempPath(), "VBAi-load.log");

        /// <summary>Écrit les diagnostics au moyen du système de fichiers natif.</summary>
        internal static Action<string, string> AppendText = File.AppendAllText;

        /// <summary>Ajoute le message daté au journal et ignore les refus d’accès et erreurs d’écriture.</summary>
        /// <param name="message">Message de diagnostic à ajouter.</param>
        internal static void Write(string message)
        {
            try
            {
                AppendText(PathName, DateTime.Now.ToString("o") + " " + message + Environment.NewLine);
            }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }
}
