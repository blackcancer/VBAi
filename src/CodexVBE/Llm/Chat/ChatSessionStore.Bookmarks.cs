using System;
using System.Collections.Generic;
using System.Linq;

namespace CodexVBE
{
    /// <summary>Repère persistant vers une position de code dans un module d’un projet.</summary>
    internal sealed class CodeBookmark
    {
        /// <summary>Obtient ou définit le libellé unique du repère.</summary>
        /// <value>Nom utilisé pour retrouver le repère sans tenir compte de la casse.</value>
        public string Name { get; set; }
        /// <summary>Obtient ou définit le nom du module ciblé.</summary>
        /// <value>Nom du module VBA contenant la position.</value>
        public string Module { get; set; }
        /// <summary>Obtient ou définit l’empreinte SHA-256 du code lors de la création.</summary>
        /// <value>Empreinte source qui permet de contextualiser le repère.</value>
        public string Sha256 { get; set; }
        /// <summary>Obtient ou définit la ligne, indexée à partir de un.</summary>
        /// <value>Numéro de ligne dans le module.</value>
        public int Line { get; set; }
        /// <summary>Obtient ou définit la colonne, indexée à partir de un.</summary>
        /// <value>Numéro de colonne dans la ligne.</value>
        public int Column { get; set; }
    }
    /// <summary>Persistance SQLite des conversations et de leurs repères de code.</summary>
    internal sealed partial class ChatSessionStore
    {
        /// <summary>Exécute une étape SQLite native ; la validation et les transactions restent dans le magasin.</summary>
        internal Func<IntPtr, int> StepNative = Native.sqlite3_step;
        /// <summary>Charge les repères du projet, triés sans tenir compte de la casse.</summary>
        /// <param name="scope">Clé du projet dont les repères sont demandés.</param>
        /// <returns>Repères désérialisés associés à cette clé.</returns>
        public List<CodeBookmark> ListBookmarks(string scope)
        {
            var result = new List<CodeBookmark>();
            using (var statement = Prepare("SELECT payload FROM code_bookmarks WHERE scope = ?1 ORDER BY name_key", scope.ToUpperInvariant()))
                while (statement.Step() == 100) result.Add(json.Deserialize<CodeBookmark>(ReadText(Native.sqlite3_column_text(statement.Handle, 0))));
            return result;
        }
        /// <summary>Insère ou remplace un repère dans une transaction, avec une limite de 200 par projet.</summary>
        /// <param name="scope">Clé du projet propriétaire.</param>
        /// <param name="bookmark">Repère à enregistrer.</param>
        public void SaveBookmark(string scope, CodeBookmark bookmark)
        {
            Execute("BEGIN IMMEDIATE");
            try
            {
                var existing = ListBookmarks(scope);
                if (existing.Count >= 200 && !existing.Any(x => string.Equals(x.Name, bookmark.Name, StringComparison.OrdinalIgnoreCase)))
                    throw new InvalidOperationException("The macro already has 200 bookmarks.");
                Execute("INSERT OR REPLACE INTO code_bookmarks(scope, name_key, payload) VALUES (?1, ?2, ?3)",
                    scope.ToUpperInvariant(), bookmark.Name.ToUpperInvariant(), json.Serialize(bookmark));
                Execute("COMMIT");
            }
            catch { Execute("ROLLBACK"); throw; }
        }
        /// <summary>Supprime un repère par son nom insensible à la casse.</summary>
        /// <param name="scope">Clé du projet propriétaire.</param>
        /// <param name="name">Nom du repère à supprimer.</param>
        /// <returns><see langword="true"/> si une ligne a été supprimée.</returns>
        public bool RemoveBookmark(string scope, string name)
        {
            Execute("DELETE FROM code_bookmarks WHERE scope = ?1 AND name_key = ?2", scope.ToUpperInvariant(), name.ToUpperInvariant());
            using (var statement = Prepare("SELECT changes()"))
                return statement.Step() == 100 && ReadText(Native.sqlite3_column_text(statement.Handle, 0)) == "1";
        }
    }
}
