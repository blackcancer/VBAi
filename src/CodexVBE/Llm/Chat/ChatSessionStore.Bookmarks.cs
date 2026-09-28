using System;
using System.Collections.Generic;
using System.Linq;

namespace CodexVBE
{
    internal sealed class CodeBookmark
    {
        public string Name { get; set; }
        public string Module { get; set; }
        public string Sha256 { get; set; }
        public int Line { get; set; }
        public int Column { get; set; }
    }
    internal sealed partial class ChatSessionStore
    {
        /// <summary>Exécute une étape SQLite native ; la validation et les transactions restent dans le magasin.</summary>
        internal Func<IntPtr, int> StepNative = Native.sqlite3_step;
        public List<CodeBookmark> ListBookmarks(string scope)
        {
            var result = new List<CodeBookmark>();
            using (var statement = Prepare("SELECT payload FROM code_bookmarks WHERE scope = ?1 ORDER BY name_key", scope.ToUpperInvariant()))
                while (statement.Step() == 100) result.Add(json.Deserialize<CodeBookmark>(ReadText(Native.sqlite3_column_text(statement.Handle, 0))));
            return result;
        }
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
        public bool RemoveBookmark(string scope, string name)
        {
            Execute("DELETE FROM code_bookmarks WHERE scope = ?1 AND name_key = ?2", scope.ToUpperInvariant(), name.ToUpperInvariant());
            using (var statement = Prepare("SELECT changes()"))
                return statement.Step() == 100 && ReadText(Native.sqlite3_column_text(statement.Handle, 0)) == "1";
        }
    }
}
