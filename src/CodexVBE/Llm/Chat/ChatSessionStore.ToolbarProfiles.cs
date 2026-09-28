using System;
using System.Collections.Generic;
using System.Linq;

namespace CodexVBE
{
    internal sealed partial class ChatSessionStore
    {
        /// <summary>Lit les seules barres personnalisées du profil d’hôte courant avec une borne avant désérialisation.</summary>
        internal VbeToolbarProfiles.Bar[] ReadToolbarProfiles()
        {
            Execute("CREATE TABLE IF NOT EXISTS vbe_toolbar_profiles (name_key TEXT PRIMARY KEY, payload TEXT NOT NULL)");
            var result = new List<VbeToolbarProfiles.Bar>();
            using (var statement = Prepare("SELECT payload FROM vbe_toolbar_profiles ORDER BY name_key LIMIT 33"))
                while (statement.Step() == 100)
                {
                    string payload = ReadText(Native.sqlite3_column_text(statement.Handle, 0));
                    if (payload.Length > 128 * 1024) throw new InvalidOperationException("Toolbar profile payload exceeds its bound.");
                    result.Add(json.Deserialize<VbeToolbarProfiles.Bar>(payload));
                }
            return result.ToArray();
        }
        /// <summary>Valide et écrit le profil sous transaction, en conservant les autres barres.</summary>
        internal void UpdateToolbarProfile(string name, VbeToolbarProfiles.Bar state, Action<VbeToolbarProfiles.Bar[]> validate)
        {
            Execute("BEGIN IMMEDIATE");
            try
            {
                var original = ReadToolbarProfiles(); validate(original);
                var current = original.Where(x => !x.Name.Equals(name, StringComparison.OrdinalIgnoreCase)).ToList();
                if (state != null) current.Add(state);
                validate(current.ToArray());
                if (state == null) Execute("DELETE FROM vbe_toolbar_profiles WHERE name_key = ?1", name.ToUpperInvariant());
                else
                {
                    if (!state.Name.Equals(name, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("Toolbar profile key differs from its name.");
                    string payload = json.Serialize(state);
                    if (payload.Length > 128 * 1024) throw new InvalidOperationException("Toolbar profile payload exceeds its bound.");
                    Execute("INSERT OR REPLACE INTO vbe_toolbar_profiles(name_key,payload) VALUES(?1,?2)", name.ToUpperInvariant(), payload);
                }
                Execute("COMMIT");
            }
            catch { Execute("ROLLBACK"); throw; }
        }
    }
}
