using System;
using System.IO;
using System.Linq;

namespace CodexVBE
{
    /// <summary>Profil borné des seules barres personnalisées persistantes créées par VBAi.</summary>
    internal sealed class VbeToolbarProfiles
    {
        internal sealed class Command { public Command() { } public int Id { get; set; } public string Caption { get; set; } public string Tag { get; set; } }
        internal sealed class Bar { public Bar() { } public string Name { get; set; } public bool Visible { get; set; } public int Position { get; set; } public int? Left { get; set; } public int? Top { get; set; } public int? RowIndex { get; set; } public Command[] Commands { get; set; } }
        private readonly string path;
        /// <summary>Utilise un fichier privé du complément ou un emplacement de test explicitement fourni.</summary>
        internal VbeToolbarProfiles(string path) { this.path = Path.GetFullPath(path); }
        internal Bar[] Read()
        {
            using (var store = new ChatSessionStore(path))
            {
                var bars = store.ReadToolbarProfiles(); Validate(bars); return bars;
            }
        }
        /// <summary>Fusionne une seule barre dans une transaction SQLite, sans remplacer le fichier de base.</summary>
        internal void Update(string name, Bar state)
        {
            using (var store = new ChatSessionStore(path)) store.UpdateToolbarProfile(name, state, Validate);
        }
        private static void Validate(Bar[] bars)
        {
            if (bars == null || bars.Length > 32 || bars.Any(x => x == null) || bars.Select(x => x.Name).Distinct(StringComparer.OrdinalIgnoreCase).Count() != bars.Length)
                throw new InvalidOperationException("Invalid toolbar profile collection.");
            foreach (var bar in bars)
                if (bar == null || string.IsNullOrEmpty(bar.Name) || !bar.Name.StartsWith("VBAi - ", StringComparison.Ordinal) || bar.Name.Length > 71 || bar.Name.Any(char.IsControl) ||
                    bar.Position < 0 || bar.Position > 4 || bar.Left < -32768 || bar.Left > 32767 || bar.Top < -32768 || bar.Top > 32767 || bar.RowIndex < 1 || bar.Commands == null || bar.Commands.Length > 128 ||
                    bar.Commands.Any(x => x == null || x.Id <= 1 || string.IsNullOrWhiteSpace(x.Caption) || x.Caption.Length > 1024 ||
                        string.IsNullOrEmpty(x.Tag) || !x.Tag.StartsWith("VBAi.ToolbarCommand.Persistent.", StringComparison.Ordinal) || x.Tag.Length > 96) ||
                    bar.Commands.Select(x => x.Tag).Distinct(StringComparer.Ordinal).Count() != bar.Commands.Length)
                    throw new InvalidOperationException("Invalid toolbar profile contents.");
        }
    }
}
