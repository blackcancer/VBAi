using System;
using System.IO;
using System.Linq;

namespace CodexVBE
{
    /// <summary>Profil borné des seules barres personnalisées persistantes créées par VBAi.</summary>
    internal sealed class VbeToolbarProfiles
    {
        /// <summary>Commande native copiée dans une barre persistante.</summary>
        internal sealed class Command { /// <summary>Crée un enregistrement de commande vide pour la désérialisation.</summary>
public Command() { } /// <summary>Identifiant de la commande intégrée au VBE.</summary>
            /// <value>Identifiant numérique de la commande.</value>
public int Id { get; set; } /// <summary>Légende attendue pour l’identification de la commande native.</summary>
            /// <value>Légende enregistrée pour vérifier la correspondance native.</value>
public string Caption { get; set; } /// <summary>Marqueur persistant unique apposé à la copie.</summary>
            /// <value>Tag stable de la commande persistante.</value>
public string Tag { get; set; } }
        /// <summary>État persistant d’une barre VBAi, incluant placement et commandes reconnues.</summary>
        internal sealed class Bar { /// <summary>Crée un enregistrement de barre vide pour la désérialisation.</summary>
public Bar() { } /// <summary>Nom complet de la barre personnalisée.</summary>
            /// <value>Nom VBAi complet de la barre.</value>
public string Name { get; set; } /// <summary>État de visibilité à restaurer.</summary>
            /// <value><see langword="true"/> si la barre doit être visible.</value>
public bool Visible { get; set; } /// <summary>Mode d’ancrage ou de flottement natif.</summary>
            /// <value>Valeur Position native enregistrée.</value>
public int Position { get; set; } /// <summary>Position horizontale d’une barre flottante, si définie.</summary>
            /// <value>Coordonnée Left en pixels, ou <see langword="null"/>.</value>
public int? Left { get; set; } /// <summary>Position verticale d’une barre flottante, si définie.</summary>
            /// <value>Coordonnée Top en pixels, ou <see langword="null"/>.</value>
public int? Top { get; set; } /// <summary>Rangée d’ancrage, si définie.</summary>
            /// <value>Indice de rangée ou <see langword="null"/> pour une barre flottante.</value>
public int? RowIndex { get; set; } /// <summary>Commandes VBAi persistantes à rétablir.</summary>
            /// <value>Commandes reconnues du profil.</value>
public Command[] Commands { get; set; } }
        /// <summary>Chemin absolu du stockage des profils.</summary>
        private readonly string path;
                /// <summary>Utilise un fichier privé du complément ou un emplacement de test explicitement fourni.</summary>
        /// <param name="path">Chemin du fichier de base de données qui contient les profils.</param>
        internal VbeToolbarProfiles(string path) { this.path = Path.GetFullPath(path); }
        /// <summary>Lit et valide les profils de barres stockés.</summary>
        /// <returns>Les profils validés, ou un tableau vide lorsque le stockage n’en contient aucun.</returns>
        internal Bar[] Read()
        {
            using (var store = new ChatSessionStore(path))
            {
                var bars = store.ReadToolbarProfiles(); Validate(bars); return bars;
            }
        }
                /// <summary>Fusionne une seule barre dans une transaction SQLite, sans remplacer le fichier de base.</summary>
        /// <param name="name">Nom de barre dont l’entrée doit être créée, remplacée ou supprimée.</param>
        /// <param name="state">Nouvel état, ou <see langword="null"/> pour supprimer le profil.</param>
        internal void Update(string name, Bar state)
        {
            using (var store = new ChatSessionStore(path)) store.UpdateToolbarProfile(name, state, Validate);
        }
        /// <summary>Vérifie les limites de collection et valide chaque profil.</summary>
        /// <param name="bars">Profils chargés ou candidats à l’écriture.</param>
        private static void Validate(Bar[] bars)
        {
            if (bars == null || bars.Length > 32 || bars.Any(x => x == null) || bars.Select(x => x.Name).Distinct(StringComparer.OrdinalIgnoreCase).Count() != bars.Length)
                throw new InvalidOperationException("Invalid toolbar profile collection.");
            foreach (var bar in bars)
                ValidateContents(bar);
        }
                /// <summary>Valide également une entrée isolée avant toute utilisation de ses propriétés.</summary>
        /// <param name="bar">Profil individuel à valider.</param>
        private static void ValidateContents(Bar bar)
        {
            if (bar == null || string.IsNullOrEmpty(bar.Name) || !bar.Name.StartsWith("VBAi - ", StringComparison.Ordinal) || bar.Name.Length > 71 || bar.Name.Any(char.IsControl) ||
                bar.Position < 0 || bar.Position > 4 || bar.Left < -32768 || bar.Left > 32767 || bar.Top < -32768 || bar.Top > 32767 || bar.RowIndex < 1 || bar.Commands == null || bar.Commands.Length > 128 ||
                bar.Commands.Any(x => x == null || x.Id <= 1 || string.IsNullOrWhiteSpace(x.Caption) || x.Caption.Length > 1024 ||
                    string.IsNullOrEmpty(x.Tag) || !x.Tag.StartsWith("VBAi.ToolbarCommand.Persistent.", StringComparison.Ordinal) || x.Tag.Length > 96) ||
                bar.Commands.Select(x => x.Tag).Distinct(StringComparer.Ordinal).Count() != bar.Commands.Length)
                throw new InvalidOperationException("Invalid toolbar profile contents.");
        }
    }
}
