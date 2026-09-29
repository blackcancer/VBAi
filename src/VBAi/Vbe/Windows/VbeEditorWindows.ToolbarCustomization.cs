using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;

namespace VBAi
{
    /// <summary>Crée et modifie les barres d’outils VBAi à partir des commandes natives du VBE.</summary>
    internal sealed partial class VbeEditorWindows
    {
        /// <summary>Préfixe réservé aux barres personnalisées créées par VBAi.</summary>
        private const string CustomToolbarPrefix = "VBAi - ";
        /// <summary>Préfixe commun aux identifiants des commandes ajoutées par VBAi.</summary>
        private const string CustomCommandTag = "VBAi.ToolbarCommand.";
                /// <summary>Retourne les commandes directes d’une barre et une version de personnalisation.</summary>
                /// <param name="request">Requête contenant le nom de barre issu de l’inventaire.</param>
                /// <returns>Les contrôles observés et leur version de personnalisation.</returns>
        public object ToolbarControls(Request request) => ReadToolbarCommands(FindNormalToolbar(request.ObjectName));
                /// <summary>Crée une barre personnalisée sans écraser une barre existante.</summary>
                /// <param name="request">Nom de barre et version attendue de la collection.</param>
                /// <returns>Résultat de création, état relu et version de collection mise à jour.</returns>
        public object CreateToolbar(Request request)
        {
            CheckToolbarCollection(request.ExpectedToolbarCollectionVersion);
            string suffix = (request.ObjectName ?? "").Trim();
            if (suffix.Length == 0 || suffix.Length > 64 || suffix.Any(char.IsControl))
                throw new ArgumentException("A printable toolbar name of 1-64 characters is required.");
            string name = CustomToolbarPrefix + suffix;
            foreach (dynamic bar in vbe.CommandBars)
                if (string.Equals((string)bar.Name, name, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("The toolbar already exists.");
            bool temporary = request.Temporary ?? true;
            object created = null; string error = null;
            try { created = (object)vbe.CommandBars.Add(name, 1, false, temporary); ((dynamic)created).Visible = true; if (!temporary) SaveToolbarProfile((dynamic)created, true); }
            catch (Exception ex) { error = ex.Message; }
            return new { Created = created != null, Verified = error == null && created != null, NativeError = error,
                Temporary = temporary, ObjectName = name, After = created == null ? null : ReadToolbarCommands(created),
                ToolbarCollectionVersion = ToolbarCollectionVersion(), PersistenceVerified = false,
                NextRead = "list_toolbars, toolbar_controls" };
        }
                /// <summary>Supprime uniquement une barre VBAi personnalisée vide et inchangée.</summary>
                /// <param name="request">Nom de barre et versions attendues de collection et de contrôles.</param>
                /// <returns>Indique si la barre a été supprimée et si son absence a été relue.</returns>
        public object RemoveToolbar(Request request)
        {
            CheckToolbarCollection(request.ExpectedToolbarCollectionVersion);
            dynamic bar = FindNormalToolbar(request.ObjectName);
            CheckToolbarCommands((object)bar, request.ExpectedToolbarControlsVersion);
            if ((bool)bar.BuiltIn || !((string)bar.Name).StartsWith(CustomToolbarPrefix, StringComparison.Ordinal) || (int)bar.Controls.Count != 0)
                throw new InvalidOperationException("Only an empty custom VBAi toolbar can be removed.");
            if (((int)bar.Protection & 1) != 0) throw new InvalidOperationException("Toolbar customization is protected.");
            bar.Delete();
            bool absent = true;
            foreach (dynamic current in vbe.CommandBars)
                if (string.Equals((string)current.Name, request.ObjectName, StringComparison.OrdinalIgnoreCase)) absent = false;
            // Un refus natif conserve le profil afin de permettre la restauration suivante.
            if (absent) ToolbarProfiles?.Update(request.ObjectName, null);
            return new { Removed = absent, Verified = absent, ToolbarCollectionVersion = ToolbarCollectionVersion() };
        }
                /// <summary>Ajoute une commande native existante, sans OnAction arbitraire.</summary>
                /// <param name="request">Barre, commande native exacte, index éventuel et version attendue des contrôles.</param>
                /// <returns>État de l’ajout et contrôles relus après l’appel natif.</returns>
        public object AddToolbarCommand(Request request)
        {
            dynamic bar = FindNormalToolbar(request.ObjectName);
            CheckToolbarCommands((object)bar, request.ExpectedToolbarControlsVersion);
            CheckToolbarCustomizable(bar);
            if (request.ControlId <= 1 || string.IsNullOrWhiteSpace(request.ControlCaption)) throw new ArgumentException("An exact built-in ControlId and caption are required.");
            dynamic source = vbe.CommandBars.FindControl(1, request.ControlId);
            if (source == null || !(bool)source.BuiltIn || (int)source.Type != 1 || (string)source.Caption != request.ControlCaption)
                throw new InvalidOperationException("The exact native built-in command is unavailable.");
            int position = request.InsertIndex ?? ((int)bar.Controls.Count + 1);
            if (position < 1 || position > (int)bar.Controls.Count + 1) throw new ArgumentOutOfRangeException(nameof(request.InsertIndex));
            bool temporary = request.Temporary ?? true;
            if (!temporary) RequirePersistentToolbar(bar);
            string tag = (temporary ? CustomCommandTag : PersistentCommandTag) + Guid.NewGuid().ToString("N");
            string error = null;
            try
            {
                dynamic added = source.Copy(bar, position);
                added.Tag = tag;
                if (!temporary) SaveToolbarProfile(bar, false);
            }
            catch (Exception ex) { error = ex.Message; }
            var after = ReadToolbarCommands((object)bar);
            var addedState = after.Controls.FirstOrDefault(x => x.Tag == tag);
            return new { Added = addedState != null, Verified = error == null && addedState != null && addedState.Id == request.ControlId && addedState.Caption == request.ControlCaption,
                Tag = tag, After = after, NativeError = error, Temporary = temporary, PersistenceVerified = false,
                NextRead = "toolbar_controls", Limit = "Inspect partial results; no automatic retry. Restart persistence must be qualified separately." };
        }
                /// <summary>Retire uniquement un bouton ajouté par VBAi après vérification de son identité et de son index.</summary>
                /// <param name="request">Barre, identité de commande, index et version attendue des contrôles.</param>
                /// <returns>Indique si la commande marquée VBAi a été retirée et relue comme absente.</returns>
        public object RemoveToolbarCommand(Request request)
        {
            dynamic bar = FindNormalToolbar(request.ObjectName);
            CheckToolbarCommands((object)bar, request.ExpectedToolbarControlsVersion);
            CheckToolbarCustomizable(bar);
            int position = request.InsertIndex ?? 0;
            if (position < 1 || position > (int)bar.Controls.Count) throw new ArgumentOutOfRangeException(nameof(request.InsertIndex));
            dynamic control = bar.Controls[position];
            string tag = (string)control.Tag ?? "";
            if (!tag.StartsWith(CustomCommandTag, StringComparison.Ordinal) || (int)control.Id != request.ControlId || (string)control.Caption != request.ControlCaption)
                throw new InvalidOperationException("Only the exact VBAi-added command can be removed.");
            string error = null;
            try { control.Delete(); SaveToolbarProfile(bar, false); } catch (Exception ex) { error = ex.Message; }
            var after = ReadToolbarCommands((object)bar);
            bool absent = !after.Controls.Any(x => x.Tag == tag);
            return new { Removed = absent, Verified = error == null && absent, After = after, NativeError = error };
        }
        /// <summary>Refuse la personnalisation d’une barre désactivée ou protégée.</summary>
        /// <param name="bar">Barre native à vérifier.</param>
        private static void CheckToolbarCustomizable(dynamic bar)
        {
            if (!(bool)bar.Enabled || ((int)bar.Protection & 1) != 0) throw new InvalidOperationException("The toolbar is disabled or customization is protected.");
        }
        /// <summary>Vérifie que la collection des barres correspond à l’empreinte précédemment lue.</summary>
        /// <param name="expected">Version fournie par l’inventaire des barres.</param>
        private void CheckToolbarCollection(string expected)
        {
            if (string.IsNullOrWhiteSpace(expected) || !string.Equals(expected, ToolbarCollectionVersion(), StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Toolbar collection changed; read list_toolbars again.");
        }
        /// <summary>Vérifie que les commandes d’une barre correspondent à leur version précédemment lue.</summary>
        /// <param name="bar">Barre dont les commandes sont vérifiées.</param>
        /// <param name="expected">Version fournie par la lecture des contrôles de cette barre.</param>
        private static void CheckToolbarCommands(object bar, string expected)
        {
            if (string.IsNullOrWhiteSpace(expected) || !string.Equals(expected, ReadToolbarCommands(bar).ToolbarControlsVersion, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Toolbar commands changed; read toolbar_controls again.");
        }
        /// <summary>Calcule une empreinte stable des noms et types des barres actuellement ouvertes.</summary>
        /// <returns>Empreinte SHA-256 de la collection triée.</returns>
        private string ToolbarCollectionVersion()
        {
            var names = new List<string>();
            foreach (dynamic bar in vbe.CommandBars)
                names.Add((string)bar.Name + "\n" + (int)bar.Type);
            return ToolbarCustomizationHash(names.OrderBy(x => x, StringComparer.Ordinal).ToArray());
        }
        /// <summary>État observable d’un contrôle de barre d’outils.</summary>
        private sealed class ToolbarCommandState
        {
            /// <summary>Gets or sets the index.</summary>
            /// <value>The current value represented by this member.</value>
            public int Index { get; set; }
            /// <summary>Gets or sets the id.</summary>
            /// <value>The current value represented by this member.</value>
            public int Id { get; set; }
            /// <summary>Gets or sets the type.</summary>
            /// <value>The current value represented by this member.</value>
            public int Type { get; set; }
            /// <summary>Gets or sets the caption.</summary>
            /// <value>The current value represented by this member.</value>
            public string Caption { get; set; }
            /// <summary>Gets or sets the tag.</summary>
            /// <value>The current value represented by this member.</value>
            public string Tag { get; set; }
            /// <summary>Gets or sets the built in.</summary>
            /// <value>The current value represented by this member.</value>
            public bool BuiltIn { get; set; }
            /// <summary>Gets or sets the visible.</summary>
            /// <value>The current value represented by this member.</value>
            public bool Visible { get; set; }
        }
        /// <summary>Instantané sérialisable des contrôles d’une barre et de leur version.</summary>
        private sealed class ToolbarCommandsState
        {
            /// <summary>Gets or sets the object name.</summary>
            /// <value>The current value represented by this member.</value>
            public string ObjectName { get; set; }
            /// <summary>Gets or sets the toolbar controls version.</summary>
            /// <value>The current value represented by this member.</value>
            public string ToolbarControlsVersion { get; set; }
            /// <summary>Gets or sets the controls.</summary>
            /// <value>The current value represented by this member.</value>
            public ToolbarCommandState[] Controls { get; set; }
        }
        /// <summary>Lit les contrôles directs d’une barre et calcule la version vérifiable de leur état.</summary>
        /// <param name="bar">Barre native à lire.</param>
        /// <returns>Instantané des commandes avec leur empreinte de contrôle.</returns>
        private static ToolbarCommandsState ReadToolbarCommands(dynamic bar)
        {
            var controls = new List<ToolbarCommandState>();
            foreach (dynamic control in bar.Controls)
                controls.Add(new ToolbarCommandState { Index = (int)control.Index, Id = (int)control.Id, Type = (int)control.Type,
                    Caption = (string)control.Caption, Tag = (string)control.Tag ?? "", BuiltIn = (bool)control.BuiltIn, Visible = (bool)control.Visible });
            var state = new ToolbarCommandsState { ObjectName = (string)bar.Name, Controls = controls.ToArray() };
            state.ToolbarControlsVersion = ToolbarCustomizationHash(new { state.ObjectName, state.Controls, Protection = (int)bar.Protection, Enabled = (bool)bar.Enabled });
            return state;
        }
        /// <summary>Calcule l’empreinte SHA-256 de l’état sérialisé fourni.</summary>
        /// <param name="state">État sérialisable dont la version est calculée.</param>
        /// <returns>Empreinte hexadécimale en minuscules.</returns>
        private static string ToolbarCustomizationHash(object state)
        {
            using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(new JavaScriptSerializer().Serialize(state)))).Replace("-", "").ToLowerInvariant();
        }
    }
}
