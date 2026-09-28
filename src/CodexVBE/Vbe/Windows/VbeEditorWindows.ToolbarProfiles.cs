using System;
using System.Collections.Generic;
using System.Linq;

namespace CodexVBE
{
    /// <summary>Restaure et persiste les profils explicites de barres d’outils VBAi.</summary>
internal sealed partial class VbeEditorWindows
    {
        /// <summary>Préfixe distinguant les commandes conservées dans un profil persistant.</summary>
private const string PersistentCommandTag = "VBAi.ToolbarCommand.Persistent.";
        /// <summary>Stockage facultatif des profils de barres d’outils.</summary>
internal VbeToolbarProfiles ToolbarProfiles;
        /// <summary>Erreurs rencontrées lors de la restauration des profils de barres d’outils.</summary>
internal readonly List<string> ToolbarProfileErrors = new List<string>();
        /// <summary>Restaure les commandes natives exactes des seules barres VBAi enregistrées explicitement.</summary>
        internal void RestoreToolbarProfiles()
        {
            if (ToolbarProfiles == null) return;
            VbeToolbarProfiles.Bar[] profiles;
            try { profiles = ToolbarProfiles.Read(); } catch (Exception ex) { ToolbarProfileErrors.Add(ex.Message); return; }
            foreach (var profile in profiles)
            {
                try
                {
                    dynamic bar = null;
                    foreach (dynamic candidate in vbe.CommandBars)
                        if (((string)candidate.Name).Equals(profile.Name, StringComparison.OrdinalIgnoreCase))
                        { if (bar != null) throw new InvalidOperationException("Toolbar profile name is ambiguous."); bar = candidate; }
                    if (bar == null) bar = vbe.CommandBars.Add(profile.Name, profile.Position, false, true);
                    if ((bool)bar.BuiltIn || (int)bar.Type != 0) throw new InvalidOperationException("Toolbar profile collides with a native menu or toolbar.");
                    CheckToolbarCustomizable(bar);
                    foreach (var command in profile.Commands)
                    {
                        var existing = ReadToolbarCommands((object)bar).Controls.Where(x => x.Tag == command.Tag).ToArray();
                        if (existing.Length == 1 && existing[0].Id == command.Id && existing[0].Type == 1 && existing[0].BuiltIn) continue;
                        if (existing.Length != 0) throw new InvalidOperationException("A persisted command identity is ambiguous or changed.");
                        dynamic source = vbe.CommandBars.FindControl(1, command.Id);
                        if (source == null || !(bool)source.BuiltIn || (int)source.Type != 1)
                            throw new InvalidOperationException("A persisted built-in command is absent or localized differently.");
                        dynamic added = source.Copy(bar, (int)bar.Controls.Count + 1);
                        added.Tag = command.Tag;
                    }
                    bar.Position = profile.Position;
                    if (profile.Position == 4 && profile.Left.HasValue && profile.Top.HasValue)
                    { bar.Left = profile.Left.Value; bar.Top = profile.Top.Value; }
                    else if (profile.Position != 4 && profile.RowIndex.HasValue) bar.RowIndex = profile.RowIndex.Value;
                    bar.Visible = profile.Visible;
                }
                catch (Exception ex) { ToolbarProfileErrors.Add(profile.Name + ": " + ex.Message); }
            }
        }
        /// <summary>Retire à la déconnexion les seules copies de commandes marquées pour la session.</summary>
        internal void RemoveTemporaryToolbarCommands()
        {
            foreach (dynamic bar in vbe.CommandBars)
                if ((int)bar.Type == 0)
                    for (int index = (int)bar.Controls.Count; index >= 1; index--)
                    {
                        dynamic control = bar.Controls[index]; string tag = (string)control.Tag;
                        if (System.Text.RegularExpressions.Regex.IsMatch(tag ?? "", @"^VBAi\.ToolbarCommand\.[0-9a-f]{32}$")) control.Delete();
                    }
        }
        /// <summary>Enregistre l’état et les commandes persistantes d’une barre si son profil est suivi.</summary>
        /// <param name="bar">Barre native à capturer.</param>
        /// <param name="create">Indique si un nouveau profil peut être créé.</param>
private void SaveToolbarProfile(dynamic bar, bool create)
        {
            if (ToolbarProfiles == null) return;
            string name = (string)bar.Name;
            if (!create && !ToolbarProfiles.Read().Any(x => x.Name.Equals(name, StringComparison.OrdinalIgnoreCase))) return;
            var controls = ReadToolbarCommands((object)bar).Controls;
            var profile = new VbeToolbarProfiles.Bar { Name = name, Position = (int)bar.Position, Visible = (bool)bar.Visible,
                Left = (int)bar.Position == 4 ? (int?)(int)bar.Left : null, Top = (int)bar.Position == 4 ? (int?)(int)bar.Top : null,
                RowIndex = (int)bar.Position == 4 ? null : (int?)(int)bar.RowIndex,
                Commands = controls.Where(x => x.Tag.StartsWith(PersistentCommandTag, StringComparison.Ordinal)).Select(x =>
                    new VbeToolbarProfiles.Command { Id = x.Id, Caption = x.Caption, Tag = x.Tag }).ToArray() };
            ToolbarProfiles.Update(name, profile);
        }
        /// <summary>Exige un profil existant avant l’ajout de commandes persistantes à une barre.</summary>
        /// <param name="bar">Barre native sur laquelle une commande persistante doit être ajoutée.</param>
private void RequirePersistentToolbar(dynamic bar)
        {
            if (ToolbarProfiles != null && !ToolbarProfiles.Read().Any(x => x.Name.Equals((string)bar.Name, StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException("Persistent buttons require a custom VBAi toolbar created with Temporary=false.");
        }
    }
}
