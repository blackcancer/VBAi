using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;

namespace CodexVBE
{
    /// <summary>Regroupe les activités du fournisseur dans la projection visuelle de la conversation.</summary>
    internal sealed partial class ChatWindow
    {
        /// <summary>Activités consécutives affichées sous leur première entrée, sans modifier l'historique persisté.</summary>
        private readonly Dictionary<ChatEntry, List<ChatEntry>> activityGroups = new Dictionary<ChatEntry, List<ChatEntry>>();
        /// <summary>Première entrée visible du groupe auquel appartient chaque activité.</summary>
        private readonly Dictionary<ChatEntry, ChatEntry> activityOwners = new Dictionary<ChatEntry, ChatEntry>();
        /// <summary>Groupes ouverts par l'utilisateur, conservés pendant les mises à jour et le recyclage.</summary>
        private readonly HashSet<ChatEntry> expandedActivityGroups = new HashSet<ChatEntry>();

        /// <summary>Identifie une activité textuelle dépourvue de carte interactive.</summary>
        /// <param name="entry">Entrée à examiner.</param>
        /// <returns>Vrai pour un outil ou un résumé de réflexion.</returns>
        private static bool IsActivity(ChatEntry entry) => entry != null && entry.Change == null && entry.FormCut == null &&
            (entry.Speaker == "Outil" || entry.Speaker == "Réflexion");

        /// <summary>Regroupe uniquement des activités adjacentes appartenant au même tour.</summary>
        /// <param name="previous">Entrée précédente.</param>
        /// <param name="entry">Entrée suivante.</param>
        /// <returns>Vrai si aucun message ni changement de tour ne sépare les activités.</returns>
        private static bool CanGroupActivities(ChatEntry previous, ChatEntry entry) =>
            IsActivity(previous) && IsActivity(entry) && previous.TurnId == entry.TurnId;

        /// <summary>Ajoute une entrée à la projection visible, en fusionnant les activités adjacentes.</summary>
        /// <param name="entry">Entrée conservée dans l'historique brut.</param>
        /// <param name="previous">Entrée brute précédente, ou null au début de la fenêtre.</param>
        private void AppendVisibleEntry(ChatEntry entry, ChatEntry previous)
        {
            if (!IsActivity(entry)) { visibleEntries.Add(entry); return; }
            if (CanGroupActivities(previous, entry) && activityOwners.TryGetValue(previous, out var owner))
            {
                activityGroups[owner].Add(entry); activityOwners[entry] = owner;
                RefreshVisibleActivity(owner);
                return;
            }
            activityGroups[entry] = new List<ChatEntry> { entry }; activityOwners[entry] = entry;
            visibleEntries.Add(entry);
        }

        /// <summary>Actualise une entrée ou son groupe sans dupliquer les lignes du transcript.</summary>
        /// <param name="entry">Entrée modifiée par le fournisseur.</param>
        private void RefreshVisibleActivity(ChatEntry entry)
        {
            if (activityOwners.TryGetValue(entry, out var owner)) entry = owner;
            int index = visibleEntries.IndexOf(entry);
            if (index >= 0) { visibleEntries.RemoveAt(index); visibleEntries.Insert(index, entry); }
        }

        /// <summary>Construit un unique bloc repliable contenant les activités et résumés dans leur ordre d'origine.</summary>
        /// <param name="owner">Première entrée servant d'identité stable au bloc.</param>
        /// <param name="entries">Activités à afficher.</param>
        /// <returns>Bloc repliable, fermé par défaut et conservant son état lors des mises à jour.</returns>
        private FrameworkElement RenderActivityGroup(ChatEntry owner, List<ChatEntry> entries)
        {
            var body = new StackPanel { Margin = new Thickness(18, 8, 4, 6) };
            foreach (var entry in entries)
            {
                if (entry.Speaker == "Réflexion") body.Children.Add(new TextBlock {
                    Text = UiText.Get("Reasoning · summary"), Foreground = Ink("#64748B"), FontSize = 12,
                    Margin = new Thickness(0, 4, 0, 4) });
                var text = SelectableText(entry.Text); text.Margin = new Thickness(0, 0, 0, 6);
                if (!string.IsNullOrEmpty(entry.StreamId)) liveTexts[entry.StreamId] = text;
                body.Children.Add(text);
            }
            var group = new Expander { Header = UiText.Get("Agent activity") + " · " + entries.Count,
                IsExpanded = expandedActivityGroups.Contains(owner), Content = body, FontSize = 12,
                Foreground = Ink("#64748B"), Margin = new Thickness(0, 4, 0, 14) };
            group.Expanded += (s, e) => expandedActivityGroups.Add(owner);
            group.Collapsed += (s, e) => expandedActivityGroups.Remove(owner);
            return group;
        }

        /// <summary>Libère les champs de flux d'un groupe lorsqu'il quitte la fenêtre virtualisée.</summary>
        /// <param name="owner">Entrée visible du groupe libéré.</param>
        private void ReleaseActivityTexts(ChatEntry owner)
        {
            if (!activityGroups.TryGetValue(owner, out var entries)) return;
            foreach (var entry in entries) if (entry.StreamId != null) liveTexts.Remove(entry.StreamId);
        }
    }
}
