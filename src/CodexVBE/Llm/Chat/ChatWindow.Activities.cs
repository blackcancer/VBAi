using System.Collections.Generic;
using System.Linq;
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
        /// <summary>Actions dont les détails ont été ouverts par l'utilisateur.</summary>
        private readonly HashSet<ChatEntry> expandedActivitySteps = new HashSet<ChatEntry>();

        /// <summary>Ajoute ou actualise une étape native, sans dupliquer son identité.</summary>
        /// <param name="activity">Données reçues du fournisseur.</param>
        private void ReceiveAgentActivity(CodexAgentActivity activity)
        {
            if (IsDisposed || conversationItems == null || activity == null || string.IsNullOrEmpty(activity.Id)) return;
            if (activity.Kind == "reasoning" && string.IsNullOrWhiteSpace(activity.Detail) && !liveEntries.ContainsKey(activity.Id)) return;
            if (!liveEntries.TryGetValue(activity.Id, out var entry))
            {
                entry = new ChatEntry { Speaker = activity.Kind == "reasoning" ? "Réflexion" : "Outil", StreamId = activity.Id };
                liveEntries[activity.Id] = entry;
                entry.Activity = new CodexAgentActivity { Id = activity.Id, Kind = activity.Kind };
                AddEntry(entry);
            }
            var previous = entry.Activity;
            entry.Activity = new CodexAgentActivity { Id = activity.Id, Kind = activity.Kind,
                Title = string.IsNullOrWhiteSpace(activity.Title) ? previous?.Title : activity.Title,
                Detail = activity.Append ? CodexAgentActivity.Limit((previous?.Detail ?? "") + activity.Detail) : CodexAgentActivity.Limit(activity.Detail),
                Status = activity.Append && previous != null && previous.Status != "inProgress" ? previous.Status : activity.Status, DurationMs = activity.DurationMs ?? previous?.DurationMs };
            entry.Text = entry.Activity.Title + "\n" + entry.Activity.Detail;
            if (activity.Status != "inProgress") completedStreams.Add(activity.Id);
            RefreshVisibleActivity(entry); FollowLatest(); ScheduleSessionSave();
        }

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
                if (entry.Activity != null) { body.Children.Add(RenderActivityStep(entry)); continue; }
                if (entry.Speaker == "Réflexion") body.Children.Add(new TextBlock {
                    Text = UiText.Get("Reasoning · summary"), Foreground = Ink("#64748B"), FontSize = 12,
                    Margin = new Thickness(0, 4, 0, 4) });
                var text = SelectableText(entry.Text); text.Margin = new Thickness(0, 0, 0, 6);
                if (!string.IsNullOrEmpty(entry.StreamId)) liveTexts[entry.StreamId] = text;
                body.Children.Add(text);
            }
            var latest = entries.LastOrDefault(entry => entry.Activity?.Status == "inProgress") ?? entries.Last();
            string preview = latest.Activity == null ? "" : " · " + CodexAgentActivity.Limit(latest.Activity.Title);
            if (preview.Length > 90) preview = preview.Substring(0, 87) + "…";
            var group = new Expander { Header = UiText.Get("Agent activity") + " · " + entries.Count + preview,
                IsExpanded = expandedActivityGroups.Contains(owner), Content = body, FontSize = 12,
                Foreground = Ink("#64748B"), Margin = new Thickness(0, 4, 0, 14) };
            group.Expanded += (s, e) => expandedActivityGroups.Add(owner);
            group.Collapsed += (s, e) => expandedActivityGroups.Remove(owner);
            return group;
        }

        /// <summary>Dessine une étape compacte avec résultat, durée native et détail dépliable.</summary>
        /// <param name="entry">Entrée enrichie de l'historique.</param>
        /// <returns>Étape de la chronologie.</returns>
        private FrameworkElement RenderActivityStep(ChatEntry entry)
        {
            var activity = entry.Activity;
            bool running = activity.Status == "inProgress", failed = activity.Status == "failed";
            string status = UiText.Get(running ? "In progress" : failed ? "Failed" :
                activity.Status == "declined" ? "Declined" : activity.Status == "completed" ? "Completed" : "Cancelled");
            var heading = new DockPanel { LastChildFill = true };
            var state = new TextBlock { Text = (running ? "● " : failed ? "× " : activity.Status == "completed" ? "✓ " : "— ") + status,
                Foreground = Ink(failed ? "#B91C1C" : running ? "#2563EB" : "#64748B"), FontSize = 11,
                Margin = new Thickness(12, 0, 0, 0), VerticalAlignment = VerticalAlignment.Center };
            if (activity.DurationMs.HasValue) state.Text += " · " + (activity.DurationMs.Value / 1000d).ToString("0.0", UiText.Culture) + " s";
            DockPanel.SetDock(state, System.Windows.Controls.Dock.Right); heading.Children.Add(state);
            heading.Children.Add(new TextBlock { Text = activity.Title, ToolTip = activity.Title, Foreground = Ink("#334155"),
                FontSize = 12, TextTrimming = TextTrimming.CharacterEllipsis, VerticalAlignment = VerticalAlignment.Center });
            var detail = SelectableText(activity.Detail, activity.Kind == "commandExecution");
            detail.Margin = new Thickness(18, 6, 4, 8);
            if (!string.IsNullOrEmpty(entry.StreamId)) liveTexts[entry.StreamId] = detail;
            var step = new Expander { Header = heading, Content = detail, IsExpanded = expandedActivitySteps.Contains(entry),
                HorizontalContentAlignment = HorizontalAlignment.Stretch, Margin = new Thickness(0, 2, 0, 4) };
            step.Expanded += (s, e) => { expandedActivitySteps.Add(entry); e.Handled = true; };
            step.Collapsed += (s, e) => { expandedActivitySteps.Remove(entry); e.Handled = true; };
            return step;
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
