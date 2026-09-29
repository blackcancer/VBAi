using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace VBAi
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
        /// <summary>Stores the collapsed activity groups used by ChatWindow.</summary>
        private readonly HashSet<ChatEntry> collapsedActivityGroups = new HashSet<ChatEntry>();
        /// <summary>Stores the collapsed activity steps used by ChatWindow.</summary>
        private readonly HashSet<ChatEntry> collapsedActivitySteps = new HashSet<ChatEntry>();

        // View state belongs to the control tree and is released with its virtualized owner.
        private sealed class ActivityGroupViewState
        {
            internal readonly Dictionary<ChatEntry, System.Windows.Forms.Control> Rows =
                new Dictionary<ChatEntry, System.Windows.Forms.Control>();
            internal bool Updating;
        }

        private sealed class ActivityStepViewState
        {
            internal string Kind;
            internal bool Initialized, Updating;
        }

        /// <summary>Ajoute ou actualise une étape native, sans dupliquer son identité.</summary>
        /// <param name="activity">Données reçues du fournisseur.</param>
        private void ReceiveAgentActivity(CodexAgentActivity activity)
        {
            if (IsDisposed || runtimeDisposed || conversationItems == null || activity == null || string.IsNullOrEmpty(activity.Id)) return;
            if (activity.Kind == "reasoning" && string.IsNullOrWhiteSpace(activity.Detail) && !liveEntries.ContainsKey(activity.Id)) return;
            bool isNew = !liveEntries.TryGetValue(activity.Id, out var entry);
            if (isNew)
            {
                entry = new ChatEntry { Speaker = activity.Kind == "reasoning" ? "Réflexion" : "Outil", StreamId = activity.Id };
                liveEntries[activity.Id] = entry;
                entry.Activity = new CodexAgentActivity { Id = activity.Id, Kind = activity.Kind };
            }
            var previous = entry.Activity;
            entry.Activity = new CodexAgentActivity { Id = activity.Id, Kind = activity.Kind,
                Title = string.IsNullOrWhiteSpace(activity.Title) ? previous?.Title : activity.Title,
                Detail = activity.Append ? CodexAgentActivity.Limit((previous?.Detail ?? "") + activity.Detail) : CodexAgentActivity.Limit(activity.Detail),
                Status = activity.Append && previous != null && previous.Status != "inProgress" ? previous.Status : activity.Status, DurationMs = activity.DurationMs ?? previous?.DurationMs };
            entry.Text = entry.Activity.Title + "\n" + entry.Activity.Detail;
            if (activity.Status != "inProgress") completedStreams.Add(activity.Id);
            // Publish a complete first snapshot; adding to a group already refreshes its view.
            if (isNew) AddEntry(entry);
            bool textOnly = activity.Append && previous != null && previous.Kind == entry.Activity.Kind &&
                previous.Title == entry.Activity.Title && previous.Status == entry.Activity.Status && previous.DurationMs == entry.Activity.DurationMs;
            if (textOnly && liveTexts.TryGetValue(activity.Id, out var live) && !live.IsDisposed)
            {
                pendingActivityText.Add(entry);
                ScheduleStreamRender();
            }
            else
            {
                pendingActivityText.Remove(entry);
                if (!isNew) RefreshVisibleActivity(entry);
                FollowLatest();
            }
            ScheduleSessionSave();
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
            if (entryViews.TryGetValue(entry, out var view) && view is ChatDesignerHost host &&
                host.View is ChatActivityGroupView group && !group.IsDisposed &&
                activityGroups.TryGetValue(entry, out var entries) && UpdateActivityGroup(group, entry, entries))
                return;
            int index = visibleEntries.IndexOf(entry);
            if (index >= 0) { visibleEntries.RemoveAt(index); visibleEntries.Insert(index, entry); }
        }

        /// <summary>Construit un unique bloc repliable contenant les activités et résumés dans leur ordre d'origine.</summary>
        /// <param name="owner">Première entrée servant d'identité stable au bloc.</param>
        /// <param name="entries">Activités à afficher.</param>
        /// <returns>Bloc repliable, fermé par défaut et conservant son état lors des mises à jour.</returns>
        private FrameworkElement RenderActivityGroup(ChatEntry owner, List<ChatEntry> entries)
        {
            var card = new ChatActivityGroupView();
            var state = new ActivityGroupViewState();
            card.Tag = state;
            UpdateActivityGroup(card, owner, entries);
            card.section.ExpansionChanged += (sender, args) => {
                if (state.Updating) return;
                if (card.section.Expanded) { expandedActivityGroups.Add(owner); collapsedActivityGroups.Remove(owner); }
                else { expandedActivityGroups.Remove(owner); collapsedActivityGroups.Add(owner); }
            };
            return new ChatDesignerHost(card) { Margin = new Thickness(0,4,0,14) };
        }

        /// <summary>Updates realized rows without replacing their native text controls or user selection.</summary>
        private bool UpdateActivityGroup(ChatActivityGroupView card, ChatEntry owner, List<ChatEntry> entries)
        {
            var state = card.Tag as ActivityGroupViewState;
            if (state == null || entries.Count == 0) return false;
            // A changed row type or damaged control requires the existing full rebuild path.
            if (state.Rows.Count > entries.Count) return false;
            int retainedRows = 0;
            foreach (var entry in entries)
            {
                if (!state.Rows.TryGetValue(entry, out var existingRow)) continue;
                retainedRows++;
                if (existingRow.IsDisposed) return false;
                if (entry.Activity != null)
                {
                    if (!(existingRow is ChatActivityStepView existingStep) || existingStep.detail.content.IsDisposed) return false;
                }
                else if (!(existingRow is ChatTextContentView existingText) || existingText.content.IsDisposed) return false;
            }
            if (retainedRows != state.Rows.Count) return false;
            state.Updating = true;
            card.SuspendLayout();
            card.section.body.SuspendLayout();
            try
            {
                foreach (var entry in entries)
                {
                    if (!state.Rows.TryGetValue(entry, out var row))
                    {
                        row = entry.Activity != null ? (System.Windows.Forms.Control)CreateActivityStep(entry) : SelectableText(entry.Text);
                        state.Rows.Add(entry, row);
                        card.section.body.Controls.Add(row);
                    }
                    else if (row is ChatActivityStepView step) UpdateActivityStep(step, entry);
                    else if (row is ChatTextContentView text) SetTranscriptText(text.content, entry.Text);
                    if (!string.IsNullOrEmpty(entry.StreamId))
                        liveTexts[entry.StreamId] = row is ChatActivityStepView activity ? activity.detail.content : ((ChatTextContentView)row).content;
                }
                var latest = entries.LastOrDefault(entry => entry.Activity?.Status == "inProgress") ?? entries.Last();
                string preview = string.IsNullOrWhiteSpace(latest.Activity?.Title) ? "" : " · " + CodexAgentActivity.Limit(latest.Activity.Title);
                if (preview.Length > 90) preview = preview.Substring(0,87) + "…";
                string title = UiText.Get(entries.All(e => e.Speaker == "Réflexion") ? "Reasoning" : "Agent activity") + " · " + entries.Count + preview;
                if (card.section.Title != title) card.section.Title = title;
                bool running = entries.Any(e => e.Activity?.Status == "inProgress" || (busy && e.Activity == null && e.StreamId != null && !completedStreams.Contains(e.StreamId)));
                card.section.Expanded = expandedActivityGroups.Contains(owner) || running && !collapsedActivityGroups.Contains(owner);
            }
            finally
            {
                card.section.body.ResumeLayout(true);
                card.ResumeLayout(true);
                state.Updating = false;
            }
            return true;
        }
        /// <summary>Dessine une étape compacte avec résultat, durée native et détail dépliable.</summary>
        /// <param name="entry">Entrée enrichie de l'historique.</param>
        /// <returns>Étape de la chronologie.</returns>
        private FrameworkElement RenderActivityStep(ChatEntry entry) => new ChatDesignerHost(CreateActivityStep(entry));
        /// <summary>Creates a transcript row for a tool activity and its displayed state.</summary>
        /// <param name="entry">The entry used by this operation.</param>
        /// <returns>The result produced by this operation.</returns>
        private ChatActivityStepView CreateActivityStep(ChatEntry entry)
        {
            var card = new ChatActivityStepView { Tag = new ActivityStepViewState() };
            UpdateActivityStep(card, entry);
            card.section.ExpansionChanged += (sender, args) => {
                if (((ActivityStepViewState)card.Tag).Updating) return;
                if (card.section.Expanded) { expandedActivitySteps.Add(entry); collapsedActivitySteps.Remove(entry); }
                else { expandedActivitySteps.Remove(entry); collapsedActivitySteps.Add(entry); }
            };
            return card;
        }

        private void UpdateActivityStep(ChatActivityStepView card, ChatEntry entry)
        {
            var activity = entry.Activity;
            var state = (ActivityStepViewState)card.Tag;
            state.Updating = true;
            try
            {
                bool running = activity.Status == "inProgress", failed = activity.Status == "failed";
                string status = UiText.Get(running ? "In progress" : failed ? "Failed" : activity.Status == "declined" ? "Declined" : activity.Status == "completed" ? "Completed" : activity.Status == "interrupted" ? "Interrupted" : "Cancelled");
                if (activity.DurationMs.HasValue) status += " · " + (activity.DurationMs.Value / 1000d).ToString("0.0", UiText.Culture) + " s";
                if (card.state.Text != status) card.state.Text = status;
                string title = string.IsNullOrWhiteSpace(activity.Title) ? UiText.Get(activity.Kind == "reasoning" ? "Reasoning" : "Tool") : activity.Title;
                if (card.section.Title != title) card.section.Title = title;
                if (!state.Initialized || state.Kind != activity.Kind)
                {
                    int start = card.detail.content.SelectionStart, length = card.detail.content.SelectionLength;
                    card.detail.ShowPlain(activity.Detail, activity.Kind == "commandExecution");
                    card.detail.content.Select(System.Math.Min(start, card.detail.content.TextLength), System.Math.Min(length, System.Math.Max(0, card.detail.content.TextLength - start)));
                    state.Kind = activity.Kind; state.Initialized = true;
                }
                else SetTranscriptText(card.detail.content, activity.Detail);
                if (!string.IsNullOrEmpty(entry.StreamId)) liveTexts[entry.StreamId] = card.detail.content;
                card.section.Expanded = expandedActivitySteps.Contains(entry) || running && activity.Kind == "reasoning" && !collapsedActivitySteps.Contains(entry);
            }
            finally { state.Updating = false; }
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
