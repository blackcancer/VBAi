using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Threading;

namespace CodexVBE
{
    /// <summary>Fenêtre de conversation qui affiche les messages et les modifications de code.</summary>
    internal sealed partial class ChatWindow
    {
        /// <summary>Défileur de la conversation une fois matérialisé par le template WPF.</summary>
        private ScrollViewer conversationScroll;
        /// <summary>Liste virtualisée des éléments affichés dans la conversation.</summary>
        private ListBox conversationItems;
        /// <summary>Entrées actuellement visibles par la liste virtualisée.</summary>
        private readonly System.Collections.ObjectModel.ObservableCollection<object> visibleEntries = new System.Collections.ObjectModel.ObservableCollection<object>();
        /// <summary>Index de la première entrée du transcript chargée dans la fenêtre visible.</summary>
        private int firstLoadedEntry;
        /// <summary>Marqueur inséré pour charger les messages antérieurs.</summary>
        private readonly object earlierEntries = new object();
        /// <summary>Indique si le défilement doit rester attaché au dernier message.</summary>
        private bool followConversation = true;
        /// <summary>Historique complet des entrées de la session courante.</summary>
        private readonly List<ChatEntry> transcriptEntries = new List<ChatEntry>();
        /// <summary>Contrôles matérialisés actuellement associés à leurs entrées.</summary>
        private readonly Dictionary<ChatEntry, FrameworkElement> entryViews = new Dictionary<ChatEntry, FrameworkElement>();
        /// <summary>Messages de flux actifs indexés par leur identifiant.</summary>
        private readonly Dictionary<string, ChatEntry> liveEntries = new Dictionary<string, ChatEntry>();
        /// <summary>Champs texte matérialisés pour afficher le texte des flux actifs.</summary>

        private readonly Dictionary<string, System.Windows.Forms.RichTextBox> liveTexts = new Dictionary<string, System.Windows.Forms.RichTextBox>();
        /// <summary>Boutons de restauration associés aux changements de code visibles.</summary>

        private readonly Dictionary<CodeChange, System.Windows.Forms.Button> rollbackButtons = new Dictionary<CodeChange, System.Windows.Forms.Button>();
        /// <summary>Libellés d’état associés aux changements de code visibles.</summary>

        private readonly Dictionary<CodeChange, System.Windows.Forms.Label> changeStates = new Dictionary<CodeChange, System.Windows.Forms.Label>();

        /// <summary>Configure la liste virtualisée, les événements de défilement, l’accessibilité et les changements de thème.</summary>
        private void InitializeTranscript()
        {
            conversationItems = new ListBox { ItemsSource = visibleEntries, Background = Ink("#F8FAFC"), BorderThickness = new Thickness(0),
                FlowDirection = UiText.Culture.TextInfo.IsRightToLeft ? FlowDirection.RightToLeft : FlowDirection.LeftToRight };
            ScrollViewer.SetCanContentScroll(conversationItems, true);
            ScrollViewer.SetHorizontalScrollBarVisibility(conversationItems, ScrollBarVisibility.Disabled);
            VirtualizingPanel.SetIsVirtualizing(conversationItems, true);
            VirtualizingPanel.SetVirtualizationMode(conversationItems, VirtualizationMode.Recycling);
            VirtualizingPanel.SetScrollUnit(conversationItems, ScrollUnit.Pixel);
            var items = new FrameworkElementFactory(typeof(VirtualizingStackPanel));
            conversationItems.ItemsPanel = new ItemsPanelTemplate(items);
            var presenter = new FrameworkElementFactory(typeof(TranscriptItem));
            presenter.SetValue(TranscriptItem.RenderProperty, new Action<TranscriptItem>(RealizeEntry));
            presenter.SetValue(TranscriptItem.ReleaseProperty, new Action<TranscriptItem>(ReleaseEntry));
            conversationItems.ItemTemplate = new DataTemplate { VisualTree = presenter };
            var style = new Style(typeof(ListBoxItem));
            style.Setters.Add(new Setter(Control.HorizontalContentAlignmentProperty, HorizontalAlignment.Stretch));
            style.Setters.Add(new Setter(Control.PaddingProperty, new Thickness(16, 4, 16, 0)));
            style.Setters.Add(new Setter(Control.BackgroundProperty, Brushes.Transparent));
            conversationItems.ItemContainerStyle = style;
            conversationItems.AddHandler(ScrollViewer.ScrollChangedEvent, new ScrollChangedEventHandler((s, e) => {
                var scroller = e.OriginalSource as ScrollViewer;
                if (scroller == null || !ReferenceEquals(scroller.TemplatedParent, conversationItems)) return;
                conversationScroll = scroller;
                if (Math.Abs(e.ExtentHeightChange) < 0.1 && Math.Abs(e.VerticalChange) > 0.1)
                    followConversation = conversationScroll.ScrollableHeight - conversationScroll.VerticalOffset < 32;
                jumpToLatest.Visible = !followConversation;
            }));
            AutomationProperties.SetName(conversationItems, UiText.Get("Conversation"));
            transcriptHost.Child = conversationItems;

            Action themeChanged = () => { if (!IsDisposed && IsHandleCreated) BeginInvoke(new Action(() => { conversationItems.Background = Ink("#F8FAFC"); prompt.Foreground = Ink("#1E293B"); referenceList.BackColor = UiTheme.Surface; UiTheme.Apply(referenceView); RefreshTranscriptWindow(firstLoadedEntry); ShowWelcome(); })); };
            UiTheme.Changed += themeChanged;
            Disposed += (s, e) => UiTheme.Changed -= themeChanged;
        }
        /// <summary>Construit le contrôle visuel d’un élément lorsque le panneau virtualisé le matérialise.</summary>
        /// <param name="item">Élément du transcript à matérialiser.</param>
        private void RealizeEntry(TranscriptItem item)
        {
            var entry = item.DataContext as ChatEntry;
            if (entry != null) { var view = RenderEntry(entry); entryViews[entry] = view; item.Content = view; RefreshCodeChangeCards(); }
            else if (ReferenceEquals(item.DataContext, earlierEntries))
            {

                var more = new ChatLinkView(); more.link.Text = UiText.Get("Load earlier messages");

                more.link.Click += (s, e) => {
                    var anchor = visibleEntries.OfType<ChatEntry>().FirstOrDefault();
                    RefreshTranscriptWindow(Math.Max(0, firstLoadedEntry - 80));
                    if (anchor != null) conversationItems.ScrollIntoView(anchor);
                };

                item.Content = new ChatDesignerHost(more);
            }
            else item.Content = item.DataContext as FrameworkElement;
        }
        /// <summary>Retire les références aux contrôles temporaires lorsqu’un élément sort de la fenêtre virtualisée.</summary>
        /// <param name="item">Élément du transcript libéré.</param>
        private void ReleaseEntry(TranscriptItem item)
        {
            if ((item.RenderedContext is ChatEntry || ReferenceEquals(item.RenderedContext, earlierEntries)) && item.Content is FrameworkElement content) DisposeEntryView(content);
            if (item.RenderedContext is ChatEntry entry && entryViews.TryGetValue(entry, out var view) && ReferenceEquals(view, item.Content)) {
                entryViews.Remove(entry);
                ReleaseActivityTexts(entry);
                if (entry.StreamId != null) liveTexts.Remove(entry.StreamId);
                if (entry.FormCut != null) formCutButtons.Remove(entry.FormCut);
                if (entry.Change != null) { rollbackButtons.Remove(entry.Change); changeStates.Remove(entry.Change); }
            }
        }
        /// <summary>Libère les hôtes natifs avant de remplacer les vues matérialisées.</summary>
        private void DisposeEntryViews()
        {
            foreach (var view in entryViews.Values.ToArray()) DisposeEntryView(view);

            foreach (var view in visibleEntries.OfType<ChatDesignerHost>().ToArray()) view.Dispose();
        }
                /// <summary>Libère récursivement les ressources WPF détenues par une entrée et ses enfants visuels/logiques.</summary>
        /// <param name="view">Racine de l’élément de transcript à nettoyer.</param>
private static void DisposeEntryView(FrameworkElement view)
        {

            if (view is ChatDesignerHost card) { card.Dispose(); return; }
            if (view is ChatDiffView diff) { diff.Dispose(); return; }
            foreach (var child in LogicalTreeHelper.GetChildren(view).OfType<FrameworkElement>()) DisposeEntryView(child);
        }
        /// <summary>Remplace la fenêtre virtualisée par les entrées commençant à l’index demandé.</summary>
        /// <param name="start">Index de départ dans le transcript complet.</param>
        private void RefreshTranscriptWindow(int start)
        {
            DisposeEntryViews();
            while (start > 0 && start < transcriptEntries.Count &&
                CanGroupActivities(transcriptEntries[start - 1], transcriptEntries[start])) start--;
            firstLoadedEntry = start;
            visibleEntries.Clear(); entryViews.Clear(); liveTexts.Clear(); rollbackButtons.Clear(); changeStates.Clear(); formCutButtons.Clear();
            if (start > 0) visibleEntries.Add(earlierEntries);
            activityGroups.Clear(); activityOwners.Clear();
            ChatEntry previous = null;
            foreach (var entry in transcriptEntries.Skip(start)) { AppendVisibleEntry(entry, previous); previous = entry; }
        }
        /// <summary>Efface le transcript complet et réinitialise les contrôles matérialisés et le suivi du défilement.</summary>
        private void ClearTranscript()
        {
            DisposeEntryViews();
            visibleEntries.Clear(); firstLoadedEntry = 0;
            transcriptEntries.Clear(); entryViews.Clear(); liveEntries.Clear(); liveTexts.Clear();
            activityGroups.Clear(); activityOwners.Clear(); expandedActivityGroups.Clear(); expandedActivitySteps.Clear();
            rollbackButtons.Clear(); changeStates.Clear(); formCutButtons.Clear(); followConversation = true;
        }
        /// <summary>Fait défiler vers le dernier élément si le suivi automatique est activé.</summary>
        private void FollowLatest()
        {
            if (!followConversation || conversationItems == null || visibleEntries.Count == 0) return;
            conversationItems.Dispatcher.BeginInvoke(DispatcherPriority.Background, new Action(() => {
                if (followConversation && visibleEntries.Count > 0) conversationItems.ScrollIntoView(visibleEntries.Last());
            }));
        }
        /// <summary>Crée un champ texte en lecture seule dont le contenu peut être sélectionné et copié.</summary>
        /// <param name="text">Texte à afficher.</param>
        /// <param name="code">Active une police monospace, une ligne non renvoyée et un flux gauche-droite.</param>
        /// <returns>Vue WinForms Designer configurée pour la sélection du texte.</returns>

        private ChatTextContentView SelectableText(string text, bool code = false)
        {

            var view = new ChatTextContentView { ErrorHandler = SetStatus }; view.ShowPlain(text,code); return view;
        }
        /// <summary>Ajoute un message simple au transcript.</summary>
        /// <param name="speaker">Locuteur ou catégorie du message.</param>

        /// <param name="content">Message text.</param>
        private void AddTranscriptMessage(string speaker, string content) { AddEntry(new ChatEntry { Speaker = speaker, Text = content }); }
        /// <summary>Ajoute une entrée à l’historique et, hors chargement de session, à la liste visible.</summary>
        /// <param name="entry">Entrée à ajouter.</param>
        private void AddEntry(ChatEntry entry)
        {
            if (conversationItems == null) return;
            if (entry.TurnId == null) entry.TurnId = activeTurnId;
            if (transcriptEntries.Count == 0) visibleEntries.Clear();
            var previous = transcriptEntries.LastOrDefault();
            transcriptEntries.Add(entry);
            if (!loadingSession) { AppendVisibleEntry(entry, previous); FollowLatest(); ScheduleSessionSave(); }
        }

        /// <summary>Construit le contrôle WPF correspondant à un message, une activité, une référence ou une pièce jointe.</summary>
        /// <param name="entry">Entrée du transcript à afficher.</param>
        /// <returns>Élément WPF matérialisant l’entrée.</returns>
        private FrameworkElement RenderEntry(ChatEntry entry)
        {
            if (activityGroups.TryGetValue(entry, out var activities)) return RenderActivityGroup(entry, activities);
            if (entry.FormCut != null) return RenderFormCut(entry.FormCut);
            if (entry.Change != null) return RenderChange(entry.Change);

            if (IsActivity(entry)) return RenderActivityGroup(entry, new List<ChatEntry> { entry });

            var card = new ChatMessageView();

            card.speaker.Text = entry.Speaker == "Vous" ? UiText.Get("YOU") : UiText.Speaker(entry.Speaker).ToUpperInvariant();

            card.copy.Click += (s,e) => CopyText(entry.Text);

            card.fork.Visible = entry.Speaker == "Vous" || entry.Speaker == "Assistant";

            card.fork.Click += (s,e) => ForkChat(entry);

            if (!string.IsNullOrEmpty(entry.StreamId) && !completedStreams.Contains(entry.StreamId)) {

                card.message.ShowPlain(entry.Text); liveTexts[entry.StreamId] = card.message.content;

            } else {

                var refs = transcriptEntries.SelectMany(x => x.References ?? new VbeChatReference[0]).GroupBy(x => x.Token).ToDictionary(x => x.Key, x => x.Last());

                card.message.ShowMarkdown(entry.Text ?? "", refs, NavigateReference, SetStatus);
            }

            card.memory.Visible = !string.IsNullOrWhiteSpace(entry.AttachedMemory);

            if (card.memory.Visible) {

                var text = SelectableText(entry.AttachedMemory); card.memory.body.Controls.Add(text);
            }

            foreach (var attachment in entry.Attachments ?? new ChatAttachment[0]) {

                var row = new ChatAttachmentView();

                row.section.Title = attachment.Label + " · " + attachment.Text.Length + UiText.Get(" characters");

                row.text.ShowPlain(attachment.Text, !string.IsNullOrEmpty(attachment.Module)); row.open.Visible = !string.IsNullOrEmpty(attachment.Module);

                row.open.Click += (s,e) => NavigateAttachment(attachment); card.attachments.Controls.Add(row);
            }

            foreach (var reference in entry.References ?? new VbeChatReference[0]) {

                var row = new ChatLinkView(); row.link.Text = reference.Token; row.link.Click += (s,e) => NavigateReference(reference);

                card.references.Controls.Add(row);
            }

            var targets = entry.Speaker == "Intervention" ? codeChanges.Where(x => x.TurnId == entry.TurnId).ToArray() : new CodeChange[0];

            foreach (var target in targets) {

                var row = new ChatLinkView(); row.link.Text = target.Label; row.link.Click += (s,e) => ShowCodeChanges(target); card.targets.Controls.Add(row);
            }

            card.undoTurn.Visible = targets.Length > 0;

            card.undoTurn.Click += (s,e) => { if (targets.Length > 0) RollbackIntervention(targets[0], null, true); };

            card.fix.Visible = entry.Speaker == "Vérification" && entry.Attachments?.Length > 0;

            card.fix.Click += (s,e) => { if (busy) return; modePicker.SelectedItem = ChatMode.Agent; prompt.Text = "/corriger " + entry.Text; draftAttachments.AddRange(entry.Attachments); RefreshContextChips(); };

            return new ChatDesignerHost(card) { Margin = new Thickness(entry.Speaker == "Vous" ? 30 : 0, 0, 4, 14) };
        }

        /// <summary>Copie le texte dans le presse-papiers et signale les erreurs à l’interface.</summary>
        /// <param name="text">Texte à copier.</param>
        private void CopyText(string text)
        {
            try { WriteClipboard(text ?? ""); }
            catch (Exception ex) { SetStatus(UiText.Get("Unable to copy: ") + ex.Message); }
        }

        /// <summary>Ajoute une entrée de changement de code puis actualise les cartes visibles.</summary>
        /// <param name="change">Changement à afficher.</param>
        private void AddCodeChangeCard(CodeChange change)
        {
            AddEntry(new ChatEntry { Speaker = "Code", Change = change });
            RefreshCodeChangeCards();
        }

        /// <summary>Construit une carte de diff avec navigation vers le module et actions de restauration.</summary>
        /// <param name="change">Changement dont il faut afficher le diff et les actions.</param>
        /// <returns>Carte WPF de changement.</returns>
        private FrameworkElement RenderChange(CodeChange change)
        {

            var card = new ChatChangeCardView();

            card.module.Text = "#" + change.Project + "." + change.Module;

            card.module.Click += (s,e) => NavigateReference(new VbeChatReference { Project = change.Project, Module = change.Module });

            card.count.Text = "+" + change.Rows.Count(r => r.Kind == CodeDiffKind.Added) + "  −" + change.Rows.Count(r => r.Kind == CodeDiffKind.Removed);

            card.diff.ShowDiff(change.Before, change.After); card.section.Expanded = true;

            card.undo.Click += async (s,e) => {
                if (busy || tools == null) return;

                try { EnsureCurrentScope(); } catch (Exception ex) { SetStatus(ex.Message); return; }
                var result = await tools.RestoreChangesAsync(new[] { change }, null);
                if (!result.Ok) AddTranscriptMessage("Erreur", UiText.Get("Unable to undo: ") + result.Error);
                RefreshCodeChangeCards(); SaveCurrentSession();
            };

            rollbackButtons[change] = card.undo; changeStates[change] = card.state;

            card.blocks.Click += (s,e) => {

                card.blockMenu.Items.Clear();
                foreach (var hunk in CodeRollback.Hunks(change.Before, change.After).Where(x => !change.RestoredHunks.Contains(x.Index))) {

                    var item = new System.Windows.Forms.ToolStripMenuItem(UiText.Get("Block ") + (hunk.Index + 1) + " · L" + (hunk.AfterStart + 1) + " · +" + hunk.After.Length + " −" + hunk.Before.Length) { Enabled = !busy };

                    item.Click += (a,b) => RollbackIntervention(change, hunk.Index, false); card.blockMenu.Items.Add(item);
                }

                card.blockMenu.Show(card.blocks, 0, card.blocks.Height);

            };

            card.undoTurn.Visible = !string.IsNullOrEmpty(change.TurnId);

            card.undoTurn.Click += (s,e) => RollbackIntervention(change, null, true);

            var host = new ChatDesignerHost(card) { Margin = new Thickness(0, 0, 4, 14) };

            RefreshCodeChangeCards(); return host;
        }

        /// <summary>Actualise l’activation et le libellé des boutons selon l’état restauré et l’activité courante.</summary>
        private void RefreshCodeChangeCards()
        {
            RefreshFormCutCards();
            foreach (var pair in rollbackButtons)
            {

                pair.Value.Enabled = !pair.Key.Restored && !busy;

                pair.Value.Text = pair.Key.Restored ? UiText.Get("Change undone") : UiText.Get("Undo change");
                changeStates[pair.Key].Text = pair.Key.Restored ? UiText.Get("Code restored") : pair.Key.RestoredHunks.Count > 0 ? UiText.Get("Partially undone") : UiText.Get("Applied");
            }
        }

        /// <summary>Identifiants des flux dont le contenu final est disponible.</summary>
        private readonly HashSet<string> completedStreams = new HashSet<string>();
        /// <summary>Texte final du flux assistant courant, utilisé pour éviter un message final en double.</summary>
        private string streamedFinalText;

        /// <summary>Ajoute ou met à jour le texte d’un flux de réponse et finalise l’entrée lorsqu’il est terminé.</summary>
        /// <param name="kind">Catégorie de flux, notamment summary, tool ou final.</param>
        /// <param name="id">Identifiant stable du flux.</param>
        /// <param name="text">Fragment reçu ou texte final.</param>
        /// <param name="complete">Indique si la mise à jour termine le flux.</param>
        private void ReceiveChatUpdate(string kind, string id, string text, bool complete)
        {
            if (IsDisposed || conversationItems == null) return;
            if (kind == "summary" && !liveEntries.ContainsKey(id) && string.IsNullOrWhiteSpace(text)) return;
            ChatEntry entry;
            if (!liveEntries.TryGetValue(id, out entry))
            {
                entry = new ChatEntry { Speaker = kind == "summary" ? "Réflexion" :
                    kind == "tool" ? "Outil" : "Assistant", StreamId = id, Text = "" };
                liveEntries[id] = entry;
                AddEntry(entry);
            }
            entry.Text = complete ? (string.IsNullOrWhiteSpace(text) ? entry.Text : text) : entry.Text + (text ?? "");
            if (complete)
            {
                completedStreams.Add(id);
                if (kind == "final") streamedFinalText = entry.Text;
                RefreshVisibleActivity(entry);
            }
            else
            {

                System.Windows.Forms.RichTextBox live;
                if (liveTexts.TryGetValue(id, out live)) live.Text = entry.Text;
            }
            FollowLatest(); ScheduleSessionSave();
        }

        /// <summary>Ajoute le texte final de l’assistant s’il n’a pas déjà été affiché par le flux.</summary>
        /// <param name="text">Réponse complète de l’assistant.</param>
        private void CompleteAssistantResponse(string text)
        {
            if (!string.Equals(streamedFinalText, text, StringComparison.Ordinal)) AddTranscriptMessage("Assistant", text);
            streamedFinalText = null;
        }

        /// <summary>Affiche l’écran d’accueil avec suggestions tant que le transcript est vide.</summary>
        private void ShowWelcome()
        {
            if (conversationItems == null || transcriptEntries.Count > 0) return;

            var welcome = new ChatWelcomeView();

            foreach (var button in new[] { welcome.explain, welcome.fix, welcome.improve }) {

                string seed = button.Text;

                button.Click += (s,e) => { prompt.Text = seed + " "; prompt.CaretIndex = prompt.Text.Length; prompt.Focus(); };
            }

            visibleEntries.Add(new ChatDesignerHost(welcome));
        }
    }
}
