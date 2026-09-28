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
        private readonly Dictionary<string, TextBox> liveTexts = new Dictionary<string, TextBox>();
        /// <summary>Boutons de restauration associés aux changements de code visibles.</summary>
        private readonly Dictionary<CodeChange, Button> rollbackButtons = new Dictionary<CodeChange, Button>();
        /// <summary>Libellés d’état associés aux changements de code visibles.</summary>
        private readonly Dictionary<CodeChange, TextBlock> changeStates = new Dictionary<CodeChange, TextBlock>();

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
            Action themeChanged = () => { if (!IsDisposed && IsHandleCreated) BeginInvoke(new Action(() => { conversationItems.Background = Ink("#F8FAFC"); prompt.Foreground = Ink("#1E293B"); referenceList.Background = Ink("#FFFFFF"); ((Border)referencePopup.Child).Background = Ink("#FFFFFF"); RefreshTranscriptWindow(firstLoadedEntry); ShowWelcome(); })); };
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
                var more = ChatButton(UiText.Get("Load earlier messages"));
                more.Click += (s, e) => {
                    var anchor = visibleEntries.OfType<ChatEntry>().FirstOrDefault();
                    RefreshTranscriptWindow(Math.Max(0, firstLoadedEntry - 80));
                    if (anchor != null) conversationItems.ScrollIntoView(anchor);
                };
                item.Content = more;
            }
            else item.Content = item.DataContext as FrameworkElement;
        }
        /// <summary>Retire les références aux contrôles temporaires lorsqu’un élément sort de la fenêtre virtualisée.</summary>
        /// <param name="item">Élément du transcript libéré.</param>
        private void ReleaseEntry(TranscriptItem item)
        {
            if (item.RenderedContext is ChatEntry entry && entryViews.TryGetValue(entry, out var view) && ReferenceEquals(view, item.Content)) {
                entryViews.Remove(entry);
                if (entry.StreamId != null) liveTexts.Remove(entry.StreamId);
                if (entry.FormCut != null) formCutButtons.Remove(entry.FormCut);
                if (entry.Change != null) { rollbackButtons.Remove(entry.Change); changeStates.Remove(entry.Change); }
            }
        }
        /// <summary>Remplace la fenêtre virtualisée par les entrées commençant à l’index demandé.</summary>
        /// <param name="start">Index de départ dans le transcript complet.</param>
        private void RefreshTranscriptWindow(int start)
        {
            firstLoadedEntry = start;
            visibleEntries.Clear(); entryViews.Clear(); liveTexts.Clear(); rollbackButtons.Clear(); changeStates.Clear(); formCutButtons.Clear();
            if (start > 0) visibleEntries.Add(earlierEntries);
            foreach (var entry in transcriptEntries.Skip(start)) visibleEntries.Add(entry);
        }
        /// <summary>Efface le transcript complet et réinitialise les contrôles matérialisés et le suivi du défilement.</summary>
        private void ClearTranscript()
        {
            visibleEntries.Clear(); firstLoadedEntry = 0;
            transcriptEntries.Clear(); entryViews.Clear(); liveEntries.Clear(); liveTexts.Clear();
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
        /// <returns>Champ WPF configuré pour la sélection du texte.</returns>
        private static TextBox SelectableText(string text, bool code = false)
        {
            return new TextBox { Text = text ?? "", IsReadOnly = true, AcceptsReturn = true,
                FlowDirection = !code && UiText.Culture.TextInfo.IsRightToLeft ? FlowDirection.RightToLeft : FlowDirection.LeftToRight,
                TextWrapping = code ? TextWrapping.NoWrap : TextWrapping.Wrap,
                BorderThickness = new Thickness(0), Background = Brushes.Transparent,
                Foreground = Ink("#334155"), FontSize = 13,
                FontFamily = new FontFamily(code ? "Consolas" : "Segoe UI"),
                Padding = new Thickness(0), IsReadOnlyCaretVisible = true };
        }
        /// <summary>Ajoute un message simple au transcript.</summary>
        /// <param name="speaker">Locuteur ou catégorie du message.</param>
        /// <param name="content">Texte du message.</param>
        private void AddTranscriptMessage(string speaker, string content) { AddEntry(new ChatEntry { Speaker = speaker, Text = content }); }
        /// <summary>Ajoute une entrée à l’historique et, hors chargement de session, à la liste visible.</summary>
        /// <param name="entry">Entrée à ajouter.</param>
        private void AddEntry(ChatEntry entry)
        {
            if (conversationItems == null) return;
            if (entry.TurnId == null) entry.TurnId = activeTurnId;
            if (transcriptEntries.Count == 0) visibleEntries.Clear();
            transcriptEntries.Add(entry);
            if (!loadingSession) { visibleEntries.Add(entry); FollowLatest(); ScheduleSessionSave(); }
        }

        /// <summary>Construit le contrôle WPF correspondant à un message, une activité, une référence ou une pièce jointe.</summary>
        /// <param name="entry">Entrée du transcript à afficher.</param>
        /// <returns>Élément WPF matérialisant l’entrée.</returns>
        private FrameworkElement RenderEntry(ChatEntry entry)
        {
            if (entry.FormCut != null) return RenderFormCut(entry.FormCut);
            if (entry.Change != null) return RenderChange(entry.Change);
            if (entry.Speaker == "Réflexion" || entry.Speaker == "Outil")
            {
                var text = SelectableText(entry.Text);
                text.Margin = new Thickness(18, 8, 4, 6);
                if (!string.IsNullOrEmpty(entry.StreamId)) liveTexts[entry.StreamId] = text;
                return new Expander { Header = entry.Speaker == "Réflexion" ? UiText.Get("Reasoning · summary") : UiText.Get("Agent activity"),
                    IsExpanded = entry.Speaker == "Réflexion", Content = text, FontSize = 12,
                    Foreground = Ink("#64748B"), Margin = new Thickness(0, 4, 0, 14) };
            }
            bool user = entry.Speaker == "Vous";
            var body = new StackPanel();
            var heading = new DockPanel { Margin = new Thickness(0, 0, 0, 8) };
            var copy = ChatButton(UiText.Get("Copy"));
            copy.ToolTip = UiText.Get("Copy the message text to the clipboard.");
            copy.FontSize = 10; copy.MinHeight = 20; copy.Padding = new Thickness(6, 2, 6, 2);
            copy.Background = Brushes.Transparent;
            copy.Click += (s, e) => CopyText(entry.Text);
            DockPanel.SetDock(copy, System.Windows.Controls.Dock.Right); heading.Children.Add(copy);
            if (entry.Speaker == "Vous" || entry.Speaker == "Assistant") {
                var fork = ChatButton(UiText.Get("Branch conversation")); fork.FontSize = 10; fork.Padding = new Thickness(6, 2, 6, 2);
                fork.ToolTip = UiText.Get("Create an independent conversation with the history up to this message.");
                fork.Click += (s, e) => ForkChat(entry); DockPanel.SetDock(fork, System.Windows.Controls.Dock.Right); heading.Children.Add(fork);
            }
            heading.Children.Add(new TextBlock { Text = user ? UiText.Get("YOU") : UiText.Speaker(entry.Speaker).ToUpperInvariant(),
                FontSize = 10, FontWeight = FontWeights.SemiBold, Foreground = Ink("#64748B"),
                VerticalAlignment = VerticalAlignment.Center });
            body.Children.Add(heading);
            if (!string.IsNullOrEmpty(entry.StreamId) && !completedStreams.Contains(entry.StreamId))
            {
                var live = SelectableText(entry.Text);
                liveTexts[entry.StreamId] = live;
                body.Children.Add(live);
            }
            else RenderMarkdown(body, entry.Text ?? "");
            if (entry.Speaker == "Intervention")
            {
                var targets = codeChanges.Where(x => x.TurnId == entry.TurnId).ToArray();
                foreach (var target in targets) { var link = ChatButton(target.Label); link.Click += (s, e) => ShowCodeChanges(target); body.Children.Add(link); }
                if (targets.Length > 0) { var undo = ChatButton(UiText.Get("Undo entire turn")); undo.ToolTip = UiText.Get("Undo changes from this turn after checking for conflicts."); undo.Click += (s, e) => RollbackIntervention(targets[0], null, true); body.Children.Add(undo); }
            }
            if (!string.IsNullOrWhiteSpace(entry.AttachedMemory))
                body.Children.Add(new Expander { Header = UiText.Get("Attached document memory"),
                    Content = SelectableText(entry.AttachedMemory), FontSize = 11, Margin = new Thickness(0, 8, 0, 0) });
            foreach (var attachment in entry.Attachments ?? new ChatAttachment[0]) {
                var content = new StackPanel();
                var attachmentText = SelectableText(attachment.Text);
                if (!string.IsNullOrEmpty(attachment.Module)) attachmentText.FlowDirection = FlowDirection.LeftToRight;
                content.Children.Add(attachmentText);
                if (!string.IsNullOrEmpty(attachment.Module)) { var navigate = ChatButton(UiText.Get("Open in the VBE")); navigate.Click += (s, e) => NavigateAttachment(attachment); content.Children.Add(navigate); }
                body.Children.Add(new Expander { Header = attachment.Label + " · " + attachment.Text.Length + UiText.Get(" characters"), Content = content, FontSize = 11 });
            }
            if (entry.Speaker == "Vérification" && entry.Attachments != null && entry.Attachments.Length > 0) {
                var fix = ChatButton(UiText.Get("Prepare a fix")); fix.Click += (s, e) => { if (busy) return; modePicker.SelectedItem = ChatMode.Agent; prompt.Text = "/corriger " + entry.Text; draftAttachments.AddRange(entry.Attachments); RefreshContextChips(); }; body.Children.Add(fix);
            }
            if (entry.References != null && entry.References.Length > 0)
            {
                var refs = new WrapPanel { Margin = new Thickness(0, 9, 0, 0) };
                foreach (var reference in entry.References)
                {
                    var link = ChatButton(reference.Token);
                    link.ToolTip = UiText.Get("Open this reference in the VBE.");
                    link.FontSize = 11; link.Margin = new Thickness(0, 0, 5, 4);
                    link.Padding = new Thickness(6, 3, 6, 3);
                    link.Click += (s, e) => NavigateReference(reference);
                    refs.Children.Add(link);
                }
                body.Children.Add(refs);
            }
            return new Border { Child = body, CornerRadius = new CornerRadius(12),
                Background = Ink(entry.Speaker == "Erreur" ? "#FEF2F2" : user ? "#EFF6FF" : "#FFFFFF"),
                BorderBrush = Ink(user ? "#DBEAFE" : "#E2E8F0"), BorderThickness = new Thickness(1),
                Padding = new Thickness(14), Margin = new Thickness(user ? 30 : 0, 0, user ? 0 : 4, 14) };
        }

        /// <summary>Copie le texte dans le presse-papiers et signale les erreurs à l’interface.</summary>
        /// <param name="text">Texte à copier.</param>
        private void CopyText(string text)
        {
            try { WriteClipboard(text ?? ""); }
            catch (Exception ex) { SetStatus(UiText.Get("Unable to copy: ") + ex.Message); }
        }

        /// <summary>Ajoute le rendu Markdown avec les références connues et leur navigation VBE.</summary>
        /// <param name="body">Conteneur auquel ajouter le rendu.</param>
        /// <param name="content">Contenu Markdown.</param>
        private void RenderMarkdown(StackPanel body, string content)
        {
            var refs = transcriptEntries.SelectMany(x => x.References ?? new VbeChatReference[0])
                .GroupBy(x => x.Token).ToDictionary(x => x.Key, x => x.Last());
            body.Children.Add(ChatMarkdown.Render(content, refs, NavigateReference, SetStatus));
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
            var body = new StackPanel();
            var heading = new DockPanel();
            var count = new TextBlock { Text = "+" + change.Rows.Count(r => r.Kind == CodeDiffKind.Added) +
                "  −" + change.Rows.Count(r => r.Kind == CodeDiffKind.Removed), Foreground = Ink("#15803D"),
                FontSize = 12, VerticalAlignment = VerticalAlignment.Center };
            DockPanel.SetDock(count, System.Windows.Controls.Dock.Right); heading.Children.Add(count);
            var link = ChatButton("#" + change.Project + "." + change.Module);
            link.ToolTip = UiText.Get("Open the modified module in the VBE.");
            link.HorizontalAlignment = HorizontalAlignment.Left; link.Background = Brushes.Transparent;
            link.Click += (s, e) => NavigateReference(new VbeChatReference { Project = change.Project, Module = change.Module });
            heading.Children.Add(link); body.Children.Add(heading);
            body.Children.Add(new Expander { Header = UiText.Get("Change diff"), Content = new ChatDiffView(change.Before, change.After),
                IsExpanded = true, Margin = new Thickness(0, 6, 0, 8), FontSize = 12, Foreground = Ink("#475569") });
            var actions = new WrapPanel();
            var restore = ChatButton(change.Restored ? UiText.Get("Change undone") : UiText.Get("Undo change"));
            restore.ToolTip = UiText.Get("Restore the code before this change after checking for conflicts.");
            restore.Click += (s, e) => {
                if (busy || tools == null) return;
                try { EnsureCurrentScope(); }
                catch (Exception ex) { SetStatus(ex.Message); return; }
                var result = tools.RestoreCodeChange(change);
                if (!result.Ok) AddTranscriptMessage("Erreur", UiText.Get("Unable to undo: ") + result.Error);
                RefreshCodeChangeCards(); SaveCurrentSession();
            };
            rollbackButtons[change] = restore;
            actions.Children.Add(restore);
            var blocks = ChatButton(UiText.Get("Undo a block…"));
            blocks.ToolTip = UiText.Get("Choose a block to undo while keeping other changes.");
            blocks.Click += (s, e) => {
                var menu = new ContextMenu();
                foreach (var hunk in CodeRollback.Hunks(change.Before, change.After).Where(x => !change.RestoredHunks.Contains(x.Index))) {
                    var item = new MenuItem { Header = UiText.Get("Block ") + (hunk.Index + 1) + " · L" + (hunk.AfterStart + 1) + " · +" + hunk.After.Length + " −" + hunk.Before.Length, IsEnabled = !busy };
                    item.Click += (a, b) => RollbackIntervention(change, hunk.Index, false); menu.Items.Add(item);
                }
                blocks.ContextMenu = menu; menu.PlacementTarget = blocks; menu.IsOpen = true;
            }; actions.Children.Add(blocks);
            if (!string.IsNullOrEmpty(change.TurnId)) { var all = ChatButton(UiText.Get("Undo turn")); all.ToolTip = UiText.Get("Undo changes from this turn after checking for conflicts."); all.Click += (s, e) => RollbackIntervention(change, null, true); actions.Children.Add(all); }
            var state = new TextBlock { FontSize = 11, Foreground = Ink("#64748B"),
                Margin = new Thickness(8, 8, 0, 0) };
            changeStates[change] = state; actions.Children.Add(state); body.Children.Add(actions);
            return new Border { Child = body, Background = Ink("#FFFFFF"), BorderBrush = Ink("#CBD5E1"),
                BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(10),
                Padding = new Thickness(10), Margin = new Thickness(0, 0, 4, 14) };
        }

        /// <summary>Actualise l’activation et le libellé des boutons selon l’état restauré et l’activité courante.</summary>
        private void RefreshCodeChangeCards()
        {
            RefreshFormCutCards();
            foreach (var pair in rollbackButtons)
            {
                pair.Value.IsEnabled = !pair.Key.Restored && !busy;
                pair.Value.Content = pair.Key.Restored ? UiText.Get("Change undone") : UiText.Get("Undo change");
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
                int index = visibleEntries.IndexOf(entry);
                if (index >= 0) { visibleEntries.RemoveAt(index); visibleEntries.Insert(index, entry); }
            }
            else
            {
                TextBox live;
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
            var welcome = new StackPanel { Margin = new Thickness(12, 45, 12, 16) };
            welcome.Children.Add(new TextBlock { Text = UiText.Get("What would you like to build?"), FontSize = 23,
                FontWeight = FontWeights.SemiBold, Foreground = Ink("#0F172A"), TextWrapping = TextWrapping.Wrap });
            welcome.Children.Add(new TextBlock { Text = UiText.Get("Add #aModule or @aFunction to work on your code."),
                FontSize = 13, Foreground = Ink("#64748B"), Margin = new Thickness(0, 12, 0, 22), TextWrapping = TextWrapping.Wrap });
            foreach (string suggestion in new[] { UiText.Get("Explain a procedure"), UiText.Get("Fix an error"), UiText.Get("Improve the code") })
            {
                string seed = suggestion;
                var button = ChatButton(seed + "  →");
                button.HorizontalContentAlignment = HorizontalAlignment.Left;
                button.Margin = new Thickness(0, 0, 0, 8);
                button.Click += (s, e) => { prompt.Text = seed + " "; prompt.CaretIndex = prompt.Text.Length; prompt.Focus(); };
                welcome.Children.Add(button);
            }
            visibleEntries.Add(welcome);
        }
    }
}
