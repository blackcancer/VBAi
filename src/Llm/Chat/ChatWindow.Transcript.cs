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
    internal sealed partial class ChatWindow
    {
        private ScrollViewer conversationScroll;
        private StackPanel conversationItems;
        private bool followConversation = true;
        private readonly List<ChatEntry> transcriptEntries = new List<ChatEntry>();
        private readonly Dictionary<ChatEntry, FrameworkElement> entryViews = new Dictionary<ChatEntry, FrameworkElement>();
        private readonly Dictionary<string, ChatEntry> liveEntries = new Dictionary<string, ChatEntry>();
        private readonly Dictionary<string, TextBox> liveTexts = new Dictionary<string, TextBox>();
        private readonly Dictionary<CodeChange, Button> rollbackButtons = new Dictionary<CodeChange, Button>();
        private readonly Dictionary<CodeChange, TextBlock> changeStates = new Dictionary<CodeChange, TextBlock>();

        private void InitializeTranscript()
        {
            conversationItems = new StackPanel { Margin = new Thickness(20, 6, 20, 12) };
            conversationScroll = new ScrollViewer { Content = conversationItems,
                Background = Ink("#F8FAFC"), HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
            conversationScroll.ScrollChanged += (s, e) => {
                if (Math.Abs(e.ExtentHeightChange) < 0.1 && Math.Abs(e.ViewportHeightChange) < 0.1 && Math.Abs(e.VerticalChange) > 0.1)
                    followConversation = conversationScroll.ScrollableHeight - conversationScroll.VerticalOffset < 32;
                jumpToLatest.Visibility = followConversation ? Visibility.Collapsed : Visibility.Visible;
            };
            AutomationProperties.SetName(conversationScroll, "Conversation");
            transcriptHost.Child = conversationScroll;
        }

        private void ClearTranscript()
        {
            conversationItems?.Children.Clear();
            transcriptEntries.Clear(); entryViews.Clear(); liveEntries.Clear(); liveTexts.Clear();
            rollbackButtons.Clear(); changeStates.Clear();
            followConversation = true;
        }

        private void FollowLatest()
        {
            if (!followConversation || conversationScroll == null) return;
            conversationScroll.Dispatcher.BeginInvoke(DispatcherPriority.Background, new Action(() => {
                if (followConversation) conversationScroll.ScrollToEnd();
            }));
        }

        private static TextBox SelectableText(string text, bool code = false)
        {
            return new TextBox { Text = text ?? "", IsReadOnly = true, AcceptsReturn = true,
                TextWrapping = code ? TextWrapping.NoWrap : TextWrapping.Wrap,
                BorderThickness = new Thickness(0), Background = Brushes.Transparent,
                Foreground = Ink(code ? "#E2E8F0" : "#334155"), FontSize = 13,
                FontFamily = new FontFamily(code ? "Consolas" : "Segoe UI"),
                Padding = new Thickness(0), IsReadOnlyCaretVisible = true };
        }

        private void AddTranscriptMessage(string speaker, string content)
        {
            AddEntry(new ChatEntry { Speaker = speaker, Text = content });
        }

        private void AddEntry(ChatEntry entry)
        {
            if (conversationItems == null) return;
            if (transcriptEntries.Count == 0) conversationItems.Children.Clear();
            transcriptEntries.Add(entry);
            var view = RenderEntry(entry);
            entryViews[entry] = view;
            conversationItems.Children.Add(view);
            FollowLatest();
            ScheduleSessionSave();
        }

        private FrameworkElement RenderEntry(ChatEntry entry)
        {
            if (entry.Change != null) return RenderChange(entry.Change);
            if (entry.Speaker == "Réflexion" || entry.Speaker == "Outil")
            {
                var text = SelectableText(entry.Text);
                text.Margin = new Thickness(18, 8, 4, 6);
                if (!string.IsNullOrEmpty(entry.StreamId)) liveTexts[entry.StreamId] = text;
                return new Expander { Header = entry.Speaker == "Réflexion" ? "Réflexion · résumé" : "Activité de l'agent",
                    IsExpanded = entry.Speaker == "Réflexion", Content = text, FontSize = 12,
                    Foreground = Ink("#64748B"), Margin = new Thickness(0, 4, 0, 14) };
            }
            bool user = entry.Speaker == "Vous";
            var body = new StackPanel();
            var heading = new DockPanel { Margin = new Thickness(0, 0, 0, 8) };
            var copy = ChatButton("Copier");
            copy.FontSize = 10; copy.MinHeight = 20; copy.Padding = new Thickness(6, 2, 6, 2);
            copy.Background = Brushes.Transparent;
            copy.Click += (s, e) => CopyText(entry.Text);
            DockPanel.SetDock(copy, System.Windows.Controls.Dock.Right); heading.Children.Add(copy);
            heading.Children.Add(new TextBlock { Text = user ? "VOUS" : entry.Speaker.ToUpperInvariant(),
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
            if (!string.IsNullOrWhiteSpace(entry.AttachedMemory))
                body.Children.Add(new Expander { Header = "Mémoire du document jointe",
                    Content = SelectableText(entry.AttachedMemory), FontSize = 11, Margin = new Thickness(0, 8, 0, 0) });
            if (entry.References != null && entry.References.Length > 0)
            {
                var refs = new WrapPanel { Margin = new Thickness(0, 9, 0, 0) };
                foreach (var reference in entry.References)
                {
                    var link = ChatButton(reference.Token);
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

        private void CopyText(string text)
        {
            try { Clipboard.SetText(text ?? ""); }
            catch (Exception ex) { SetStatus("Copie impossible : " + ex.Message); }
        }

        private void RenderMarkdown(StackPanel body, string content)
        {
            string[] parts = content.Split(new[] { "```" }, StringSplitOptions.None);
            for (int i = 0; i < parts.Length; i++)
            {
                string part = parts[i].Trim('\r', '\n');
                if (part.Length == 0) continue;
                if (i % 2 == 1)
                {
                    string language = "CODE";
                    int end = part.IndexOf('\n');
                    if (end > 0 && end < 20 && Regex.IsMatch(part.Substring(0, end).Trim(), "^[A-Za-z0-9_+.-]+$"))
                    { language = part.Substring(0, end).Trim().ToUpperInvariant(); part = part.Substring(end + 1); }
                    string source = part;
                    var code = new StackPanel();
                    var header = new DockPanel { Margin = new Thickness(0, 0, 0, 8) };
                    var copy = ChatButton("Copier le code");
                    copy.MinHeight = 22; copy.FontSize = 10; copy.Padding = new Thickness(6, 2, 6, 2);
                    copy.Click += (s, e) => CopyText(source);
                    DockPanel.SetDock(copy, System.Windows.Controls.Dock.Right); header.Children.Add(copy);
                    header.Children.Add(new TextBlock { Text = language, Foreground = Ink("#94A3B8"), FontSize = 10 });
                    code.Children.Add(header);
                    code.Children.Add(new ScrollViewer { Content = SelectableText(source, true),
                        HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                        VerticalScrollBarVisibility = ScrollBarVisibility.Auto, MaxHeight = 320 });
                    body.Children.Add(new Border { Child = code, Background = Ink("#172033"),
                        CornerRadius = new CornerRadius(8), Padding = new Thickness(12), Margin = new Thickness(0, 8, 0, 8) });
                }
                else
                {
                    // Render links only for structured references that were actually selected.
                    var refs = transcriptEntries.SelectMany(x => x.References ?? new VbeChatReference[0])
                        .GroupBy(x => x.Token).ToDictionary(x => x.Key, x => x.Last());
                    foreach (string line in part.Split('\n'))
                    {
                        var text = new TextBlock { TextWrapping = TextWrapping.Wrap, FontSize = 13,
                            Foreground = Ink("#334155"), Margin = new Thickness(0, 0, 0, 5) };
                        string value = line.TrimEnd('\r');
                        if (value.StartsWith("# "))
                        { value = value.Substring(2); text.FontSize = 18; text.FontWeight = FontWeights.SemiBold; }
                        else if (value.StartsWith("## "))
                        { value = value.Substring(3); text.FontSize = 15; text.FontWeight = FontWeights.SemiBold; }
                        if (value.StartsWith("- ")) value = "• " + value.Substring(2);
                        foreach (string token in Regex.Split(value, "([#@][\\p{L}\\p{N}_.:]+|\\*\\*[^*]+\\*\\*|\x60[^\x60]+\x60)"))
                        {
                            VbeChatReference reference;
                            if (refs.TryGetValue(token, out reference))
                            {
                                var link = new Hyperlink(new Run(token)) { Foreground = Ink("#2563EB") };
                                var target = reference;
                                link.Click += (s, e) => NavigateReference(target);
                                text.Inlines.Add(link);
                            }
                            else if (token.StartsWith("**") && token.EndsWith("**") && token.Length > 4)
                                text.Inlines.Add(new Bold(new Run(token.Substring(2, token.Length - 4))));
                            else if (token.StartsWith("\x60") && token.EndsWith("\x60") && token.Length > 2)
                                text.Inlines.Add(new Run(token.Substring(1, token.Length - 2)) { FontFamily = new FontFamily("Consolas"), Background = Ink("#F1F5F9") });
                            else text.Inlines.Add(new Run(token));
                        }
                        body.Children.Add(text);
                    }
                }
            }
        }

        private void AddCodeChangeCard(CodeChange change)
        {
            AddEntry(new ChatEntry { Speaker = "Code", Change = change });
            RefreshCodeChangeCards();
        }

        private FrameworkElement RenderChange(CodeChange change)
        {
            var body = new StackPanel();
            var heading = new DockPanel();
            var count = new TextBlock { Text = "+" + change.Rows.Count(r => r.Kind == CodeDiffKind.Added) +
                "  −" + change.Rows.Count(r => r.Kind == CodeDiffKind.Removed), Foreground = Ink("#15803D"),
                FontSize = 12, VerticalAlignment = VerticalAlignment.Center };
            DockPanel.SetDock(count, System.Windows.Controls.Dock.Right); heading.Children.Add(count);
            var link = ChatButton("#" + change.Project + "." + change.Module);
            link.HorizontalAlignment = HorizontalAlignment.Left; link.Background = Brushes.Transparent;
            link.Click += (s, e) => NavigateReference(new VbeChatReference { Project = change.Project, Module = change.Module });
            heading.Children.Add(link); body.Children.Add(heading);
            var diff = new DataGrid { ItemsSource = change.Rows, AutoGenerateColumns = false, IsReadOnly = true,
                CanUserAddRows = false, CanUserDeleteRows = false, CanUserSortColumns = false,
                CanUserResizeRows = false, HeadersVisibility = DataGridHeadersVisibility.None,
                GridLinesVisibility = DataGridGridLinesVisibility.None, BorderThickness = new Thickness(0),
                Background = Brushes.White, FontFamily = new FontFamily("Consolas"), FontSize = 12,
                RowHeight = 24, MaxHeight = 260, EnableRowVirtualization = true,
                SelectionMode = DataGridSelectionMode.Extended, SelectionUnit = DataGridSelectionUnit.Cell };
            diff.Columns.Add(new DataGridTextColumn { Binding = new Binding("OldLine"), Width = 40 });
            diff.Columns.Add(new DataGridTextColumn { Binding = new Binding("NewLine"), Width = 40 });
            diff.Columns.Add(new DataGridTextColumn { Binding = new Binding("Sign"), Width = 24 });
            diff.Columns.Add(new DataGridTextColumn { Binding = new Binding("Text"), Width = DataGridLength.SizeToCells });
            var style = new Style(typeof(DataGridRow));
            foreach (var kind in new[] { CodeDiffKind.Added, CodeDiffKind.Removed })
            {
                var trigger = new DataTrigger { Binding = new Binding("Kind"), Value = kind };
                trigger.Setters.Add(new Setter(Control.BackgroundProperty, Ink(kind == CodeDiffKind.Added ? "#E8F7ED" : "#FFF0F0")));
                trigger.Setters.Add(new Setter(Control.ForegroundProperty, Ink(kind == CodeDiffKind.Added ? "#166534" : "#991B1B")));
                style.Triggers.Add(trigger);
            }
            diff.RowStyle = style;
            body.Children.Add(new Expander { Header = "Diff de la modification", Content = diff,
                IsExpanded = true, Margin = new Thickness(0, 6, 0, 8), FontSize = 12, Foreground = Ink("#475569") });
            var actions = new WrapPanel();
            var restore = ChatButton(change.Restored ? "Modification annulée" : "Annuler la modification");
            restore.Click += (s, e) => {
                if (busy || tools == null) return;
                try { EnsureCurrentScope(); }
                catch (Exception ex) { SetStatus(ex.Message); return; }
                var result = tools.RestoreCodeChange(change);
                if (!result.Ok) AddTranscriptMessage("Erreur", "Annulation impossible : " + result.Error);
                RefreshCodeChangeCards(); SaveCurrentSession();
            };
            rollbackButtons[change] = restore;
            actions.Children.Add(restore);
            var state = new TextBlock { FontSize = 11, Foreground = Ink("#64748B"),
                Margin = new Thickness(8, 8, 0, 0) };
            changeStates[change] = state; actions.Children.Add(state); body.Children.Add(actions);
            return new Border { Child = body, Background = Brushes.White, BorderBrush = Ink("#CBD5E1"),
                BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(10),
                Padding = new Thickness(10), Margin = new Thickness(0, 0, 4, 14) };
        }

        private void RefreshCodeChangeCards()
        {
            foreach (var pair in rollbackButtons)
            {
                pair.Value.IsEnabled = !pair.Key.Restored && !busy;
                pair.Value.Content = pair.Key.Restored ? "Modification annulée" : "Annuler la modification";
                changeStates[pair.Key].Text = pair.Key.Restored ? "Code restauré" : "Appliquée";
            }
        }

        private readonly HashSet<string> completedStreams = new HashSet<string>();
        private string streamedFinalText;

        private void ReceiveChatUpdate(string kind, string id, string text, bool complete)
        {
            if (IsDisposed || conversationItems == null) return;
            ChatEntry entry;
            if (!liveEntries.TryGetValue(id, out entry))
            {
                entry = new ChatEntry { Speaker = kind == "summary" ? "Réflexion" :
                    kind == "tool" ? "Outil" : "Assistant", StreamId = id, Text = "" };
                liveEntries[id] = entry;
                AddEntry(entry);
            }
            entry.Text = complete ? (text ?? entry.Text) : entry.Text + (text ?? "");
            if (complete)
            {
                completedStreams.Add(id);
                if (kind == "final") streamedFinalText = entry.Text;
                FrameworkElement old;
                if (entryViews.TryGetValue(entry, out old))
                {
                    int index = conversationItems.Children.IndexOf(old);
                    var view = RenderEntry(entry);
                    conversationItems.Children.RemoveAt(index);
                    conversationItems.Children.Insert(index, view);
                    entryViews[entry] = view;
                }
            }
            else
            {
                TextBox live;
                if (liveTexts.TryGetValue(id, out live)) live.Text = entry.Text;
            }
            FollowLatest(); ScheduleSessionSave();
        }

        private void CompleteAssistantResponse(string text)
        {
            if (!string.Equals(streamedFinalText, text, StringComparison.Ordinal)) AddTranscriptMessage("Assistant", text);
            streamedFinalText = null;
        }

        private void ShowWelcome()
        {
            if (conversationItems == null || transcriptEntries.Count > 0) return;
            var welcome = new StackPanel { Margin = new Thickness(12, 45, 12, 16) };
            welcome.Children.Add(new TextBlock { Text = "Que souhaitez-vous développer ?", FontSize = 23,
                FontWeight = FontWeights.SemiBold, Foreground = Ink("#0F172A"), TextWrapping = TextWrapping.Wrap });
            welcome.Children.Add(new TextBlock { Text = "Ajoutez #unModule ou @uneFonction pour travailler sur votre code.",
                FontSize = 13, Foreground = Ink("#64748B"), Margin = new Thickness(0, 12, 0, 22), TextWrapping = TextWrapping.Wrap });
            foreach (string suggestion in new[] { "Expliquer une procédure", "Corriger une erreur", "Améliorer le code" })
            {
                string seed = suggestion;
                var button = ChatButton(seed + "  →");
                button.HorizontalContentAlignment = HorizontalAlignment.Left;
                button.Margin = new Thickness(0, 0, 0, 8);
                button.Click += (s, e) => { prompt.Text = seed + " "; prompt.CaretIndex = prompt.Text.Length; prompt.Focus(); };
                welcome.Children.Add(button);
            }
            conversationItems.Children.Add(welcome);
        }
    }
}
