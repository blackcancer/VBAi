using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Markup;
using System.Windows.Automation;
using System.Windows.Data;
using System.Windows.Media.Effects;
using Forms = System.Windows.Forms;
using WpfTextBox = System.Windows.Controls.TextBox;

namespace CodexVBE
{
    internal sealed partial class ChatWindow
    {
        private WpfTextBox prompt;
        private Popup referencePopup;
        private ListBox referenceList;
        private TextBlock referenceStatus;
        private VbeChatReferences referenceIndex;
        private Forms.Timer referenceTimer;
        private readonly List<VbeChatReference> selectedReferences = new List<VbeChatReference>();
        private bool referenceIndexReady;
        private int referenceStart = -1;
        private int acceptedTokenEnd = -1;

        private void InitializeComposer(VbeSession session)
        {
            prompt = new WpfTextBox {
                AcceptsReturn = true, TextWrapping = TextWrapping.Wrap,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                BorderThickness = new Thickness(0), FontFamily = new FontFamily("Segoe UI"),
                FontSize = 14, Foreground = Ink("#1E293B"),
                Padding = new Thickness(12), Background = Brushes.Transparent,
                Language = XmlLanguage.GetLanguage(UiText.Culture.Name)
            };
            prompt.SpellCheck.IsEnabled = true;
            AutomationProperties.SetName(prompt, UiText.Get("Your request; Enter to send, Shift+Enter for a new line"));
            promptHost.Child = prompt;
            prompt.PreviewKeyDown += PromptKeyDown;
            prompt.TextChanged += (sender, args) => {
                acceptedTokenEnd = -1;
                UpdateReferences();
                RefreshContextChips();
                ScheduleSessionSave();
            };
            prompt.SelectionChanged += (sender, args) => UpdateReferences();

            referenceList = new ListBox { Width = 350, MaxHeight = 240,
                BorderThickness = new Thickness(0), Background = Brushes.White };
            ScrollViewer.SetHorizontalScrollBarVisibility(referenceList, ScrollBarVisibility.Disabled);
            var itemLayout = new FrameworkElementFactory(typeof(DockPanel));
            var tokenText = new FrameworkElementFactory(typeof(TextBlock));
            tokenText.SetBinding(TextBlock.TextProperty, new Binding("DisplayToken"));
            tokenText.SetValue(TextBlock.FontFamilyProperty, new FontFamily("Segoe UI Semibold"));
            tokenText.SetValue(TextBlock.FontSizeProperty, 12.0);
            tokenText.SetValue(TextBlock.TextTrimmingProperty, TextTrimming.CharacterEllipsis);
            var kindText = new FrameworkElementFactory(typeof(TextBlock));
            kindText.SetBinding(TextBlock.TextProperty, new Binding("DisplayKind"));
            kindText.SetValue(TextBlock.MarginProperty, new Thickness(12, 0, 0, 0));
            kindText.SetValue(TextBlock.ForegroundProperty,
                new SolidColorBrush(Color.FromRgb(100, 116, 139)));
            kindText.SetValue(DockPanel.DockProperty, System.Windows.Controls.Dock.Right);
            itemLayout.AppendChild(kindText);
            itemLayout.AppendChild(tokenText);
            referenceList.ItemTemplate = new DataTemplate { VisualTree = itemLayout };
            var itemStyle = new Style(typeof(ListBoxItem));
            itemStyle.Setters.Add(new Setter(Control.PaddingProperty, new Thickness(10, 7, 10, 7)));
            itemStyle.Setters.Add(new Setter(Control.BorderThicknessProperty, new Thickness(0)));
            itemStyle.Setters.Add(new Setter(Control.HorizontalContentAlignmentProperty, HorizontalAlignment.Stretch));
            itemStyle.Setters.Add(new Setter(FrameworkElement.ToolTipProperty, new Binding("DisplayToken")));
            var selectedStyle = new Trigger { Property = ListBoxItem.IsSelectedProperty, Value = true };
            selectedStyle.Setters.Add(new Setter(Control.BackgroundProperty,
                new SolidColorBrush(Color.FromRgb(229, 240, 255))));
            selectedStyle.Setters.Add(new Setter(Control.ForegroundProperty,
                new SolidColorBrush(Color.FromRgb(29, 78, 216))));
            itemStyle.Triggers.Add(selectedStyle);
            referenceList.ItemContainerStyle = itemStyle;
            referenceList.PreviewMouseLeftButtonUp += (sender, args) => AcceptReference();
            referenceStatus = new TextBlock { FontSize = 11, Foreground = Ink("#64748B"),
                Margin = new Thickness(9, 5, 9, 5), TextWrapping = TextWrapping.Wrap };
            var referenceBody = new StackPanel();
            referenceBody.Children.Add(referenceList);
            referenceBody.Children.Add(referenceStatus);
            referencePopup = new Popup {
                PlacementTarget = prompt, Placement = PlacementMode.Relative,
                StaysOpen = false, AllowsTransparency = true,
                Child = new Border { Background = Brushes.White,
                    BorderBrush = new SolidColorBrush(Color.FromRgb(203, 213, 225)),
                    BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(8),
                    Padding = new Thickness(4), Child = referenceBody,
                    Effect = new DropShadowEffect { BlurRadius = 14, ShadowDepth = 3, Opacity = 0.18 } }
            };
            referenceIndex = new VbeChatReferences(session);
            referenceIndex.Changed += UpdateReferences;
            referenceTimer = new Forms.Timer { Interval = 30 };
            referenceTimer.Tick += (sender, args) => {
                referenceIndex.Step();
                if (!referenceIndex.IsLoading) referenceTimer.Stop();
            };
        }

        private void PromptKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter &&
                (System.Windows.Input.Keyboard.Modifiers & System.Windows.Input.ModifierKeys.Shift) != 0)
            { HideReferences(); return; }
            if (referencePopup.IsOpen)
            {
                if (e.Key == Key.Escape) { HideReferences(); e.Handled = true; return; }
                if (e.Key == Key.Down || e.Key == Key.Up)
                {
                    int count = referenceList.Items.Count;
                    if (count > 0) referenceList.SelectedIndex = Math.Max(0,
                        Math.Min(count - 1, referenceList.SelectedIndex + (e.Key == Key.Down ? 1 : -1)));
                    e.Handled = true;
                    return;
                }
                if ((e.Key == Key.Enter || e.Key == Key.Tab) && referenceList.SelectedItem != null)
                { AcceptReference(); e.Handled = true; return; }
            }
            if (e.Key == Key.Enter && (System.Windows.Input.Keyboard.Modifiers & System.Windows.Input.ModifierKeys.Shift) == 0)
            {
                e.Handled = true;
                _ = SendAsync();
            }
        }

        private void UpdateReferences()
        {
            if (prompt == null || referencePopup == null) return;
            int caret = prompt.CaretIndex;
            if (caret == acceptedTokenEnd) { HideReferences(); return; }
            if (TryShowCommands(caret)) return;
            int start = caret;
            while (start > 0 && IsReferenceChar(prompt.Text[start - 1])) start--;
            if (start == 0 || (prompt.Text[start - 1] != '#' && prompt.Text[start - 1] != '@') ||
                (start > 1 && IsReferenceChar(prompt.Text[start - 2])))
            { HideReferences(); return; }
            referenceStart = start - 1;
            if (!referenceIndexReady)
            {
                referenceIndexReady = true;
                referenceIndex.Refresh();
                if (referenceIndex.IsLoading) referenceTimer.Start();
            }
            string query = prompt.Text.Substring(start, caret - start);
            char prefix = prompt.Text[start - 1];
            var matches = referenceIndex.MatchPrefix(query, prefix).ToArray();
            referenceList.ItemsSource = matches;
            referenceList.SelectedIndex = matches.Length > 0 ? 0 : -1;
            var rectangle = prompt.GetRectFromCharacterIndex(caret);
            referenceList.Width = prompt.ActualWidth > 0 ? Math.Max(240, Math.Min(480, prompt.ActualWidth - 14)) : 350;
            referenceStatus.MaxWidth = referenceList.Width - 12;
            referencePopup.HorizontalOffset = Math.Max(0, Math.Min(rectangle.Left, prompt.ActualWidth - referenceList.Width - 10));
            referencePopup.VerticalOffset = rectangle.Bottom + 3;
            referenceList.Visibility = matches.Length == 0 ? Visibility.Collapsed : Visibility.Visible;
            referenceStatus.Text = referenceIndex.IsLoading ? UiText.Get("Searching the project…") :
                !string.IsNullOrWhiteSpace(referenceIndex.Error) ? referenceIndex.Error :
                matches.Length == 0 ? UiText.Get("No matching target") : UiText.Get("↑ ↓ Browse · Enter Insert · Esc Close");
            referencePopup.IsOpen = true;
        }

        private bool TryShowCommands(int caret)
        {
            if (!prompt.Text.StartsWith("/") || caret < 1 || prompt.Text.Substring(0, caret).Any(char.IsWhiteSpace)) return false;
            var matches = ChatCommand.All.Where(x => x.Token.StartsWith(prompt.Text.Substring(0, caret), StringComparison.OrdinalIgnoreCase) || x.EnglishToken.StartsWith(prompt.Text.Substring(0, caret), StringComparison.OrdinalIgnoreCase)).ToArray();
            referenceList.ItemsSource = matches; referenceList.SelectedIndex = matches.Length > 0 ? 0 : -1;
            referenceList.Visibility = Visibility.Visible; referenceStatus.Text = UiText.Get("Command · Enter to choose · add your instructions");
            referencePopup.HorizontalOffset = 0; referencePopup.VerticalOffset = prompt.GetRectFromCharacterIndex(caret).Bottom + 3;
            referenceList.Width = Math.Max(240, Math.Min(420, prompt.ActualWidth - 14));
            referenceStatus.MaxWidth = referenceList.Width - 12;
            referencePopup.IsOpen = true; return true;
        }

        private static bool IsReferenceChar(char value)
        {
            return char.IsLetterOrDigit(value) || value == '_' || value == '.' || value == ':';
        }

        private void AcceptReference()
        {
            var command = referenceList.SelectedItem as ChatCommand;
            if (command != null) {
                prompt.Select(0, prompt.CaretIndex); prompt.SelectedText = command.DisplayToken + " ";
                if (!busy) modePicker.SelectedItem = command.Mode;
                HideReferences(); prompt.CaretIndex = command.DisplayToken.Length + 1; return;
            }
            var reference = referenceList.SelectedItem as VbeChatReference;
            if (reference == null || referenceStart < 0) return;
            int caret = prompt.CaretIndex;
            int start = referenceStart;
            prompt.Select(referenceStart, caret - referenceStart);
            prompt.SelectedText = reference.Token;
            selectedReferences.Add(reference);
            HideReferences();
            prompt.Focus();
            acceptedTokenEnd = start + reference.Token.Length;
            prompt.CaretIndex = acceptedTokenEnd;
            RefreshContextChips();
        }

        private void HideReferences()
        {
            if (referencePopup != null) referencePopup.IsOpen = false;
            referenceTimer?.Stop();
            referenceIndexReady = false;
            referenceStart = -1;
        }

        private string ResolveReferences(string question)
        {
            var selected = selectedReferences.Where(item => ContainsToken(question, item.Token))
                .GroupBy(item => item.Token, StringComparer.Ordinal).Select(group => group.Last()).ToArray();
            if (selected.Length == 0) return question;
            var context = new StringBuilder(question).Append("\n\n<references-vbe>\n");
            foreach (var item in selected)
            {
                if (context.Length > 48000) throw new InvalidOperationException(UiText.Get("The # context exceeds 48,000 characters."));
                context.Append(referenceIndex.Resolve(item)).Append("\n---\n");
            }
            if (context.Length > 48000) throw new InvalidOperationException(UiText.Get("The # context exceeds 48,000 characters."));
            return context.Append("</references-vbe>").ToString();
        }

        private VbeChatReference[] CurrentReferences(string text)
        {
            return selectedReferences.Where(item => ContainsToken(text, item.Token))
                .GroupBy(item => item.Token, StringComparer.Ordinal).Select(group => group.Last()).ToArray();
        }

        private static Forms.Button ContextButton(string text)
        {
            return new Forms.Button { Text = text, AutoSize = true, Height = 27,
                FlatStyle = Forms.FlatStyle.Flat, BackColor = System.Drawing.Color.FromArgb(239, 246, 255),
                ForeColor = System.Drawing.Color.FromArgb(29, 78, 216), Margin = new Forms.Padding(2),
                Padding = new Forms.Padding(4, 0, 4, 0), Cursor = Forms.Cursors.Hand };
        }

        private void RefreshContextChips()
        {
            if (contextChips == null || prompt == null) return;
            contextChips.SuspendLayout();
            try
            {
                while (contextChips.Controls.Count > 0) contextChips.Controls[0].Dispose();
                if (attachMemory.Checked && !string.IsNullOrWhiteSpace(projectMemory))
                {
                    var memory = ContextButton(UiText.Get("Attached memory · ×"));
                    toolTips.SetToolTip(memory, UiText.Get("Remove notes from the next message"));
                    memory.Click += (s, e) => attachMemory.Checked = false;
                    contextChips.Controls.Add(memory);
                }
                foreach (var item in CurrentReferences(prompt.Text))
                {
                    var row = new Forms.FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Forms.Padding(0) };
                    var link = ContextButton(item.Token);
                    toolTips.SetToolTip(link, UiText.Get("Open in the VBE"));
                    link.Click += (s, e) => NavigateReference(item); row.Controls.Add(link);
                    var remove = ContextButton("×");
                    toolTips.SetToolTip(remove, UiText.Get("Remove this reference from the context"));
                    remove.Click += (s, e) => { selectedReferences.RemoveAll(value => value.Token == item.Token); RefreshContextChips(); ScheduleSessionSave(); };
                    row.Controls.Add(remove); contextChips.Controls.Add(row);
                }
                foreach (var attachment in draftAttachments.ToArray())
                {
                    var chip = ContextButton(attachment.Label + " · ×");
                    toolTips.SetToolTip(chip, UiText.Get("Remove this selection"));
                    chip.Click += (s, e) => { draftAttachments.Remove(attachment); RefreshContextChips(); ScheduleSessionSave(); };
                    contextChips.Controls.Add(chip);
                }
                contextChips.Visible = contextChips.Controls.Count > 0;
            }
            finally { contextChips.ResumeLayout(true); }
        }

        private void NavigateReference(VbeChatReference reference)
        {
            try
            {
                if (reference.Module == null)
                {
                    prompt.Text = reference.Token + ".";
                    prompt.CaretIndex = prompt.Text.Length;
                    prompt.Focus();
                    return;
                }
                Response result = referenceIndex.Navigate(reference);
                if (!result.Ok) SetStatus(result.Error);
            }
            catch (Exception ex) { SetStatus(UiText.Get("Unable to navigate: ") + ex.Message); }
        }

        private static bool ContainsToken(string text, string token)
        {
            int start = 0;
            while ((start = text.IndexOf(token, start, StringComparison.Ordinal)) >= 0)
            {
                int end = start + token.Length;
                if ((start == 0 || !IsReferenceChar(text[start - 1])) &&
                    (end == text.Length || !IsReferenceChar(text[end]))) return true;
                start++;
            }
            return false;
        }

        private void DisposeComposer()
        {
            referenceTimer?.Stop();
            referenceTimer?.Dispose();
            HideReferences();
        }
    }
}
