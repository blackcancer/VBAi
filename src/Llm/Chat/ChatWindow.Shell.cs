using System;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace CodexVBE
{
    internal sealed partial class ChatWindow
    {
        private static SolidColorBrush Ink(string hex)
        {
            return new SolidColorBrush((Color)ColorConverter.ConvertFromString(UiTheme.Map(hex)));
        }

        private static Button ChatButton(string text, bool primary = false)
        {
            var button = new Button {
                Content = text, Padding = new Thickness(12, 7, 12, 7),
                Background = Ink(primary ? "#2563EB" : "#F1F5F9"),
                Foreground = primary ? Brushes.White : Ink("#334155"),
                BorderBrush = Brushes.Transparent, BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand, FontSize = 12, FontFamily = new FontFamily("Segoe UI"),
                MinHeight = 32
            };
            var border = new FrameworkElementFactory(typeof(Border));
            border.SetValue(Border.CornerRadiusProperty, new CornerRadius(7));
            border.SetBinding(Border.BackgroundProperty, new System.Windows.Data.Binding("Background") { RelativeSource = System.Windows.Data.RelativeSource.TemplatedParent });
            border.SetBinding(Border.PaddingProperty, new System.Windows.Data.Binding("Padding") { RelativeSource = System.Windows.Data.RelativeSource.TemplatedParent });
            var content = new FrameworkElementFactory(typeof(ContentPresenter));
            content.SetValue(FrameworkElement.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            content.SetValue(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Center);
            border.AppendChild(content);
            button.Template = new ControlTemplate(typeof(Button)) { VisualTree = border };
            var style = new Style(typeof(Button));
            var hover = new Trigger { Property = UIElement.IsMouseOverProperty, Value = true };
            hover.Setters.Add(new Setter(UIElement.OpacityProperty, 0.82));
            style.Triggers.Add(hover);
            var disabled = new Trigger { Property = UIElement.IsEnabledProperty, Value = false };
            disabled.Setters.Add(new Setter(UIElement.OpacityProperty, 0.45));
            style.Triggers.Add(disabled);
            button.Style = style;
            AutomationProperties.SetName(button, text);
            return button;
        }

        // The fixed control tree belongs exclusively to InitializeComponent.
        // Runtime setup only fills data and activates the dynamic chat surfaces.
        private void InitializeShell()
        {
            historyPanel.Visible = false;
            memoryPanel.Visible = false;
            historyLayout.RowStyles[7].Height = 0;
            contextPanel.Visible = false;
            contextChips.Visible = false;
            activityBar.Visible = false;
            jumpToLatest.Visible = false;
            transcriptPlaceholder.Visible = false;
            promptPlaceholder.Visible = false;
            foreach (var mode in Enum.GetValues(typeof(ChatMode))) modePicker.Items.Add(mode);
            modePicker.FormattingEnabled = true;
            modePicker.Format += (sender, args) => { if (args.ListItem is ChatMode) args.Value = UiText.Get(args.ListItem.ToString() == "Discussion" ? "Chat" : args.ListItem.ToString()); };
            modePicker.SelectedItem = ChatMode.Agent;
        }

        private void NewChat_Click(object sender, EventArgs e) { StartNewChat(); }
        private void Options_Click(object sender, EventArgs e) { optionsMenu.Show(options, 0, options.Height); }
        private void Docking_Click(object sender, EventArgs e) { if (!busy) DockRequested?.Invoke(); }
        private void History_Click(object sender, EventArgs e) { historyPanel.Visible = !historyPanel.Visible; if (historyPanel.Visible) historyPanel.BringToFront(); }
        private void JumpToLatest_Click(object sender, EventArgs e) { followConversation = true; FollowLatest(); }
        private void ShowArchived_CheckedChanged(object sender, EventArgs e) { RefreshHistory(); }
        private void Rename_Click(object sender, EventArgs e) { RenameCurrentChat(); }
        private void Archive_Click(object sender, EventArgs e) { ToggleArchiveCurrentChat(); }
        private void SaveMemory_Click(object sender, EventArgs e) { SaveProjectMemory(); }
        private void AttachMemory_CheckedChanged(object sender, EventArgs e) { RefreshContextChips(); }
        private void Selection_Click(object sender, EventArgs e) { CaptureSelection(); }
        private void Modules_Click(object sender, EventArgs e) { InsertReferencePrefix('#'); }
        private void Methods_Click(object sender, EventArgs e) { InsertReferencePrefix('@'); }
        private void Export_Click(object sender, EventArgs e) { ExportCurrentChat(); }
        private void GitHub_Click(object sender, EventArgs e)
        {
            if (busy) return;
            try
            {
                EnsureCurrentScope();
                var scope = scopePicker.SelectedItem as MacroScope;
                if (scope == null || scope.Key.StartsWith("temporary:", StringComparison.Ordinal))
                    throw new InvalidOperationException(UiText.Get("Save the document before linking it to GitHub."));
                using (var window = new GitWindow(scopeSession.GitProject(scope.Project, scope.Key), scope.Key, scope.Label, settings.GitHubAccount))
                    window.ShowDialog(this);
            }
            catch (Exception ex) { SetStatus(UiText.Get("GitHub: ") + ex.Message); }
        }
        private void Pin_Click(object sender, EventArgs e)
        {
            if (busy || currentSession == null) return;
            currentSession.Pinned = !currentSession.Pinned; SaveCurrentSession(); RefreshHistory();
        }
        private void MemoryToggle_Click(object sender, EventArgs e)
        {
            memoryPanel.Visible = !memoryPanel.Visible;
            historyLayout.RowStyles[7].Height = memoryPanel.Visible ? 226 : 0;
        }
        private void ContextToggle_Click(object sender, EventArgs e)
        {
            contextPanel.Visible = !contextPanel.Visible;
            if (contextPanel.Visible) RefreshContextPreview();
        }
        private void ModePicker_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (loadingSession || busy || currentSession == null || modePicker.SelectedItem == null) return;
            currentSession.Mode = (ChatMode)modePicker.SelectedItem;
            if (tools != null) tools.Mode = currentSession.Mode;
            ScheduleSessionSave();
            SetStatus(currentSession.Mode == ChatMode.Agent ? UiText.Get("Agent: edits allowed by the VBE policy") : UiText.Get(currentSession.Mode == ChatMode.Discussion ? "Chat" : "Plan") + UiText.Get(": no edits or macro execution"));
        }
        private async void Compile_Click(object sender, EventArgs e)
        {
            if (busy || tools == null) return;
            SetBusy(true);
            try { await VerifyProjectAsync(); } finally { SetBusy(false); SaveCurrentSession(); }
        }
        protected override bool ProcessCmdKey(ref System.Windows.Forms.Message msg, System.Windows.Forms.Keys keyData)
        {
            if (keyData == (System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.N) && !busy)
            { StartNewChat(); return true; }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void InsertReferencePrefix(char prefix)
        {
            if (prompt == null) return;
            prompt.Focus();
            int caret = prompt.CaretIndex;
            prompt.SelectedText = (caret > 0 && !char.IsWhiteSpace(prompt.Text[caret - 1]) ? " " : "") + prefix;
            prompt.CaretIndex = prompt.SelectionStart + prompt.SelectionLength;
            UpdateReferences();
        }

        private void StartNewChat()
        {
            NewSession();
            prompt?.Focus();
        }
    }
}
