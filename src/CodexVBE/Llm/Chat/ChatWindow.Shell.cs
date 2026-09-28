using System;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace CodexVBE
{
    /// <summary>Fenêtre de conversation avec commandes de navigation, de sélection et de gestion du contexte.</summary>
    internal sealed partial class ChatWindow
    {
        /// <summary>Crée un pinceau WPF à partir d’une couleur de thème exprimée en hexadécimal.</summary>
        /// <param name="hex">Couleur à résoudre via le thème actif.</param>
        /// <returns>Pinceau solide correspondant à la couleur résolue.</returns>
        private static SolidColorBrush Ink(string hex)
        {
            return new SolidColorBrush((Color)ColorConverter.ConvertFromString(UiTheme.Map(hex)));
        }

        /// <summary>Crée un bouton WPF cohérent avec le thème, son style interactif et son nom d’accessibilité.</summary>
        /// <param name="text">Texte visible et nom d’accessibilité du bouton.</param>
        /// <param name="primary">Applique la couleur de fond et le texte du bouton principal.</param>
        /// <returns>Bouton configuré.</returns>
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

        /// <summary>Masque les surfaces initialement inactives et remplit les options du sélecteur de mode.</summary>
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

        /// <summary>Crée une conversation depuis le bouton Nouveau chat.</summary>
        /// <param name="sender">Bouton déclencheur.</param>
        /// <param name="e">Données de l’événement.</param>
        private void NewChat_Click(object sender, EventArgs e) { StartNewChat(); }
        /// <summary>Ouvre le menu des options de conversation.</summary>
        /// <param name="sender">Bouton déclencheur.</param>
        /// <param name="e">Données de l’événement.</param>
        private void Options_Click(object sender, EventArgs e) { optionsMenu.Show(options, 0, options.Height); }
        /// <summary>Présente les informations produit depuis le menu de la conversation.</summary>
        private void About_Click(object sender, EventArgs e)
        {
            using (var dialog = new AboutWindow()) ShowModal(dialog, this);
        }
        /// <summary>Demande l’ancrage de la fenêtre si aucun tour n’est actif.</summary>
        /// <param name="sender">Bouton déclencheur.</param>
        /// <param name="e">Données de l’événement.</param>
        private void Docking_Click(object sender, EventArgs e) { if (!busy) DockRequested?.Invoke(); }
        /// <summary>Affiche ou masque le panneau d’historique.</summary>
        /// <param name="sender">Bouton déclencheur.</param>
        /// <param name="e">Données de l’événement.</param>
        private void History_Click(object sender, EventArgs e) { historyPanel.Visible = !historyPanel.Visible; if (historyPanel.Visible) historyPanel.BringToFront(); }
        /// <summary>Réactive le défilement automatique et revient au dernier message.</summary>
        /// <param name="sender">Bouton déclencheur.</param>
        /// <param name="e">Données de l’événement.</param>
        private void JumpToLatest_Click(object sender, EventArgs e) { followConversation = true; FollowLatest(); }
        /// <summary>Actualise l’historique lorsque l’option des sessions archivées change.</summary>
        /// <param name="sender">Contrôle déclencheur.</param>
        /// <param name="e">Données de l’événement.</param>
        private void ShowArchived_CheckedChanged(object sender, EventArgs e) { RefreshHistory(); }
        /// <summary>Applique le titre saisi à la session active.</summary>
        /// <param name="sender">Bouton déclencheur.</param>
        /// <param name="e">Données de l’événement.</param>
        private void Rename_Click(object sender, EventArgs e) { RenameCurrentChat(); }
        /// <summary>Archive ou restaure la session active.</summary>
        /// <param name="sender">Bouton déclencheur.</param>
        /// <param name="e">Données de l’événement.</param>
        private void Archive_Click(object sender, EventArgs e) { ToggleArchiveCurrentChat(); }
        /// <summary>Enregistre la mémoire de projet affichée.</summary>
        /// <param name="sender">Bouton déclencheur.</param>
        /// <param name="e">Données de l’événement.</param>
        private void SaveMemory_Click(object sender, EventArgs e) { SaveProjectMemory(); }
        /// <summary>Actualise les puces de contexte lorsque l’inclusion de mémoire change.</summary>
        /// <param name="sender">Case à cocher déclencheuse.</param>
        /// <param name="e">Données de l’événement.</param>
        private void AttachMemory_CheckedChanged(object sender, EventArgs e) { RefreshContextChips(); }
        /// <summary>Capture le code sélectionné dans le VBE pour le contexte de conversation.</summary>
        /// <param name="sender">Bouton déclencheur.</param>
        /// <param name="e">Données de l’événement.</param>
        private void Selection_Click(object sender, EventArgs e) { CaptureSelection(); }
        /// <summary>Insère le préfixe de référence d’un projet ou module dans le compositeur.</summary>
        /// <param name="sender">Bouton déclencheur.</param>
        /// <param name="e">Données de l’événement.</param>
        private void Modules_Click(object sender, EventArgs e) { InsertReferencePrefix('#'); }
        /// <summary>Insère le préfixe de référence d’une procédure dans le compositeur.</summary>
        /// <param name="sender">Bouton déclencheur.</param>
        /// <param name="e">Données de l’événement.</param>
        private void Methods_Click(object sender, EventArgs e) { InsertReferencePrefix('@'); }
        /// <summary>Exporte la conversation courante.</summary>
        /// <param name="sender">Bouton déclencheur.</param>
        /// <param name="e">Données de l’événement.</param>
        private void Export_Click(object sender, EventArgs e) { ExportCurrentChat(); }
        /// <summary>Ouvre l’intégration GitHub pour le document enregistré de la portée courante.</summary>
        /// <param name="sender">Commande GitHub.</param>
        /// <param name="e">Données de l’événement.</param>
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
                    ShowModal(window,this);
            }
            catch (Exception ex) { SetStatus(UiText.Get("GitHub: ") + ex.Message); }
        }
        /// <summary>Inverse l’état épinglé de la session active et actualise son entrée d’historique.</summary>
        /// <param name="sender">Bouton déclencheur.</param>
        /// <param name="e">Données de l’événement.</param>
        private void Pin_Click(object sender, EventArgs e)
        {
            if (busy || currentSession == null) return;
            currentSession.Pinned = !currentSession.Pinned; SaveCurrentSession(); RefreshHistory();
        }
        /// <summary>Affiche ou masque l’éditeur de mémoire de projet.</summary>
        /// <param name="sender">Bouton déclencheur.</param>
        /// <param name="e">Données de l’événement.</param>
        private void MemoryToggle_Click(object sender, EventArgs e)
        {
            memoryPanel.Visible = !memoryPanel.Visible;
            historyLayout.RowStyles[7].Height = memoryPanel.Visible ? 226 : 0;
        }
        /// <summary>Affiche ou masque l’aperçu du contexte et le recalcule à l’ouverture.</summary>
        /// <param name="sender">Bouton déclencheur.</param>
        /// <param name="e">Données de l’événement.</param>
        private void ContextToggle_Click(object sender, EventArgs e)
        {
            contextPanel.Visible = !contextPanel.Visible;
            if (contextPanel.Visible) RefreshContextPreview();
        }
        /// <summary>Applique le mode sélectionné à la session et aux outils, puis programme sa sauvegarde.</summary>
        /// <param name="sender">Sélecteur de mode.</param>
        /// <param name="e">Données de l’événement.</param>
        private void ModePicker_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (loadingSession || busy || currentSession == null || modePicker.SelectedItem == null) return;
            currentSession.Mode = (ChatMode)modePicker.SelectedItem;
            if (tools != null) tools.Mode = currentSession.Mode;
            ScheduleSessionSave();
            SetStatus(currentSession.Mode == ChatMode.Agent ? UiText.Get("Agent: edits allowed by the VBE policy") : UiText.Get(currentSession.Mode == ChatMode.Discussion ? "Chat" : "Plan") + UiText.Get(": no edits or macro execution"));
        }
        /// <summary>Lance la vérification du projet et restaure l’état inactif du formulaire à la fin.</summary>
        /// <param name="sender">Bouton de compilation.</param>
        /// <param name="e">Données de l’événement.</param>
        private async void Compile_Click(object sender, EventArgs e)
        {
            if (busy || tools == null) return;
            SetBusy(true);
            try { await VerifyProjectAsync(); } finally { SetBusy(false); SaveCurrentSession(); }
        }
        /// <summary>Intercepte Ctrl+N pour créer une conversation, puis délègue les autres raccourcis au formulaire.</summary>
        /// <param name="msg">Message Windows traité par le contrôle.</param>
        /// <param name="keyData">Touche et modificateurs pressés.</param>
        /// <returns><see langword="true"/> si Ctrl+N a été traité ; sinon résultat de la classe de base.</returns>
        protected override bool ProcessCmdKey(ref System.Windows.Forms.Message msg, System.Windows.Forms.Keys keyData)
        {
            if (keyData == (System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.N) && !busy)
            { StartNewChat(); return true; }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        /// <summary>Insère le préfixe demandé à la position du curseur et actualise les suggestions de référence.</summary>
        /// <param name="prefix">Caractère de préfixe de référence.</param>
        private void InsertReferencePrefix(char prefix)
        {
            if (prompt == null) return;
            prompt.Focus();
            int caret = prompt.CaretIndex;
            prompt.SelectedText = (caret > 0 && !char.IsWhiteSpace(prompt.Text[caret - 1]) ? " " : "") + prefix;
            prompt.CaretIndex = prompt.SelectionStart + prompt.SelectionLength;
            UpdateReferences();
        }

        /// <summary>Crée une session de chat et replace le focus dans le compositeur.</summary>
        private void StartNewChat()
        {
            NewSession();
            prompt?.Focus();
        }
    }
}
