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
    /// <summary>Compose les messages et gère la recherche de références du projet dans l’éditeur.</summary>
    internal sealed partial class ChatWindow
    {
        /// <summary>Zone de saisie WPF des demandes de conversation.</summary>
        private WpfTextBox prompt;
        /// <summary>Fenêtre contextuelle des commandes et références trouvées.</summary>
        private Popup referencePopup;
        /// <summary>Résultats sélectionnables de la recherche contextuelle.</summary>
        private Forms.ListBox referenceList;
        /// <summary>Hosts the live reference suggestions shown below the composer.</summary>
private ChatSuggestionsView referenceView;
        /// <summary>État de la recherche ou instructions de sélection affichés sous la liste.</summary>
        private Forms.Label referenceStatus;
        /// <summary>Index des références du projet VBA courant.</summary>
        private VbeChatReferences referenceIndex;
        /// <summary>Minuteur WinForms qui fait progresser la construction de l’index.</summary>
        private Forms.Timer referenceTimer;
        /// <summary>Références insérées dans la saisie et disponibles pour le prochain message.</summary>
        private readonly List<VbeChatReference> selectedReferences = new List<VbeChatReference>();
        /// <summary>Indique si l’index a été démarré pour la recherche courante.</summary>
        private bool referenceIndexReady;
        /// <summary>Position du préfixe # ou @ de la référence actuellement recherchée.</summary>
        private int referenceStart = -1;
        /// <summary>Position finale du jeton de référence inséré pour masquer les suggestions.</summary>
        private int acceptedTokenEnd = -1;

        /// <summary>Construit la zone de saisie, ses suggestions et les abonnements associés.</summary>
        /// <param name="session">Session VBE qui fournit l’index des références du projet.</param>
        private void InitializeComposer(VbeSession session)
        {
            prompt = promptHost.Editor;
            prompt.Language = XmlLanguage.GetLanguage(UiText.Culture.Name);
            prompt.FlowDirection = UiText.Culture.TextInfo.IsRightToLeft ? FlowDirection.RightToLeft : FlowDirection.LeftToRight;
            AutomationProperties.SetName(prompt, UiText.Get("Your request; Enter to send, Shift+Enter for a new line"));

            referenceView = new ChatSuggestionsView();
            referenceList = referenceView.targets; referenceStatus = referenceView.status;
            referenceList.MouseUp += (sender,args) => AcceptReference();
            referenceList.DrawItem += (sender,args) => {
                if (args.Index < 0 || args.Index >= referenceList.Items.Count) return;
                args.DrawBackground();
                var value = referenceList.Items[args.Index];
                string token = value is VbeChatReference reference ? reference.DisplayToken : ((ChatCommand)value).DisplayToken;
                string kind = value is VbeChatReference item ? item.DisplayKind : ((ChatCommand)value).DisplayKind;
                var bounds = args.Bounds; bounds.Inflate(-8,0);
                Forms.TextRenderer.DrawText(args.Graphics, token + " · " + kind, args.Font, bounds, args.ForeColor,
                    Forms.TextFormatFlags.VerticalCenter | Forms.TextFormatFlags.EndEllipsis | Forms.TextFormatFlags.NoPrefix);
                args.DrawFocusRectangle();
            };
            referencePopup = new Popup { PlacementTarget = prompt, Placement = PlacementMode.Relative,
                StaysOpen = false, AllowsTransparency = true, Child = new ChatDesignerHost(referenceView) };
            referenceIndex = new VbeChatReferences(session);
            referenceTimer = new Forms.Timer { Interval = 30 };
            referenceTimer.Tick += (sender, args) => {
                referenceIndex.Step();
                if (!referenceIndex.IsLoading) referenceTimer.Stop();
            };
            // Event callbacks see a complete composer, including its popup and timer.
            prompt.PreviewKeyDown += PromptKeyDown;
            prompt.TextChanged += (sender, args) => {
                UpdateBudgetControls();
                acceptedTokenEnd = -1;
                UpdateReferences();
                RefreshContextChips();
                ScheduleSessionSave();
            };
            prompt.SelectionChanged += (sender, args) => UpdateReferences();
            referenceIndex.Changed += UpdateReferences;
        }

        /// <summary>Traite les touches de retour, navigation et validation des suggestions.</summary>
        /// <param name="sender">Contrôle à l’origine de la touche reçue.</param>
        /// <param name="e">Informations sur la touche et état de traitement.</param>
        private void PromptKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter &&
                (ReadModifiers() & System.Windows.Input.ModifierKeys.Shift) != 0)
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
            if (e.Key == Key.Enter && (ReadModifiers() & System.Windows.Input.ModifierKeys.Shift) == 0)
            {
                e.Handled = true;
                _ = SendAsync();
            }
        }

        /// <summary>Met à jour les suggestions de commandes et références autour du curseur.</summary>
        private void UpdateReferences()
        {
            if (prompt == null) return;
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
            SetSuggestionTargets(matches);
            referenceList.SelectedIndex = matches.Length > 0 ? 0 : -1;
            var rectangle = prompt.GetRectFromCharacterIndex(caret);
            referenceView.Width = prompt.ActualWidth > 0 ? (int)Math.Max(240, Math.Min(480, prompt.ActualWidth - 14)) : 350;
            referenceList.Height = Math.Max(30, Math.Min(240, matches.Length * referenceList.ItemHeight));
            ((FrameworkElement)referencePopup.Child).Width = referenceView.Width;
            referenceStatus.MaximumSize = new System.Drawing.Size(Math.Max(200, referenceView.Width - 24), 0);
            referencePopup.HorizontalOffset = Math.Max(0, Math.Min(rectangle.Left, prompt.ActualWidth - referenceView.Width - 10));
            referencePopup.VerticalOffset = rectangle.Bottom + 3;
            referenceList.Visible = matches.Length > 0;
            referenceStatus.Text = referenceIndex.IsLoading ? UiText.Get("Searching the project…") :
                !string.IsNullOrWhiteSpace(referenceIndex.Error) ? referenceIndex.Error :
                matches.Length == 0 ? UiText.Get("No matching target") : UiText.Get("↑ ↓ Browse · Enter Insert · Esc Close");
            referencePopup.IsOpen = true;
        }

        /// <summary>Replaces variable suggestion data without rebuilding the Designer list.</summary>
        /// <param name="targets">Reference or command records.</param>
        private void SetSuggestionTargets(object[] targets)
        {
            referenceList.BeginUpdate();
            try {
                referenceList.DataSource = null;
                referenceList.Items.Clear(); referenceList.Items.AddRange(targets);
            } finally { referenceList.EndUpdate(); }
        }
        /// <summary>Affiche les commandes qui correspondent au préfixe saisi en début de message.</summary>
        /// <param name="caret">Position du curseur dans la zone de saisie.</param>
        /// <returns>true si une liste de commandes correspondantes est affichée.</returns>
        private bool TryShowCommands(int caret)
        {
            if (!prompt.Text.StartsWith("/") || caret < 1 || prompt.Text.Substring(0, caret).Any(char.IsWhiteSpace)) return false;
            var matches = ChatCommand.All.Where(x => x.Token.StartsWith(prompt.Text.Substring(0, caret), StringComparison.OrdinalIgnoreCase) || x.EnglishToken.StartsWith(prompt.Text.Substring(0, caret), StringComparison.OrdinalIgnoreCase)).ToArray();
            SetSuggestionTargets(matches); referenceList.SelectedIndex = matches.Length > 0 ? 0 : -1;
            referenceList.Visible = true; referenceStatus.Text = UiText.Get("Command · Enter to choose · add your instructions");
            referencePopup.HorizontalOffset = 0; referencePopup.VerticalOffset = prompt.GetRectFromCharacterIndex(caret).Bottom + 3;
            referenceView.Width = (int)Math.Max(240, Math.Min(420, prompt.ActualWidth - 14));
            ((FrameworkElement)referencePopup.Child).Width = referenceView.Width;
            referenceList.Height = Math.Max(30, Math.Min(240, matches.Length * referenceList.ItemHeight));
            referenceStatus.MaximumSize = new System.Drawing.Size(Math.Max(200, referenceView.Width - 24), 0);
            referencePopup.IsOpen = true; return true;
        }

        /// <summary>Indique si le caractère peut appartenir à un identifiant de référence.</summary>
        /// <param name="value">Caractère à tester.</param>
        /// <returns>true si le caractère est autorisé dans un jeton de référence.</returns>
        private static bool IsReferenceChar(char value)
        {
            return char.IsLetterOrDigit(value) || value == '_' || value == '.' || value == ':';
        }

        /// <summary>Insère la commande ou référence sélectionnée dans le message.</summary>
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

        /// <summary>Ferme les suggestions et réinitialise l’état de recherche.</summary>
        private void HideReferences()
        {
            if (referencePopup != null) referencePopup.IsOpen = false;
            referenceTimer?.Stop();
            referenceIndexReady = false;
            referenceStart = -1;
        }

        /// <summary>Remplace les jetons sélectionnés par le contexte source sous la limite de taille.</summary>
        /// <param name="question">Texte saisi avant résolution des références sélectionnées.</param>
        /// <returns>Texte avec le contexte source des références sélectionnées, si elles sont présentes.</returns>
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

        /// <summary>Retourne les références sélectionnées dont le jeton apparaît dans le texte.</summary>
        /// <param name="text">Texte dans lequel rechercher les références sélectionnées.</param>
        /// <returns>Références sélectionnées dont le jeton est présent, sans doublons.</returns>
        private VbeChatReference[] CurrentReferences(string text)
        {
            return selectedReferences.Where(item => ContainsToken(text, item.Token))
                .GroupBy(item => item.Token, StringComparer.Ordinal).Select(group => group.Last()).ToArray();
        }

        /// <summary>Reconstruit les boutons de mémoire, références et pièces jointes.</summary>
        private void RefreshContextChips()
        {
            if (prompt == null) return;
            contextChips.SuspendLayout();
            try
            {
                while (contextChips.Controls.Count > 0) contextChips.Controls[0].Dispose();
                if (attachMemory.Checked && !string.IsNullOrWhiteSpace(queuedDraftMemory ?? projectMemory))
                {
                    var memory = new ChatContextChipView();
                    memory.ShowItem(UiText.Get("Attached memory · ×"), false, null, UiText.Get("Remove notes from the next message"));
                    memory.RemoveRequested += (s, e) => attachMemory.Checked = false;
                    UiTheme.Apply(memory);
                    contextChips.Controls.Add(memory);
                }
                foreach (var item in CurrentReferences(prompt.Text))
                {
                    var chip = new ChatContextChipView();
                    chip.ShowItem(item.Token, true, UiText.Get("Open in the VBE"), UiText.Get("Remove this reference from the context"));
                    chip.OpenRequested += (s, e) => NavigateReference(item);
                    chip.RemoveRequested += (s, e) => { selectedReferences.RemoveAll(value => value.Token == item.Token); RefreshContextChips(); ScheduleSessionSave(); };
                    UiTheme.Apply(chip);
                    contextChips.Controls.Add(chip);
                }
                foreach (var attachment in draftAttachments.ToArray())
                {
                    var chip = new ChatContextChipView();
                    chip.ShowItem(attachment.Label + " · ×", false, null, UiText.Get("Remove this selection"));
                    chip.RemoveRequested += (s, e) => { draftAttachments.Remove(attachment); RefreshContextChips(); ScheduleSessionSave(); };
                    UiTheme.Apply(chip);
                    contextChips.Controls.Add(chip);
                }
                contextChips.Visible = contextChips.Controls.Count > 0;
            }
            finally { contextChips.ResumeLayout(true); }
        }

        /// <summary>Navigue vers le module référencé ou insère le nom du projet dans la saisie.</summary>
        /// <param name="reference">Référence VBE à ouvrir ou insérer.</param>
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
                referenceIndex.Navigate(reference);
            }
            catch (Exception ex) { SetStatus(UiText.Get("Unable to navigate: ") + ex.Message); }
        }

        /// <summary>Indique si un jeton apparaît comme référence complète, délimité par des caractères non identifiants.</summary>
        /// <param name="text">Texte dans lequel rechercher les références sélectionnées.</param>
        /// <param name="token">Jeton dont la présence comme référence entière doit être vérifiée.</param>
        /// <returns>true si le jeton est entouré de séparateurs ou de bornes de texte.</returns>
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

        /// <summary>Arrête et libère les ressources de recherche de références.</summary>
        private void DisposeComposer()
        {
            referenceTimer?.Stop();
            referenceTimer?.Dispose();
            referenceView?.Dispose();
            HideReferences();
        }
    }
}
