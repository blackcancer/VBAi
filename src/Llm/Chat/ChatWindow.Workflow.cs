using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace CodexVBE
{
    internal sealed partial class ChatWindow
    {
        private ComboBox modePicker;
        private CheckBox verifyAfterEdit;
        private StackPanel contextPreview;
        private readonly List<ChatAttachment> draftAttachments = new List<ChatAttachment>();
        private string activeTurnId;
        public event Action DockRequested;
        public void ReportDockFailure(string reason) { SetStatus("Ancrage indisponible : " + reason + " · vérifiez l’installation du contrôle COM."); }

        private void BuildWorkflowControls(StackPanel top, WrapPanel historyButtons, StackPanel composer)
        {
            var row = new WrapPanel { Margin = new Thickness(20, 0, 20, 8) };
            modePicker = ChatPicker("Mode de travail");
            foreach (var mode in Enum.GetValues(typeof(ChatMode))) modePicker.Items.Add(mode);
            modePicker.SelectedItem = ChatMode.Agent;
            modePicker.SelectionChanged += (s, e) => {
                if (loadingSession || busy || currentSession == null) return;
                currentSession.Mode = (ChatMode)modePicker.SelectedItem;
                if (tools != null) tools.Mode = currentSession.Mode;
                ScheduleSessionSave();
                SetStatus(currentSession.Mode == ChatMode.Agent ? "Agent : modifications autorisées par la politique VBE" : currentSession.Mode + " : aucune modification ni exécution de macro");
            };
            row.Children.Add(modePicker);
            var selection = ChatButton("Joindre la sélection"); selection.Margin = new Thickness(6, 0, 0, 0);
            selection.Click += (s, e) => CaptureSelection(); row.Children.Add(selection);
            var compile = ChatButton("Vérifier VBA"); compile.Margin = new Thickness(6, 0, 0, 0);
            compile.Click += async (s, e) => {
                if (busy) return;
                SetBusy(true);
                try { await VerifyProjectAsync(); } finally { SetBusy(false); SaveCurrentSession(); }
            };
            row.Children.Add(compile);
            top.Children.Add(row);
            verifyAfterEdit = new CheckBox { Content = "Compiler après les modifications", IsChecked = true,
                FontSize = 11, Margin = new Thickness(20, 0, 20, 8) };
            top.Children.Add(verifyAfterEdit);
            var pin = ChatButton("Épingler / détacher");
            pin.Click += (s, e) => { if (busy || currentSession == null) return; currentSession.Pinned = !currentSession.Pinned; SaveCurrentSession(); RefreshHistory(); };
            historyButtons.Children.Add(pin);
            var export = ChatButton("Exporter Markdown"); export.Click += (s, e) => ExportCurrentChat(); historyButtons.Children.Add(export);
            contextPreview = new StackPanel();
            var preview = new Expander { Header = "Contexte envoyé · inspecter et actualiser", Content = new ScrollViewer {
                Content = contextPreview, MaxHeight = 150, VerticalScrollBarVisibility = ScrollBarVisibility.Auto },
                Margin = new Thickness(12, 4, 12, 4), FontSize = 11 };
            preview.Expanded += (s, e) => RefreshContextPreview();
            composer.Children.Add(preview);
        }

        private ChatAttachment[] PrepareAttachments(string question)
        {
            var attachments = new List<ChatAttachment>();
            foreach (var reference in CurrentReferences(question))
                attachments.Add(new ChatAttachment { Label = reference.Token, Text = referenceIndex.Resolve(reference),
                    Project = reference.Project, Module = reference.Module, Sha256 = reference.Sha256, StartLine = reference.StartLine });
            foreach (var attachment in draftAttachments)
            {
                if (!string.IsNullOrEmpty(attachment.Module))
                {
                    var module = ReadWorkflow("read_module", attachment.Project, attachment.Module);
                    if (!string.Equals(Convert.ToString(module["Sha256"]), attachment.Sha256, StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException(attachment.Label + " est périmée. Retirez-la puis joignez la sélection actuelle.");
                }
                attachments.Add(attachment);
            }
            if (attachments.Sum(x => x.Text.Length) + question.Length > 48000)
                throw new InvalidOperationException("Le contexte dépasse 48 000 caractères. Retirez des éléments.");
            return attachments.ToArray();
        }

        private void RefreshContextPreview()
        {
            if (contextPreview == null || prompt == null) return;
            contextPreview.Children.Clear();
            try
            {
                EnsureCurrentScope();
                var attachments = PrepareAttachments(prompt.Text);
                foreach (var attachment in attachments)
                    contextPreview.Children.Add(new Expander { Header = attachment.Label + " · " + attachment.Text.Length + " caractères",
                        Content = SelectableText(attachment.Text) });
                int memorySize = attachMemory.IsChecked == true ? projectMemory.Length : 0;
                if (memorySize > 0) contextPreview.Children.Add(new Expander { Header = "Mémoire · " + memorySize + " caractères", Content = SelectableText(projectMemory) });
                contextPreview.Children.Add(new TextBlock { Text = (attachments.Sum(x => x.Text.Length) + memorySize) +
                    " caractères de contexte explicite. L’historique de cette conversation accompagne aussi la demande.", TextWrapping = TextWrapping.Wrap });
            }
            catch (Exception ex) { contextPreview.Children.Add(new TextBlock { Text = ex.Message, TextWrapping = TextWrapping.Wrap, Foreground = Ink("#B91C1C") }); }
        }

        private IDictionary<string, object> ReadWorkflow(string command, string project = null, string module = null)
        {
            if (scopeSession == null) throw new InvalidOperationException("Aucun hôte VBE connecté.");
            var result = scopeSession.Execute(new Request { Command = command, Project = project, Module = module });
            if (!result.Ok) throw new InvalidOperationException(result.Error);
            return json.DeserializeObject(json.Serialize(result.Data)) as IDictionary<string, object> ?? throw new InvalidOperationException("Réponse VBE inattendue.");
        }

        private void CaptureSelection()
        {
            if (busy) return;
            try
            {
                EnsureCurrentScope();
                var pane = ReadWorkflow("code_panes")["ActiveCodePane"] as IDictionary<string, object>;
                var properties = pane?["Properties"] as IDictionary<string, object>;
                if (properties == null || !properties.ContainsKey("Selection")) throw new InvalidOperationException("Sélectionnez du code dans le VBE.");
                string project = Convert.ToString(properties["Project"]), module = Convert.ToString(properties["Module"]);
                if ((scopePicker.SelectedItem as MacroScope)?.Project != project) throw new InvalidOperationException("La sélection appartient à un autre document. Choisissez sa conversation.");
                var selection = (IDictionary<string, object>)properties["Selection"];
                int start = Convert.ToInt32(selection["StartLine"]), end = Convert.ToInt32(selection["EndLine"]);
                int startColumn = Convert.ToInt32(selection["StartColumn"]), endColumn = Convert.ToInt32(selection["EndColumn"]);
                var data = ReadWorkflow("read_module", project, module);
                var lines = CodeRollback.Lines(Convert.ToString(data["Code"]));
                if (start < 1 || end < start || end > lines.Length) throw new InvalidOperationException("Sélection de code vide ou invalide.");
                var selected = lines.Skip(start - 1).Take(end - start + 1).ToArray();
                if (start != end || startColumn != endColumn)
                {
                    selected[selected.Length - 1] = selected.Last().Substring(0, Math.Min(selected.Last().Length, Math.Max(0, endColumn - 1)));
                    selected[0] = selected[0].Substring(Math.Min(selected[0].Length, Math.Max(0, startColumn - 1)));
                }
                var attachment = new ChatAttachment { Label = project + "." + module + " · sélection L" + start + "–" + end,
                    Text = string.Join("\n", selected), Project = project, Module = module, StartLine = start, Sha256 = Convert.ToString(data["Sha256"]) };
                draftAttachments.RemoveAll(x => x.Label == attachment.Label); draftAttachments.Add(attachment);
                RefreshContextChips(); RefreshContextPreview(); ScheduleSessionSave();
                SetStatus("Sélection jointe au prochain message");
            }
            catch (Exception ex) { SetStatus("Sélection : " + ex.Message); }
        }

        public void PrepareEditorAction(string command)
        {
            if (busy) { SetStatus("Attendez la fin de la réponse avant de préparer une action."); return; }
            CaptureSelection();
            var action = ChatCommand.All.First(x => x.Token == command);
            modePicker.SelectedItem = action.Mode;
            prompt.Text = command + " "; prompt.CaretIndex = prompt.Text.Length; prompt.Focus();
        }

        private async Task VerifyProjectAsync()
        {
            try
            {
                EnsureCurrentScope();
                var scope = scopePicker.SelectedItem as MacroScope;
                if (scope == null || tools == null) throw new InvalidOperationException("Aucun projet connecté.");
                SetStatus("Compilation VBA en cours…");
                var response = json.Deserialize<Response>(await tools.InvokeAsync("compile_project", json.Serialize(new { Project = scope.Project, ExpectedMode = 2 })));
                if (response == null || !response.Ok) throw new InvalidOperationException(response?.Error ?? "Réponse vide.");
                var data = json.DeserializeObject(json.Serialize(response.Data)) as IDictionary<string, object>;
                bool compiled = data != null && data.ContainsKey("Compiled") && Convert.ToBoolean(data["Compiled"]);
                string diagnostic = compiled ? "Compilation terminée : aucun diagnostic natif observé. Les macros n’ont pas été exécutées." : Convert.ToString(data?["Diagnostic"]);
                ChatAttachment[] location = null;
                if (!compiled)
                {
                    try {
                        var state = ReadWorkflow("debug_state", scope.Project);
                        string module = Convert.ToString(state["ActiveModule"]);
                        var selection = state["Selection"] as IDictionary<string, object>;
                        if (Convert.ToString(state["SelectedProject"]) == scope.Project && !string.IsNullOrEmpty(module) && selection != null)
                        {
                            var source = ReadWorkflow("read_module", scope.Project, module);
                            location = new[] { new ChatAttachment { Label = "Ouvrir " + module + " L" + selection["StartLine"], Text = diagnostic,
                                Project = scope.Project, Module = module, Sha256 = Convert.ToString(source["Sha256"]), StartLine = Convert.ToInt32(selection["StartLine"]) } };
                        }
                    } catch (Exception ex) { diagnostic += "\nLocalisation indisponible : " + ex.Message; }
                }
                AddEntry(new ChatEntry { Speaker = "Vérification", Text = diagnostic, Attachments = location });
                SetStatus(compiled ? "Vérification VBA terminée" : "Erreur de compilation · voir le diagnostic dans le chat");
            }
            catch (Exception ex) { AddEntry(new ChatEntry { Speaker = "Vérification", Text = "Compilation non vérifiée : " + ex.Message }); SetStatus("Compilation non vérifiée"); }
        }

        private void ExportCurrentChat()
        {
            if (currentSession == null || busy) return;
            SaveCurrentSession();
            using (var dialog = new System.Windows.Forms.SaveFileDialog { Filter = "Markdown (*.md)|*.md", FileName = "conversation-vba.md" })
            {
                if (dialog.ShowDialog(this) != System.Windows.Forms.DialogResult.OK) return;
                try { System.IO.File.WriteAllText(dialog.FileName, ChatHistory.Export(currentSession), new UTF8Encoding(false)); SetStatus("Conversation exportée"); }
                catch (Exception ex) { SetStatus("Export impossible : " + ex.Message); }
            }
        }

        private void ForkChat(ChatEntry lastEntry)
        {
            if (busy || currentSession == null) return;
            int index = transcriptEntries.IndexOf(lastEntry);
            if (index < 0) return;
            SaveCurrentSession();
            var entries = json.Deserialize<List<ChatEntry>>(json.Serialize(transcriptEntries.Take(index + 1)));
            // A branch is conversational context, never a second owner of rollback controls.
            foreach (var entry in entries) if (entry.Change != null) { entry.Text = entry.Change.Label + "\n" + entry.Change.Diff; entry.Change = null; }
            var fork = new ChatSessionState { Scope = currentSession.Scope, Title = currentSession.Title + " · branche",
                Provider = currentSession.Provider, Model = currentSession.Model, Effort = currentSession.Effort, Mode = currentSession.Mode, Entries = entries };
            var history = new List<object> { new { role = "system", content = LlmVbeContext.DeveloperInstructions } };
            foreach (var entry in entries.Where(x => x.Speaker == "Vous" || x.Speaker == "Assistant"))
                history.Add(new { role = entry.Speaker == "Vous" ? "user" : "assistant", content = entry.Text });
            fork.MessagesJson = json.Serialize(history);
            fork.ResumeContext = ChatHistory.Export(fork);
            scopeSessions.Insert(0, fork); ActivateSession(fork, false); SaveCurrentSession();
        }

        private void NavigateAttachment(ChatAttachment attachment)
        {
            try
            {
                EnsureCurrentScope();
                var result = scopeSession.Execute(new Request { Command = "select_code", Project = attachment.Project,
                    Module = attachment.Module, StartLine = Math.Max(1, attachment.StartLine), ExpectedSha256 = attachment.Sha256 });
                if (!result.Ok) SetStatus(result.Error);
            }
            catch (Exception ex) { SetStatus(ex.Message); }
        }

        private void RollbackIntervention(CodeChange change, int? hunk, bool entireTurn)
        {
            if (busy || tools == null) return;
            try
            {
                EnsureCurrentScope();
                var targets = entireTurn && !string.IsNullOrEmpty(change.TurnId) ? codeChanges.Where(x => x.TurnId == change.TurnId).ToArray() : new[] { change };
                var result = tools.RestoreChanges(targets, hunk);
                AddTranscriptMessage(result.Ok ? "Code restauré" : "Erreur", result.Ok ? "Annulation effectuée. Les autres modifications du module ont été conservées." : result.Error);
                RefreshCodeChangeCards(); SaveCurrentSession();
            }
            catch (Exception ex) { SetStatus(ex.Message); }
        }
    }
}
