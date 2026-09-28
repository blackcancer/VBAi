using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace CodexVBE
{
    /// <summary>Gère la collecte, vérification, prévisualisation et restauration du contexte de conversation.</summary>
    internal sealed partial class ChatWindow
    {
        /// <summary>Pièces jointes sélectionnées et en attente du prochain message.</summary>
        private readonly List<ChatAttachment> draftAttachments = new List<ChatAttachment>();
        /// <summary>Identifiant du tour de conversation courant.</summary>
        private string activeTurnId;
        /// <summary>Notifie l’hôte qu’un changement d’attachement de la fenêtre est demandé.</summary>
        public event Action DockRequested;
        /// <summary>Affiche une explication lorsque le VBE ne peut pas attacher la fenêtre.</summary>
        /// <param name="reason">Raison technique de l’échec de l’attachement.</param>
        public void ReportDockFailure(string reason) { SetStatus(UiText.Get("Docking unavailable: ") + reason + UiText.Get(" · check the COM control installation.")); }

        /// <summary>Résout les références actuelles, valide les sélections épinglées et la taille totale.</summary>
        /// <param name="question">Message dont les références et pièces jointes doivent être assemblées.</param>
        /// <returns>Pièces jointes valides et présentes dans le contexte du message.</returns>
        private ChatAttachment[] PrepareAttachments(string question)
        {
            var attachments = new List<ChatAttachment>();
            foreach (var reference in CurrentReferences(question))
            {
                tools?.RequireProjectRead(reference.Project);
                attachments.Add(new ChatAttachment { Label = reference.Token, Text = referenceIndex.Resolve(reference),
                    Project = reference.Project, Module = reference.Module, Sha256 = reference.Sha256, StartLine = reference.StartLine });
            }
            foreach (var attachment in draftAttachments)
            {
                if (!string.IsNullOrEmpty(attachment.Project)) tools?.RequireProjectRead(attachment.Project);
                if (!string.IsNullOrEmpty(attachment.EditorDocumentId))
                {
                    var document = scopeSession?.ModernEditor?.Invoke(false)?.Documents.FirstOrDefault(d => d.Id == attachment.EditorDocumentId);
                    if (document == null || EditorDocument.Hash(document.Text) != attachment.Sha256)
                        throw new InvalidOperationException(attachment.Label + UiText.Get(" is stale. Remove it and attach the current selection."));
                }
                else if (!string.IsNullOrEmpty(attachment.Module))
                {
                    var module = ReadWorkflow("read_module", attachment.Project, attachment.Module);
                    if (!string.Equals(Convert.ToString(module["Sha256"]), attachment.Sha256, StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException(attachment.Label + UiText.Get(" is stale. Remove it and attach the current selection."));
                }
                attachments.Add(attachment);
            }
            if (attachments.Sum(x => x.Text.Length) + question.Length > 48000)
                throw new InvalidOperationException(UiText.Get("Context exceeds 48,000 characters. Remove some items."));
            return attachments.ToArray();
        }

        /// <summary>Actualise l’aperçu des références et de la mémoire qui partiront au prochain envoi.</summary>
        private void RefreshContextPreview()
        {
            if (prompt == null) return;
            while (contextPreview.Controls.Count > 0) contextPreview.Controls[0].Dispose();
            try
            {
                EnsureCurrentScope();
                var attachments = PrepareAttachments(prompt.Text);
                foreach (var attachment in attachments)
                    AddContextPreview(attachment.Label + " · " + attachment.Text.Length + UiText.Get(" characters"), attachment.Text);
                int memorySize = attachMemory.Checked ? projectMemory.Length : 0;
                if (memorySize > 0) AddContextPreview(UiText.Get("Memory · ") + memorySize + UiText.Get(" characters"), projectMemory);
                AddContextPreview(UiText.Get("Explicit context"), (attachments.Sum(x => x.Text.Length) + memorySize) +
                    UiText.Get(" characters. This conversation's history also accompanies the request."));
            }
            catch (Exception ex) { AddContextPreview(UiText.Get("Context unavailable"), ex.Message); }
        }

        /// <summary>Ajoute un groupe de lecture seule pour un segment de contexte.</summary>
        /// <param name="title">Titre du segment d’aperçu.</param>
        /// <param name="text">Texte de lecture seule à afficher.</param>
        private void AddContextPreview(string title, string text)
        {
            var preview = new ChatContextPreviewView();
            preview.Width = Math.Max(200, contextPreview.ClientSize.Width - 26);
            preview.ShowContent(title, text);
            UiTheme.Apply(preview);
            contextPreview.Controls.Add(preview);
        }

        /// <summary>Exécute une commande VBE dans le périmètre de la conversation et désérialise son résultat.</summary>
        /// <param name="command">Nom de la commande protocole VBE.</param>
        /// <param name="project">Projet VBA ciblé, si nécessaire à la commande.</param>
        /// <param name="module">Module VBA ciblé, si nécessaire à la commande.</param>
        /// <returns>Dictionnaire de la réponse VBE sérialisée.</returns>
        private IDictionary<string, object> ReadWorkflow(string command, string project = null, string module = null)
        {
            if (scopeSession == null) throw new InvalidOperationException(UiText.Get("No VBE host connected."));
            var result = ReadHost(scopeSession, new Request { Command = command, Project = project, Module = module });
            if (!result.Ok) throw new InvalidOperationException(result.Error);
            return json.DeserializeObject(json.Serialize(result.Data)) as IDictionary<string, object> ?? throw new InvalidOperationException(UiText.Get("Unexpected VBE response."));
        }

        /// <summary>Capture la sélection du volet actif si elle appartient au projet lié.</summary>
        private void CaptureSelection()
        {
            if (busy) return;
            try
            {
                EnsureCurrentScope();
                var pane = ReadWorkflow("code_panes")["ActiveCodePane"] as IDictionary<string, object>;
                var properties = pane?["Properties"] as IDictionary<string, object>;
                if (properties == null || !properties.ContainsKey("Selection")) throw new InvalidOperationException(UiText.Get("Select code in the VBE."));
                string project = Convert.ToString(properties["Project"]), module = Convert.ToString(properties["Module"]);
                var scope = scopePicker.SelectedItem as MacroScope;
                string activeSelector = scope != null && !scope.Key.StartsWith("temporary:", StringComparison.Ordinal)
                    ? Convert.ToString(properties.ContainsKey("ProjectPath") ? properties["ProjectPath"] : null) : project;
                if (scope == null || !string.Equals(scope.Project, activeSelector, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException(UiText.Get("The selection belongs to another document. Choose its conversation."));
                var selection = (IDictionary<string, object>)properties["Selection"];
                int start = Convert.ToInt32(selection["StartLine"]), end = Convert.ToInt32(selection["EndLine"]);
                int startColumn = Convert.ToInt32(selection["StartColumn"]), endColumn = Convert.ToInt32(selection["EndColumn"]);
                var data = ReadWorkflow("read_module", scope.Project, module);
                var lines = CodeRollback.Lines(Convert.ToString(data["Code"]));
                if (start < 1 || end < start || end > lines.Length) throw new InvalidOperationException(UiText.Get("Empty or invalid code selection."));
                var selected = lines.Skip(start - 1).Take(end - start + 1).ToArray();
                if (start != end || startColumn != endColumn)
                {
                    selected[selected.Length - 1] = selected.Last().Substring(0, Math.Min(selected.Last().Length, Math.Max(0, endColumn - 1)));
                    selected[0] = selected[0].Substring(Math.Min(selected[0].Length, Math.Max(0, startColumn - 1)));
                }
                var attachment = new ChatAttachment { Label = project + "." + module + UiText.Get(" · selection L") + start + "–" + end,
                    Text = string.Join("\n", selected), Project = scope.Project, Module = module, StartLine = start, Sha256 = Convert.ToString(data["Sha256"]) };
                draftAttachments.RemoveAll(x => x.Label == attachment.Label); draftAttachments.Add(attachment);
                RefreshContextChips(); RefreshContextPreview(); ScheduleSessionSave();
                SetStatus(UiText.Get("Selection attached to the next message"));
            }
            catch (Exception ex) { SetStatus(UiText.Get("Selection: ") + ex.Message); }
        }

        /// <summary>Prépare un message rapide depuis une commande lancée dans l’éditeur.</summary>
        /// <param name="command">Nom de commande VBE à exécuter.</param>
        public void PrepareEditorAction(string command)
        {
            if (busy) { SetStatus(UiText.Get("Wait for the response to finish before preparing an action.")); return; }
            CaptureSelection();
            var action = ChatCommand.All.First(x => x.Token == command);
            modePicker.SelectedItem = action.Mode;
            prompt.Text = command + " "; prompt.CaretIndex = prompt.Text.Length; prompt.Focus();
        }

        internal void PrepareMonacoAction(string command, ChatAttachment attachment)
        {
            if (busy) { SetStatus(UiText.Get("Wait for the response to finish before preparing an action.")); return; }
            var action = ChatCommand.All.First(x => x.Token == command);
            var scope = scopePicker.Items.Cast<MacroScope>().FirstOrDefault(item => string.Equals(item.Project, attachment.Project, StringComparison.OrdinalIgnoreCase));
            if (scope == null) { SetStatus(UiText.Get("The selection belongs to another document. Choose its conversation.")); return; }
            scopePicker.SelectedItem = scope; EnsureCurrentScope();
            draftAttachments.RemoveAll(item => item.Label == attachment.Label); draftAttachments.Add(attachment);
            modePicker.SelectedItem = action.Mode;
            prompt.Text = command + " "; prompt.CaretIndex = prompt.Text.Length;
            RefreshContextChips(); RefreshContextPreview(); ScheduleSessionSave(); prompt.Focus();
        }

        /// <summary>Compile le projet pour recueillir un diagnostic, sans exécuter les macros.</summary>
        /// <returns>Tâche terminée après ajout du résultat de vérification au transcript.</returns>
        private async Task VerifyProjectAsync()
        {
            try
            {
                EnsureCurrentScope();
                var scope = scopePicker.SelectedItem as MacroScope;
                if (scope == null || tools == null) throw new InvalidOperationException(UiText.Get("No connected project."));
                SetStatus(UiText.Get("Compiling VBA…"));
                var response = json.Deserialize<Response>(await InvokeTool(tools, "compile_project", json.Serialize(new { Project = scope.Project, ExpectedMode = 2 })));
                if (response == null || !response.Ok) throw new InvalidOperationException(response?.Error ?? UiText.Get("Empty response."));
                var data = json.DeserializeObject(json.Serialize(response.Data)) as IDictionary<string, object>;
                bool compiled = data != null && data.ContainsKey("Compiled") && Convert.ToBoolean(data["Compiled"]);
                string diagnostic = compiled ? UiText.Get("Compilation finished: no native diagnostics observed. Macros were not executed.") : Convert.ToString(data?["Diagnostic"]);
                ChatAttachment[] location = null;
                if (!compiled)
                {
                    try {
                        var state = ReadWorkflow("debug_state", scope.Project);
                        string module = Convert.ToString(state["ActiveModule"]);
                        var selection = state["Selection"] as IDictionary<string, object>;
                        string selectedSelector = scope.Key.StartsWith("temporary:", StringComparison.Ordinal)
                            ? Convert.ToString(state["SelectedProject"]) : Convert.ToString(state["SelectedProjectPath"]);
                        if (string.Equals(selectedSelector, scope.Project, StringComparison.OrdinalIgnoreCase) &&
                            !string.IsNullOrEmpty(module) && selection != null)
                        {
                            var source = ReadWorkflow("read_module", scope.Project, module);
                            location = new[] { new ChatAttachment { Label = UiText.Get("Open ") + module + " L" + selection["StartLine"], Text = diagnostic,
                                Project = scope.Project, Module = module, Sha256 = Convert.ToString(source["Sha256"]), StartLine = Convert.ToInt32(selection["StartLine"]) } };
                        }
                    } catch (Exception ex) { diagnostic += UiText.Get("\nLocation unavailable: ") + ex.Message; }
                }
                AddEntry(new ChatEntry { Speaker = "Vérification", Text = diagnostic, Attachments = location });
                SetStatus(compiled ? UiText.Get("VBA verification complete") : UiText.Get("Compilation error · see the diagnostic in the chat"));
            }
            catch (Exception ex) { AddEntry(new ChatEntry { Speaker = "Vérification", Text = UiText.Get("Compilation not verified: ") + ex.Message }); SetStatus(UiText.Get("Compilation not verified")); }
        }

        /// <summary>Enregistre la session courante dans un fichier Markdown.</summary>
        private void ExportCurrentChat()
        {
            if (currentSession == null || busy) return;
            SaveCurrentSession();
            using (var dialog = new System.Windows.Forms.SaveFileDialog { Filter = "Markdown (*.md)|*.md", FileName = "conversation-vba.md" })
            {
                if (ShowSaveDialog(dialog,this) != System.Windows.Forms.DialogResult.OK) return;
                try { System.IO.File.WriteAllText(dialog.FileName, ChatHistory.Export(currentSession), new UTF8Encoding(false)); SetStatus(UiText.Get("Conversation exported")); }
                catch (Exception ex) { SetStatus(UiText.Get("Unable to export: ") + ex.Message); }
            }
        }

        /// <summary>Crée une branche de conversation sans partager les actions de rollback.</summary>
        /// <param name="lastEntry">Dernière entrée de transcript qui définit le point de branchement.</param>
        private void ForkChat(ChatEntry lastEntry)
        {
            if (busy || currentSession == null) return;
            int index = transcriptEntries.IndexOf(lastEntry);
            if (index < 0) return;
            int start = Math.Max(0, currentSession.ProviderHistoryStartIndex);
            if (index < start)
            { SetStatus(UiText.Get("This message predates the project privacy upgrade. Use current authorized references in a new conversation.")); return; }
            SaveCurrentSession();
            var entries = json.Deserialize<List<ChatEntry>>(json.Serialize(transcriptEntries.Skip(start).Take(index + 1 - start)));
            // A branch is conversational context, never a second owner of rollback controls.
            foreach (var entry in entries) if (entry.Change != null) { entry.Text = entry.Change.Label + "\n" + entry.Change.Diff; entry.Change = null; }
            var fork = new ChatSessionState { Scope = currentSession.Scope, Title = currentSession.Title + UiText.Get(" · branch"),
                Provider = currentSession.Provider, Model = currentSession.Model, Effort = currentSession.Effort, Mode = currentSession.Mode, Entries = entries,
                ReadProjectGrants = currentSession.ReadProjectGrants?.ToArray(), SharedContextReadAllowed = currentSession.SharedContextReadAllowed };
            var history = new List<object> { new { role = "system", content = LlmVbeContext.DeveloperInstructions } };
            foreach (var entry in entries.Where(x => x.Speaker == "Vous" || x.Speaker == "Assistant"))
                history.Add(new { role = entry.Speaker == "Vous" ? "user" : "assistant", content = entry.Text });
            fork.MessagesJson = json.Serialize(history);
            fork.ResumeContext = ChatHistory.Export(fork);
            scopeSessions.Insert(0, fork); ActivateSession(fork, false); SaveCurrentSession();
        }

        /// <summary>Sélectionne le code source correspondant à une pièce jointe dans le VBE.</summary>
        /// <param name="attachment">Référence ou sélection à ouvrir dans le VBE.</param>
        private async void NavigateAttachment(ChatAttachment attachment)
        {
            try
            {
                EnsureCurrentScope();
                if (!string.IsNullOrEmpty(attachment.EditorDocumentId))
                {
                    var editor = scopeSession?.ModernEditor?.Invoke(false);
                    var document = editor?.Documents.FirstOrDefault(d => d.Id == attachment.EditorDocumentId);
                    if (document == null) throw new InvalidOperationException(UiText.Get("Select code in the VBE."));
                    await editor.CaptureForTool();
                    if (EditorDocument.Hash(document.Text) != attachment.Sha256)
                        throw new InvalidOperationException(attachment.Label + UiText.Get(" is stale. Remove it and attach the current selection."));
                    scopeSession.ModernEditor(true);
                    await editor.OpenModule(document.Module);
                    await editor.Script("reveal", Math.Max(1, attachment.StartLine), 1);
                    return;
                }
                var result = ReadHost(scopeSession, new Request { Command = "select_code", Project = attachment.Project,
                    Module = attachment.Module, StartLine = Math.Max(1, attachment.StartLine), ExpectedSha256 = attachment.Sha256 });
                if (!result.Ok) SetStatus(result.Error);
            }
            catch (Exception ex) { SetStatus(ex.Message); }
        }

        /// <summary>Restaure un changement ou les changements du tour après vérification du périmètre.</summary>
        /// <param name="change">Changement enregistré à restaurer.</param>
        /// <param name="hunk">Index facultatif du bloc de diff à restaurer.</param>
        /// <param name="entireTurn">Indique si les changements du même tour doivent être restaurés ensemble.</param>
        private async void RollbackIntervention(CodeChange change, int? hunk, bool entireTurn)
        {
            if (busy || tools == null) return;
            try
            {
                EnsureCurrentScope();
                var targets = entireTurn && !string.IsNullOrEmpty(change.TurnId) ? codeChanges.Where(x => x.TurnId == change.TurnId).ToArray() : new[] { change };
                var result = await tools.RestoreChangesAsync(targets, hunk);
                AddTranscriptMessage(result.Ok ? UiText.Get("Code restored") : "Erreur", result.Ok ? UiText.Get("Undo complete. Other changes in the module were preserved.") : result.Error);
                RefreshCodeChangeCards(); SaveCurrentSession();
            }
            catch (Exception ex) { SetStatus(ex.Message); }
        }
    }
}
