using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

namespace CodexVBE
{
    internal sealed partial class ChatWindow
    {
        private ChatSessionStore sessionStore;
        private ChatSessionState currentSession;
        private bool loadingSession;
        private bool storageFailed;
        private VbeSession scopeSession;
        private readonly List<ChatSessionState> scopeSessions = new List<ChatSessionState>();
        private readonly Dictionary<string, List<ChatSessionState>> cachedScopes = new Dictionary<string, List<ChatSessionState>>();
        private DispatcherTimer saveTimer;
        private ComboBox scopePicker;
        private ListBox sessionList;
        private TextBox historySearch;
        private Border historyPanel;
        private TextBlock sessionTitle;
        private TextBox chatTitleEditor;
        private CheckBox showArchived;
        private TextBox memoryEditor;
        private CheckBox attachMemory;
        private string projectMemory = "";

        private sealed class MacroScope
        {
            public string Key;
            public string Label;
            public string Project;
            public override string ToString() { return Label; }
        }

        private void InitializeSessions(VbeSession session)
        {
            scopeSession = session;
            saveTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(700) };
            saveTimer.Tick += (s, e) => { saveTimer.Stop(); SaveCurrentSession(); };
            try
            {
                sessionStore = new ChatSessionStore(Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "CodexVBE", "chat.db"));
            }
            catch (Exception ex) { storageFailed = true; SetStatus("Historique non enregistré : " + ex.Message); }
            var projects = session.Execute(new Request { Command = "list_projects" });
            if (projects.Ok)
            {
                var values = json.DeserializeObject(json.Serialize(projects.Data)) as object[];
                if (values != null) foreach (var raw in values)
                {
                    var project = raw as IDictionary<string, object>;
                    if (project == null) continue;
                    string name = Convert.ToString(project["Name"]);
                    string path = Convert.ToString(project["FileName"]);
                    bool saved = !string.IsNullOrWhiteSpace(path) && Path.IsPathRooted(path);
                    scopePicker.Items.Add(new MacroScope {
                        Project = name,
                        Key = saved ? Path.GetFullPath(path).ToUpperInvariant() : "temporary:" + Guid.NewGuid().ToString("N"),
                        Label = name + " · " + (saved ? Path.GetFileName(path) : "document non enregistré")
                    });
                }
            }
            scopePicker.SelectionChanged += (s, e) => ChangeScope();
            sessionList.SelectionChanged += (s, e) => {
                var selected = sessionList.SelectedItem as ChatSessionState;
                if (!loadingSession && selected != null && selected != currentSession && !busy) ActivateSession(selected);
            };
            historySearch.TextChanged += (s, e) => RefreshHistory();
            if (scopePicker.Items.Count > 0)
            {
                int selected = 0;
                var first = (MacroScope)scopePicker.Items[0];
                var state = session.Execute(new Request { Command = "debug_state", Project = first.Project });
                if (state.Ok)
                {
                    var data = json.DeserializeObject(json.Serialize(state.Data)) as IDictionary<string, object>;
                    object active;
                    if (data != null && data.TryGetValue("SelectedProject", out active))
                        for (int i = 0; i < scopePicker.Items.Count; i++)
                            if (((MacroScope)scopePicker.Items[i]).Project == Convert.ToString(active)) selected = i;
                }
                scopePicker.SelectedIndex = selected;
            }
            else
            {
                send.IsEnabled = false;
                SetStatus(projects.Ok ? "Ouvrez un projet VBA pour démarrer une conversation." : projects.Error);
            }
        }

        private void ChangeScope()
        {
            if (busy || loadingSession) return;
            SaveCurrentSession();
            var scope = scopePicker.SelectedItem as MacroScope;
            if (scope == null) return;
            if (currentSession != null) cachedScopes[currentSession.Scope] = scopeSessions.ToList();
            scopeSessions.Clear();
            projectMemory = "";
            try {
                List<ChatSessionState> cached;
                if (cachedScopes.TryGetValue(scope.Key, out cached)) scopeSessions.AddRange(cached);
                else if (sessionStore != null) scopeSessions.AddRange(sessionStore.List(scope.Key));
                if (sessionStore != null) projectMemory = sessionStore.ReadMemory(scope.Key);
            }
            catch (Exception ex) { SetStatus("Historique indisponible : " + ex.Message); }
            memoryEditor.Text = projectMemory;
            attachMemory.IsChecked = false;
            if (!scopeSessions.Any(item => !item.Archived))
                scopeSessions.Add(new ChatSessionState { Scope = scope.Key, Provider = settings.ProviderName });
            ActivateSession(scopeSessions.First(item => !item.Archived), false);
        }

        private void ActivateSession(ChatSessionState session, bool savePrevious = true)
        {
            if (busy) return;
            if (savePrevious) SaveCurrentSession();
            loadingSession = true;
            try
            {
                codex?.Dispose(); codex = null;
                currentSession = session;
                modePicker.SelectedItem = session.Mode;
                if (tools != null) tools.Mode = session.Mode;
                draftAttachments.Clear();
                if (session.DraftAttachments != null) draftAttachments.AddRange(session.DraftAttachments);
                attachMemory.IsChecked = false;
                ClearTranscript(); codeChanges.Clear(); completedStreams.Clear(); streamedFinalText = null;
                messages.Clear();
                var saved = string.IsNullOrEmpty(session.MessagesJson) ? null :
                    json.DeserializeObject(session.MessagesJson) as object[];
                if (saved != null) messages.AddRange(saved);
                RepairInterruptedToolHistory();
                if (messages.Count == 0) messages.Add(new { role = "system", content = LlmVbeContext.DeveloperInstructions });
                foreach (var entry in session.Entries)
                {
                    if (entry.Change != null) codeChanges.Add(entry.Change);
                    if (!string.IsNullOrEmpty(entry.StreamId)) completedStreams.Add(entry.StreamId);
                    AddEntry(entry);
                }
                selectedReferences.Clear();
                if (session.DraftReferences != null) selectedReferences.AddRange(session.DraftReferences);
                prompt.Text = session.Draft ?? "";
                prompt.CaretIndex = prompt.Text.Length;
                HideReferences();
                RefreshContextChips();
                changes.Content = "Modifications · " + codeChanges.Count;
                changes.IsEnabled = codeChanges.Count > 0;
                RefreshCodeChangeCards();
                int provider = Array.FindIndex(LlmProvider.All, item => item.Name == session.Provider);
                providerPicker.SelectedIndex = provider < 0 ? 0 : provider;
                settings.ProviderName = ((LlmProvider)providerPicker.SelectedItem).Name;
                sessionTitle.Text = session.Title;
                chatTitleEditor.Text = session.Title;
                historyPanel.Visibility = Visibility.Collapsed;
                RefreshHistory();
                ShowWelcome();
            }
            finally { loadingSession = false; }
            _ = LoadModelsAsync();
        }

        private void RefreshHistory()
        {
            if (sessionList == null) return;
            bool previous = loadingSession;
            loadingSession = true;
            string query = historySearch.Text ?? "";
            sessionList.ItemsSource = scopeSessions.Where(x => (!x.Archived || showArchived.IsChecked == true) &&
                ChatHistory.Matches(x, query)).OrderByDescending(x => x.Pinned).ToArray();
            sessionList.SelectedItem = currentSession;
            loadingSession = previous;
        }

        private void ScheduleSessionSave()
        {
            if (loadingSession || saveTimer == null || currentSession == null) return;
            saveTimer.Stop(); saveTimer.Start();
        }

        private void SaveCurrentSession()
        {
            if (loadingSession || currentSession == null) return;
            saveTimer?.Stop();
            currentSession.Entries = transcriptEntries.ToList();
            currentSession.MessagesJson = json.Serialize(messages);
            currentSession.Draft = prompt.Text;
            currentSession.DraftAttachments = draftAttachments.ToArray();
            currentSession.DraftReferences = CurrentReferences(prompt.Text);
            if (codex != null && !string.IsNullOrEmpty(codex.ThreadId)) currentSession.CodexThreadId = codex.ThreadId;
            try { sessionStore?.Save(currentSession); }
            catch (Exception ex) { storageFailed = true; SetStatus("Historique non enregistré : " + ex.Message); }
        }

        private void NewSession(string provider = null)
        {
            if (busy || settings == null) return;
            var scope = scopePicker.SelectedItem as MacroScope;
            if (scope == null) return;
            SaveCurrentSession();
            var session = new ChatSessionState { Scope = scope.Key, Provider = provider ?? currentSession?.Provider ?? settings.ProviderName };
            scopeSessions.Insert(0, session);
            ActivateSession(session, false);
            SaveCurrentSession();
        }

        private void RenameFromQuestion(string text)
        {
            if (currentSession == null || currentSession.Title != "Nouvelle conversation" || transcriptEntries.Any(x => x.Speaker == "Vous")) return;
            currentSession.Title = text.Replace("\r", " ").Replace("\n", " ");
            if (currentSession.Title.Length > 64) currentSession.Title = currentSession.Title.Substring(0, 61) + "…";
            sessionTitle.Text = currentSession.Title;
            chatTitleEditor.Text = currentSession.Title;
            RefreshHistory();
        }

        private void RenameCurrentChat()
        {
            if (busy || currentSession == null || string.IsNullOrWhiteSpace(chatTitleEditor.Text)) return;
            currentSession.Title = chatTitleEditor.Text.Trim();
            if (currentSession.Title.Length > 120) currentSession.Title = currentSession.Title.Substring(0, 120);
            sessionTitle.Text = currentSession.Title;
            SaveCurrentSession(); RefreshHistory();
        }

        private void ToggleArchiveCurrentChat()
        {
            if (busy || currentSession == null) return;
            currentSession.Archived = !currentSession.Archived;
            SaveCurrentSession();
            if (currentSession.Archived) NewSession();
            else RefreshHistory();
        }

        private void SaveProjectMemory()
        {
            var scope = scopePicker.SelectedItem as MacroScope;
            if (scope == null || busy) return;
            if (sessionStore == null) { SetStatus("La mémoire nécessite un historique SQLite disponible."); return; }
            try
            {
                sessionStore.SaveMemory(scope.Key, memoryEditor.Text);
                projectMemory = memoryEditor.Text;
                RefreshContextChips();
                SetStatus("Mémoire du document enregistrée localement");
            }
            catch (Exception ex) { SetStatus("Mémoire non enregistrée : " + ex.Message); }
        }

        private void EnsureCurrentScope()
        {
            var scope = scopePicker.SelectedItem as MacroScope;
            if (scope == null || scopeSession == null) return;
            if (tools != null) tools.BoundProject = scope.Project;
            var response = scopeSession.Execute(new Request { Command = "list_projects" });
            if (!response.Ok) throw new InvalidOperationException(response.Error);
            var projects = json.DeserializeObject(json.Serialize(response.Data)) as object[];
            var matches = (projects ?? new object[0]).OfType<IDictionary<string, object>>()
                .Where(item => Convert.ToString(item["Name"]) == scope.Project).ToArray();
            if (matches.Length != 1) throw new InvalidOperationException("Le projet de cette conversation est fermé ou ambigu.");
            if (!scope.Key.StartsWith("temporary:", StringComparison.Ordinal))
            {
                string path = Convert.ToString(matches[0]["FileName"]);
                if (string.IsNullOrEmpty(path) || !string.Equals(Path.GetFullPath(path), scope.Key, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Le document VBA a changé. Rouvrez le chat pour choisir le document actuel.");
            }
        }

        private void RepairInterruptedToolHistory()
        {
            int userIndex = -1;
            var pending = new HashSet<string>();
            for (int i = 0; i < messages.Count; i++)
            {
                var item = messages[i] as IDictionary<string, object>;
                if (item == null || !item.ContainsKey("role")) continue;
                string role = Convert.ToString(item["role"]);
                if (role == "user") { userIndex = i; pending.Clear(); }
                object raw;
                if (role == "assistant" && item.TryGetValue("tool_calls", out raw) && raw is object[])
                    foreach (var value in (object[])raw)
                    {
                        var call = value as IDictionary<string, object>;
                        if (call != null && call.ContainsKey("id")) pending.Add(Convert.ToString(call["id"]));
                    }
                if (role == "tool" && item.ContainsKey("tool_call_id")) pending.Remove(Convert.ToString(item["tool_call_id"]));
            }
            if (pending.Count == 0 || userIndex < 0) return;
            messages.RemoveRange(userIndex + 1, messages.Count - userIndex - 1);
            messages.Add(new { role = "assistant", content = "La réponse précédente a été interrompue. Relire le code vivant avant de poursuivre les modifications." });
        }
    }
}
