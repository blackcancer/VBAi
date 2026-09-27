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
        private DispatcherTimer projectRetryTimer;
        private string projectMemory = "";

        private sealed class MacroScope
        {
            public string Key;
            public string Label;
            public string Project;
            public string Name;
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
            catch (Exception ex) { storageFailed = true; SetStatus(UiText.Get("History not saved: ") + ex.Message); }
            var projects = PopulateProjectScopes(session);
            scopePicker.SelectedIndexChanged += (s, e) => ChangeScope();
            sessionList.SelectedIndexChanged += (s, e) => {
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
                        {
                            var candidate = (MacroScope)scopePicker.Items[i];
                            string selectedPath = data.ContainsKey("SelectedProjectPath")
                                ? Convert.ToString(data["SelectedProjectPath"]) : null;
                            if (!candidate.Key.StartsWith("temporary:", StringComparison.Ordinal)
                                ? string.Equals(candidate.Project, selectedPath, StringComparison.OrdinalIgnoreCase)
                                : candidate.Name == Convert.ToString(active)) selected = i;
                        }
                }
                scopePicker.SelectedIndex = selected;
            }
            else
            {
                send.Enabled = false;
                SetStatus(projects.Ok ? UiText.Get("Open a VBA project to start a conversation.") : projects.Error);
                projectRetryTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(2) };
                projectRetryTimer.Tick += (s, e) => {
                    if (IsDisposed) { projectRetryTimer.Stop(); return; }
                    try { if (scopePicker.Items.Count == 0) PopulateProjectScopes(session); }
                    catch (Exception ex) { LoadLog.Write("Chat project discovery retry failed: " + ex.Message); return; }
                    if (scopePicker.Items.Count == 0) return;
                    projectRetryTimer.Stop();
                    send.Enabled = true;
                    scopePicker.SelectedIndex = 0;
                };
                projectRetryTimer.Start();
            }
        }

        private Response PopulateProjectScopes(VbeSession session)
        {
            var projects = session.Execute(new Request { Command = "list_projects" });
            if (!projects.Ok) return projects;
            var values = json.DeserializeObject(json.Serialize(projects.Data)) as object[];
            if (values == null) return projects;
            foreach (var raw in values)
            {
                var project = raw as IDictionary<string, object>;
                if (project == null) continue;
                string name = Convert.ToString(project["Name"]);
                string path = Convert.ToString(project["FileName"]);
                bool saved = !string.IsNullOrWhiteSpace(path) && Path.IsPathRooted(path);
                scopePicker.Items.Add(new MacroScope {
                    Project = saved ? Path.GetFullPath(path) : name,
                    Name = name,
                    Key = saved ? Path.GetFullPath(path).ToUpperInvariant() : "temporary:" + Guid.NewGuid().ToString("N"),
                    Label = name + " · " + (saved ? Path.GetFileName(path) : UiText.Get("unsaved document"))
                });
            }
            return projects;
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
            catch (Exception ex) { SetStatus(UiText.Get("History unavailable: ") + ex.Message); }
            memoryEditor.Text = projectMemory;
            attachMemory.Checked = false;
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
                attachMemory.Checked = false;
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
                changes.Text = UiText.Get("Changes · ") + codeChanges.Count;
                changes.Enabled = codeChanges.Count > 0;
                RefreshCodeChangeCards();
                int provider = Array.FindIndex(LlmProvider.All, item => item.Name == session.Provider);
                providerPicker.SelectedIndex = provider < 0 ? 0 : provider;
                sessionTitle.Text = session.DisplayTitle;
                chatTitleEditor.Text = session.DisplayTitle;
                historyPanel.Visible = false;
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
            sessionList.BeginUpdate();
            try
            {
                string query = historySearch.Text ?? "";
                sessionList.Items.Clear();
                sessionList.Items.AddRange(scopeSessions.Where(x => (!x.Archived || showArchived.Checked) &&
                    ChatHistory.Matches(x, query)).OrderByDescending(x => x.Pinned).ToArray());
                sessionList.SelectedItem = currentSession;
            }
            finally { sessionList.EndUpdate(); loadingSession = previous; }
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
            catch (Exception ex) { storageFailed = true; SetStatus(UiText.Get("History not saved: ") + ex.Message); }
        }

        private void NewSession(string provider = null)
        {
            if (busy || settings == null) return;
            var scope = scopePicker.SelectedItem as MacroScope;
            if (scope == null) return;
            SaveCurrentSession();
            var session = new ChatSessionState { Scope = scope.Key, Provider = provider ?? settings.ProviderName };
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
            if (sessionStore == null) { SetStatus(UiText.Get("Memory requires an available SQLite history store.")); return; }
            try
            {
                sessionStore.SaveMemory(scope.Key, memoryEditor.Text);
                projectMemory = memoryEditor.Text;
                RefreshContextChips();
                SetStatus(UiText.Get("Document memory saved locally"));
            }
            catch (Exception ex) { SetStatus(UiText.Get("Memory not saved: ") + ex.Message); }
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
                .Where(item => scope.Key.StartsWith("temporary:", StringComparison.Ordinal)
                    ? Convert.ToString(item["Name"]) == scope.Project
                    : string.Equals(Convert.ToString(item["FileName"]), scope.Project, StringComparison.OrdinalIgnoreCase)).ToArray();
            if (matches.Length != 1) throw new InvalidOperationException(UiText.Get("The project for this conversation is closed or ambiguous."));
            if (!scope.Key.StartsWith("temporary:", StringComparison.Ordinal))
            {
                string path = Convert.ToString(matches[0]["FileName"]);
                if (string.IsNullOrEmpty(path) || !string.Equals(Path.GetFullPath(path), scope.Key, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException(UiText.Get("The VBA document has changed. Reopen the chat to choose the current document."));
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
            messages.Add(new { role = "assistant", content = UiText.Get("The previous response was interrupted. Read the live code again before making further edits.") });
        }
    }
}
