using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

namespace VBAi
{

    /// <summary>Fenêtre de conversation avec gestion des sessions persistées.</summary>
    internal sealed partial class ChatWindow
    {

        /// <summary>Magasin SQLite partagé par les sessions et la mémoire de projet.</summary>
        private ChatSessionStore sessionStore;

        /// <summary>Background SQLite writer that receives immutable session snapshots and reports completion on the UI dispatcher.</summary>
        private ChatPersistenceWorker persistenceWorker;

        /// <summary>Serializer configured for session payloads up to 32 MiB.</summary>
        private readonly System.Web.Script.Serialization.JavaScriptSerializer persistenceJson =
            new System.Web.Script.Serialization.JavaScriptSerializer { MaxJsonLength = 32 * 1024 * 1024 };

        /// <summary>Session actuellement affichée.</summary>
        private ChatSessionState currentSession;

        /// <summary>Indique qu’un chargement de session est en cours et bloque les sauvegardes déclenchées par l’interface.</summary>
        private bool loadingSession;

        /// <summary>Prevents scope-selection handlers from starting another load while one is in progress.</summary>
        private bool loadingScope;

        /// <summary>Hides session-history controls when the SQLite store could not be initialized.</summary>
        private bool sessionViewUnavailable;

        /// <summary>Current asynchronous scope-load operation, observed before changing scope again.</summary>
        private System.Threading.Tasks.Task scopeLoad = System.Threading.Tasks.Task.CompletedTask;

        /// <summary>Worker read seam used to load detached session/memory snapshots from a database path and scope.</summary>
        internal Func<string, string, bool, System.Threading.Tasks.Task<ChatSessionStore.ScopeSnapshot>> ReadScope = ChatSessionStore.ReadScopeAsync;

        /// <summary>Indique qu’une erreur de stockage a empêché une sauvegarde.</summary>
        private bool storageFailed;

        /// <summary>Session VBE utilisée pour actualiser les portées de projet.</summary>
        private VbeSession scopeSession;

        /// <summary>Sessions chargées ou créées pour la portée sélectionnée.</summary>
        private readonly List<ChatSessionState> scopeSessions = new List<ChatSessionState>();

        /// <summary>Cache des sessions par portée de projet.</summary>
        private readonly Dictionary<string, List<ChatSessionState>> cachedScopes = new Dictionary<string, List<ChatSessionState>>();

        /// <summary>Unsaved notes keyed by temporary scope; these notes are never written to SQLite.</summary>
        private readonly Dictionary<string, string> transientMemory = new Dictionary<string, string>();

        /// <summary>Resolver for a borrowed live VBProject used only to capture unsaved-project identity.</summary>
        internal static Func<VbeSession, string, object> ReadScopeProject = (session, selector) => session.ProjectScopeSource(selector);

        /// <summary>Minuterie qui regroupe les sauvegardes rapprochées du brouillon.</summary>
        private DispatcherTimer saveTimer;

        /// <summary>Debounces session-history search input on the WPF dispatcher.</summary>
        private DispatcherTimer historySearchTimer;

        /// <summary>Minuterie de nouvelle tentative de découverte lorsque aucun projet n’est ouvert.</summary>
        private DispatcherTimer projectRetryTimer;

        /// <summary>Mémoire locale de la portée de projet courante.</summary>
        private string projectMemory = "";

        /// <summary>Décrit une portée de projet sélectionnable dans l’interface.</summary>
        private sealed class MacroScope
        {

            /// <summary>Clé stable utilisée pour les sessions et la mémoire.</summary>
            public string Key;

            /// <summary>Libellé visible dans le sélecteur de portée.</summary>
            public string Label;

            /// <summary>Chemin du projet enregistré ou son nom s’il est temporaire.</summary>
            public string Project;

            /// <summary>Nom du projet VBE.</summary>
            public string Name;

            /// <summary>Owned IUnknown identity lease for an unsaved live project; saved scopes use their canonical full path.</summary>
            public ScopeProjectLease Identity;

            /// <summary>First path observed after the unsaved project is saved, plus any local-history promotion error.</summary>
            public string FirstSavedPath, PromotionError;

            /// <summary>Blocks another scope-promotion attempt after SQLite could not prove the first transaction outcome.</summary>
            public bool PromotionBlocked;

            /// <summary>Retourne le libellé du sélecteur.</summary>
            /// <returns>Valeur de <see cref="Label"/>.</returns>
            public override string ToString() { return Label; }
        }

        /// <summary>Owns only its acquired IUnknown reference; shared project RCWs are never released.</summary>
        private sealed class ScopeProjectLease : IDisposable
        {

            /// <summary>Borrowed project RCW retained only to compare COM identity; this wrapper does not release it.</summary>
            private object project;

            /// <summary>One IUnknown reference acquired by this lease and released exactly once during disposal.</summary>
            private IntPtr unknown;

            /// <summary>Retains the project RCW for identity checks and acquires one owned IUnknown reference when it is a COM object.</summary>
            /// <param name="project">Live VBProject whose identity distinguishes an unsaved document from another project.</param>
            internal ScopeProjectLease(object project)
            {
                this.project = project;
                if (System.Runtime.InteropServices.Marshal.IsComObject(project))
                    unknown = System.Runtime.InteropServices.Marshal.GetIUnknownForObject(project);
            }

            /// <summary>Compares the leased project with a current project without transferring COM reference ownership.</summary>
            /// <param name="candidate">Current VBProject candidate.</param>
            /// <returns><see langword="false"/> for null, different, or failed COM identity comparisons.</returns>
            internal bool Matches(object candidate)
            {
                try { return project != null && VbeProjectHostPath.SameProject(project, candidate); }
                catch { return false; }
            }

            /// <summary>Clears the borrowed RCW and releases only the IUnknown reference acquired by this lease.</summary>
            public void Dispose()
            {
                project = null;
                if (unknown != IntPtr.Zero) { System.Runtime.InteropServices.Marshal.Release(unknown); unknown = IntPtr.Zero; }
            }
        }

        /// <summary>Attempts to resolve a live project and converts resolution failures to an unavailable identity.</summary>
        /// <param name="session">VBE session that owns project resolution.</param>
        /// <param name="selector">Project name or path currently associated with the scope.</param>
        /// <returns>Borrowed VBProject object, or null when resolution fails.</returns>
        private static object TryReadScopeProject(VbeSession session, string selector)
        {
            try { return ReadScopeProject(session, selector); }
            catch { return null; }
        }

        /// <summary>Captures a lease for one resolved unsaved project so renames do not silently change its conversation scope.</summary>
        /// <param name="session">VBE session used to resolve the live project.</param>
        /// <param name="selector">Current project selector.</param>
        /// <returns>Identity lease, or null when project resolution or COM identity acquisition fails.</returns>
        private static ScopeProjectLease CaptureScopeIdentity(VbeSession session, string selector)
        {
            object project = TryReadScopeProject(session, selector);
            if (project == null) return null;
            try { return new ScopeProjectLease(project); }
            catch { return null; }
        }

        /// <summary>Disposes scope identities for chat window.</summary>
        private void DisposeScopeIdentities()
        {
            foreach (var scope in scopePicker.Items.OfType<MacroScope>()) scope.Identity?.Dispose();
        }

        /// <summary>Promotes the exact unsaved project scope to its first saved path, atomically claiming an empty SQLite scope when storage is available.</summary>
        /// <param name="scope">Temporary scope whose live project identity must still match.</param>
        /// <param name="project">First canonical saved path observed for that project.</param>
        /// <param name="name">Current project name used to confirm the live scope identity.</param>
        /// <returns><see langword="false"/> when identity or destination ownership blocks promotion; true when memory was moved or retained locally.</returns>
        private bool PromoteScope(MacroScope scope, string project, string name)
        {
            string oldKey = scope.Key, newKey = project.ToUpperInvariant();
            if (scope.FirstSavedPath != null && !string.Equals(scope.FirstSavedPath, project, StringComparison.OrdinalIgnoreCase)) return false;
            scope.FirstSavedPath = project;
            scope.Project = project;
            // A previously managed saved scope must retain its separate conversation authority.
            if (cachedScopes.TryGetValue(newKey, out var occupied) && occupied.Count != 0) return false;
            if (sessionStore == null || scope.PromotionBlocked)
            {
                scope.PromotionError = scope.PromotionError ?? UiText.Get("Memory requires an available SQLite history store.");
                SetStatus(UiText.Get("Local history is kept in memory because storage is unavailable."));
                return true;
            }
            if (currentSession?.Scope == oldKey) SaveCurrentSession();
            cachedScopes.TryGetValue(oldKey, out var cached);
            var sessions = scopeSessions.Where(item => item.Scope == oldKey).Concat(cached ?? new List<ChatSessionState>())
                .GroupBy(item => item.Id).Select(group => group.First()).ToList();
            transientMemory.TryGetValue(oldKey, out var memory);
            Dictionary<string, string> versions;
            try
            {
                var rows = new List<ChatSessionStore.PromotionRow>();
                foreach (var item in sessions)
                {
                    var fields = (Dictionary<string, object>)persistenceJson.DeserializeObject(persistenceJson.Serialize(item));
                    fields[nameof(ChatSessionState.Scope)] = newKey;
                    rows.Add(new ChatSessionStore.PromotionRow { Id = item.Id, Title = item.Title, Payload = persistenceJson.Serialize(fields) });
                }
                // No temporary writes enter the worker; the initial claim owns this UI connection.
                versions = sessionStore.PromoteEmptyScope(newKey, rows, memory);
            }
            catch (Exception error)
            {
                // Keep the original snapshots and notes. Never replay an unverified initial claim.
                scope.PromotionBlocked = error is ChatSessionStore.PromotionOutcomeUnverifiedException; scope.PromotionError = error.Message;
                SetStatus(UiText.Get("Local history is kept in memory because storage is unavailable.") + " " + error.Message);
                return true;
            }
            if (versions == null) return false;
            foreach (var item in sessions) { item.Scope = newKey; item.StorageVersion = versions[item.Id]; }
            cachedScopes.Remove(oldKey); cachedScopes[newKey] = sessions;
            transientMemory.Remove(oldKey);
            scope.Key = newKey; scope.Name = name;
            scope.Label = name + " · " + Path.GetFileName(project);
            scope.Identity?.Dispose(); scope.Identity = null;
            return true;
        }

        /// <summary>Initialise le stockage, découvre les portées de projet et connecte les sélecteurs de session.</summary>
        /// <param name="session">Session VBE qui fournit la liste des projets.</param>
        private void InitializeSessions(VbeSession session)
        {
            scopeSession = session;
            InitializeContextMonitor(session);
            saveTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(700) };
            saveTimer.Tick += (s, e) => { saveTimer.Stop(); SaveCurrentSession(); };
            try
            {
                sessionStore = OpenHistory(HistoryPath());
                var dispatcher = Dispatcher.CurrentDispatcher;
                persistenceWorker = new ChatPersistenceWorker(sessionStore.DatabasePath, (snapshot, version, error) => {
                    if (error != null) LoadLog.Write("Chat history persistence failed: " + error.GetType().Name);
                    if (dispatcher.HasShutdownStarted) return;
                    dispatcher.BeginInvoke(new Action(() => {
                        if (runtimeDisposed || IsDisposed) return;
                        if (error != null)
                        {
                            storageFailed = true;
                            SetStatus(UiText.Get("History not saved: ") + error.Message);
                        }
                        else
                        {
                            var saved = scopeSessions.Concat(cachedScopes.Values.SelectMany(items => items))
                                .FirstOrDefault(item => item.WriterId == snapshot.Writer);
                            if (saved != null) saved.StorageVersion = version;
                        }
                    }));
                });
                if (Directory.Exists(persistenceWorker.RecoveryDirectory) && Directory.EnumerateFiles(persistenceWorker.RecoveryDirectory, "*.json").Any())
                {
                    storageFailed = true;
                    SetStatus(UiText.Get("Local history recovery copies are available. See the troubleshooting guide."));
                }
            }
            catch (Exception ex) { storageFailed = true; SetStatus(UiText.Get("History not saved: ") + ex.Message); }
            var projects = PopulateProjectScopes(session);
            scopePicker.SelectedIndexChanged += (s, e) => ChangeScope();
            sessionList.SelectedIndexChanged += (s, e) => {
                var selected = sessionList.SelectedItem as ChatSessionState;
                if (!loadingSession && selected != null && selected != currentSession && !busy) ActivateSession(selected);
                UpdateDeleteSessionButton();
            };
            historySearch.TextChanged += (s, e) => ScheduleHistorySearch();
            if (scopePicker.Items.Count > 0)
            {
                int selected = 0;
                var first = (MacroScope)scopePicker.Items[0];
                var state = ReadHost(session, new Request { Command = "debug_state", Project = first.Project });
                if (state.Ok)
                {
                    var data = json.DeserializeObject(json.Serialize(state.Data)) as IDictionary<string, object>;
                    object active;
                    if (data != null && data.TryGetValue("SelectedProject", out active))
                        for (int i = 0; i < scopePicker.Items.Count; i++)
                        {
                            var candidate = (MacroScope)scopePicker.Items[i];
                            string selectedPath = data.ContainsKey("SelectedHostPath") ? Convert.ToString(data["SelectedHostPath"]) : VbeProjectHostPath.AllowsLegacyPath && data.ContainsKey("SelectedProjectPath")
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

        /// <summary>Ajoute au sélecteur les projets accessibles, avec une clé de portée adaptée aux documents temporaires.</summary>
        /// <param name="session">Session VBE utilisée pour la commande de liste.</param>
        /// <returns>Réponse de la commande de découverte.</returns>
        private Response PopulateProjectScopes(VbeSession session)
        {
            var projects = ReadHost(session, new Request { Command = "list_projects" });
            if (!projects.Ok) return projects;
            var values = json.DeserializeObject(json.Serialize(projects.Data)) as object[];
            if (values == null) return projects;
            foreach (var raw in values)
            {
                var project = raw as IDictionary<string, object>;
                if (project == null) continue;
                string name = Convert.ToString(project["Name"]);
                string path = VbeProjectHostPath.FromFields(project);
                bool saved = !string.IsNullOrWhiteSpace(path) && Path.IsPathRooted(path);
                var scope = new MacroScope {
                    Project = saved ? Path.GetFullPath(path) : name,
                    Name = name,
                    Key = saved ? Path.GetFullPath(path).ToUpperInvariant() : "temporary:" + Guid.NewGuid().ToString("N"),
                    Label = name + " · " + (saved ? Path.GetFileName(path) : UiText.Get("unsaved document"))
                };
                if (!saved) scope.Identity = CaptureScopeIdentity(session, scope.Project);
                scopePicker.Items.Add(scope);
            }
            return projects;
        }

        /// <summary>Sauvegarde la session courante et charge les sessions et la mémoire de la nouvelle portée.</summary>
        private void ChangeScope()
        {
            if (loadingSession || runtimeDisposed || IsDisposed) return;
            scopeLoad = ChangeScopeAsync();
        }

        /// <summary>Saves the outgoing conversation, loads the selected scope off-thread, and publishes detached sessions/memory only if the selection is still current.</summary>
        /// <returns>Task that settles after loading, UI publication, or error recovery completes.</returns>
        private async System.Threading.Tasks.Task ChangeScopeAsync()
        {
            if (busy || loadingSession || runtimeDisposed || IsDisposed) return;
            var scope = scopePicker.SelectedItem as MacroScope;
            if (scope == null) return;
            SaveCurrentSession();
            if (currentSession != null) cachedScopes[currentSession.Scope] = scopeSessions.ToList();
            cachedScopes.TryGetValue(scope.Key, out var cached);
            string path = sessionStore?.DatabasePath;
            bool previousEnabled = rootLayout.Enabled;
            bool previousWaitCursor = UseWaitCursor;
            loadingScope = loadingSession = true;
            rootLayout.Enabled = false;
            UseWaitCursor = true;
            // Ignore catalog results belonging to the outgoing session.
            catalogueVersion++;
            try
            {
                ChatSessionStore.ScopeSnapshot snapshot = null;
                Exception failure = null;
                if (path != null && !ChatSessionStore.IsTransientScope(scope.Key))
                {
                    try { snapshot = await ReadScope(path, scope.Key, cached == null); }
                    catch (Exception error) { failure = error; }
                }
                if (runtimeDisposed || IsDisposed) return;
                // A programmatic scope change can still occur while controls are disabled.
                if (!ReferenceEquals(scopePicker.SelectedItem, scope)) return;
                scopeSessions.Clear();
                if (cached != null) scopeSessions.AddRange(cached);
                else if (snapshot?.Sessions != null) scopeSessions.AddRange(snapshot.Sessions);
                projectMemory = transientMemory.TryGetValue(scope.Key, out var temporaryNotes)
                    ? temporaryNotes : snapshot?.Memory ?? "";
                memoryEditor.Text = projectMemory;
                attachMemory.Checked = false;
                if (!scopeSessions.Any(item => !item.Archived))
                    scopeSessions.Add(new ChatSessionState { Scope = scope.Key, Provider = settings.ProviderName });
                loadingScope = loadingSession = false;
                ActivateSession(scopeSessions.First(item => !item.Archived), false);
                if (failure != null) SetStatus(UiText.Get("History unavailable: ") + failure.Message);
                else if (ChatSessionStore.IsTransientScope(scope.Key)) SetStatus(UiText.Get(scope.FirstSavedPath == null
                    ? "Temporary document: VBAi history and project notes stay in memory until the document is saved."
                    : "Local history is kept in memory because storage is unavailable."));
            }
            catch (Exception error)
            {
                if (!runtimeDisposed && !IsDisposed) SetStatus(UiText.Get("History unavailable: ") + error.Message);
            }
            finally
            {
                loadingScope = loadingSession = false;
                if (!runtimeDisposed && !IsDisposed)
                {
                    rootLayout.Enabled = previousEnabled;
                    UseWaitCursor = previousWaitCursor;
                    if (!ReferenceEquals(scopePicker.SelectedItem, scope)) ChangeScope();
                }
            }
        }

        /// <summary>Restaure les messages, pièces jointes, références, fournisseur et brouillon d’une session.</summary>
        /// <param name="session">Session à activer.</param>
        /// <param name="savePrevious">Indique s’il faut sauvegarder la session actuellement affichée avant le basculement.</param>
        private void ActivateSession(ChatSessionState session, bool savePrevious = true)
        {
            if (busy || loadingScope) return;
            if (savePrevious) SaveCurrentSession();
            loadingSession = true;
            try
            {
                codex?.Dispose(); codex = null;
                currentSession = session;
                modePicker.SelectedItem = session.Mode;
                if (tools != null) { tools.Mode = session.Mode; tools.ResetCatalog(); }
                draftAttachments.Clear();
                if (session.DraftAttachments != null) draftAttachments.AddRange(session.DraftAttachments);
                queuedDraftMemory = session.DraftCapturedMemory;
                attachMemory.Checked = !string.IsNullOrEmpty(queuedDraftMemory);
                ClearTranscript(); codeChanges.Clear(); completedStreams.Clear(); streamedFinalText = null;
                messages.Clear();
                var saved = string.IsNullOrEmpty(session.MessagesJson) ? null :
                    json.DeserializeObject(session.MessagesJson) as object[];
                if (saved != null) messages.AddRange(saved);
                RepairInterruptedToolHistory();
                if (messages.Count == 0) messages.Add(new { role = "system", content = LlmVbeContext.DeveloperInstructions });
                foreach (var entry in session.Entries)
                {
                    if (entry.Activity?.Status == "inProgress") entry.Activity.Status = "interrupted";
                    if (entry.Change != null) codeChanges.Add(entry.Change);
                    if (!string.IsNullOrEmpty(entry.StreamId)) completedStreams.Add(entry.StreamId);
                    AddEntry(entry);
                }
                MigrateProviderPrivacy();
                RefreshTranscriptWindow(Math.Max(0, transcriptEntries.Count - 80));
                FollowLatest();
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
                sessionViewUnavailable = false;
            }
            finally { loadingSession = false; }
            UpdateDeleteSessionButton();
            UpdateBudgetControls();
            RefreshPendingMessages();
            _ = LoadModelsAsync();
        }

        /// <summary>Filtre, trie et remplit la liste des sessions visibles, en conservant la sélection courante.</summary>
        private void RefreshHistory()
        {
            historySearchTimer?.Stop();
            if (runtimeDisposed || IsDisposed || sessionList == null) return;
            bool previous = loadingSession;
            loadingSession = true;
            sessionList.BeginUpdate();
            try
            {
                string query = historySearch.Text;
                sessionList.Items.Clear();
                sessionList.Items.AddRange(scopeSessions.Where(x => (!x.Archived || showArchived.Checked) &&
                    ChatHistory.Matches(x, query)).OrderByDescending(x => x.Pinned).ToArray());
                sessionList.SelectedItem = currentSession;
            }
            finally { sessionList.EndUpdate(); loadingSession = previous; }
            UpdateDeleteSessionButton();
        }

        /// <summary>Updates delete session button for chat window.</summary>
        private void UpdateDeleteSessionButton()
        {
            deleteSession.Enabled = !busy && !loadingSession && !loadingScope && !runtimeDisposed &&
                sessionList.SelectedItem is ChatSessionState selected && (ChatSessionStore.IsTransientScope(selected.Scope) || sessionStore != null);
        }

        /// <summary>Deletes a confirmed local history entry after its pending writes have finished.</summary>
        /// <returns>task produced by the operation for delete selected session async on chat window.</returns>
        private async System.Threading.Tasks.Task DeleteSelectedSessionAsync()
        {
            var selected = sessionList.SelectedItem as ChatSessionState;
            if (busy || loadingSession || loadingScope || runtimeDisposed || IsDisposed || selected == null ||
                (!ChatSessionStore.IsTransientScope(selected.Scope) && sessionStore == null)) return;
            if (ShowNotice(this, string.Format(UiText.Get("Delete conversation \"{0}\" from local history? This cannot be undone."), selected.DisplayTitle),
                UiText.Get("Delete conversation"), System.Windows.Forms.MessageBoxButtons.YesNo,
                System.Windows.Forms.MessageBoxIcon.Warning) != System.Windows.Forms.DialogResult.Yes) return;
            if (busy || loadingSession || loadingScope || runtimeDisposed || IsDisposed || sessionList.SelectedItem != selected) return;
            bool previousEnabled = rootLayout.Enabled;
            bool previousWaitCursor = UseWaitCursor;
            var selectedScope = scopePicker.SelectedItem;
            loadingScope = loadingSession = true;
            rootLayout.Enabled = false;
            UseWaitCursor = true;
            saveTimer?.Stop();
            UpdateDeleteSessionButton();
            bool deletedFromStore = false;
            try
            {
                Exception recoveryWarning = null;
                if (!ChatSessionStore.IsTransientScope(selected.Scope))
                {
                    if (persistenceWorker != null)
                        recoveryWarning = await persistenceWorker.DeleteAsync(new ChatPersistenceWorker.Snapshot(selected, null));
                    else sessionStore.Delete(selected.Id, selected.Scope, selected.StorageVersion);
                }
                deletedFromStore = true;
                if (runtimeDisposed || IsDisposed) return;
                scopeSessions.RemoveAll(item => item.Id == selected.Id && item.Scope == selected.Scope);
                foreach (var cached in cachedScopes.Values)
                    cached.RemoveAll(item => item.Id == selected.Id && item.Scope == selected.Scope);
                loadingScope = loadingSession = false;
                if (currentSession == selected)
                {
                    currentSession = null; // Never save the deleted conversation while activating its replacement.
                    var replacement = scopeSessions.FirstOrDefault(item => !item.Archived);
                    if (replacement == null)
                    {
                        replacement = new ChatSessionState { Scope = selected.Scope, Provider = settings.ProviderName };
                        scopeSessions.Insert(0, replacement);
                    }
                    ActivateSession(replacement, false);
                    SaveCurrentSession();
                }
                else RefreshHistory();
                historyPanel.Visible = true;
                historyPanel.BringToFront();
                SetStatus(UiText.Get("Conversation deleted from local history") + (recoveryWarning == null ? "" :
                    " · " + UiText.Get("Local history recovery copies are available. See the troubleshooting guide.")));
            }
            catch (Exception error)
            {
                if (!runtimeDisposed && !IsDisposed)
                {
                    if (deletedFromStore)
                    {
                        // Storage succeeded even if rebuilding a different conversation failed.
                        currentSession = null;
                        sessionViewUnavailable = true;
                        send.Enabled = false;
                        scopeSessions.RemoveAll(item => item.Id == selected.Id && item.Scope == selected.Scope);
                        foreach (var cached in cachedScopes.Values)
                            cached.RemoveAll(item => item.Id == selected.Id && item.Scope == selected.Scope);
                        try { RefreshHistory(); }
                        catch (Exception refreshError) { LoadLog.Write("Deleted conversation history refresh failed: " + refreshError.GetType().Name); }
                        SetStatus(UiText.Get("Conversation deleted from local history") + " · " + UiText.Get("History unavailable: ") + error.Message);
                    }
                    else SetStatus(UiText.Get("Conversation not deleted: ") + error.Message);
                }
            }
            finally
            {
                loadingScope = loadingSession = false;
                if (!runtimeDisposed && !IsDisposed)
                {
                    rootLayout.Enabled = previousEnabled;
                    UseWaitCursor = previousWaitCursor;
                    UpdateDeleteSessionButton();
                    if (!ReferenceEquals(scopePicker.SelectedItem, selectedScope)) ChangeScope();
                }
            }
        }

        /// <summary>Coalesces search edits before scanning and rebinding the session history.</summary>
        private void ScheduleHistorySearch()
        {
            if (runtimeDisposed || IsDisposed) return;
            if (historySearchTimer == null)
            {
                historySearchTimer = new DispatcherTimer(DispatcherPriority.Background) {
                    Interval = TimeSpan.FromMilliseconds(200)
                };
                historySearchTimer.Tick += (sender, args) => RefreshHistory();
            }
            historySearchTimer.Stop();
            historySearchTimer.Start();
        }

        /// <summary>Programme une sauvegarde différée du brouillon lorsque le chargement ne bloque pas les modifications.</summary>
        private void ScheduleSessionSave()
        {
            if (loadingSession || saveTimer == null || currentSession == null) return;
            saveTimer.Stop(); saveTimer.Start();
        }

        /// <summary>Copie l’état d’interface dans la session courante et demande sa persistance au magasin.</summary>
        private void SaveCurrentSession()
        {
            if (loadingSession || currentSession == null) return;
            try
            {
                saveTimer?.Stop();
                FlushStreamText();
                // Capture fallible values before replacing the last complete in-memory snapshot.
                string messageSnapshot = json.Serialize(messages);
                string draft = prompt.Text;
                var references = CurrentReferences(draft);
                currentSession.Entries = transcriptEntries.ToList();
                currentSession.MessagesJson = messageSnapshot;
                currentSession.Draft = draft;
                currentSession.DraftCapturedMemory = attachMemory.Checked ? queuedDraftMemory : null;
                currentSession.DraftAttachments = draftAttachments.ToArray();
                currentSession.DraftReferences = references;
                if (codex != null && !string.IsNullOrEmpty(codex.ThreadId)) currentSession.CodexThreadId = codex.ThreadId;
                if (ChatSessionStore.IsTransientScope(currentSession.Scope)) return;
                // Only immutable strings cross this boundary; mutable activity/card models stay on the UI thread.
                if (persistenceWorker != null)
                    persistenceWorker.Enqueue(new ChatPersistenceWorker.Snapshot(currentSession, persistenceJson.Serialize(currentSession)));
                else sessionStore?.Save(currentSession);
            }
            catch (Exception ex)
            {
                if (ChatSessionStore.IsTransientScope(currentSession.Scope))
                {
                    LoadLog.Write("Temporary conversation snapshot failed without creating disk recovery: " + ex.GetType().Name);
                    try { SetStatus(UiText.Get((scopePicker.SelectedItem as MacroScope)?.FirstSavedPath == null
                        ? "Temporary document: VBAi history and project notes stay in memory until the document is saved."
                        : "Local history is kept in memory because storage is unavailable.")); }
                    catch { }
                    return;
                }
                storageFailed = true;
                LoadLog.Write("Chat history save failed: " + ex.Message);
                try
                {
                    string recovery = (sessionStore?.DatabasePath ?? HistoryPath()) + ".recovery";
                    // Preserve plain text even when an activity/card cannot be serialized. No tool replay.
                    string payload = new System.Web.Script.Serialization.JavaScriptSerializer { MaxJsonLength = 64 * 1024 * 1024 }
                        .Serialize(new { FormatVersion = 1, RecoveryKind = "draft-text", currentSession.Id, currentSession.Scope,
                            Draft = prompt.Text, Entries = transcriptEntries.Select(entry => new { entry.Speaker, entry.Text }).ToArray() });
                    UpdatePaths.WriteAtomic(Path.Combine(recovery, "capture-" + Guid.NewGuid().ToString("N") + ".json"), payload);
                }
                catch (Exception recoveryError) { LoadLog.Write("Chat emergency recovery failed: " + recoveryError.GetType().Name); }
                // Shutdown must continue even if an already-disposed status surface cannot be updated.
                try { SetStatus(UiText.Get("History not saved: ") + ex.Message); }
                catch (Exception statusError) { LoadLog.Write("Chat save status unavailable: " + statusError.Message); }
            }
        }

        /// <summary>Crée une session dans la portée courante et l’active avec le fournisseur indiqué ou celui des paramètres.</summary>
        /// <param name="provider">Fournisseur initial facultatif.</param>
        private void NewSession(string provider = null)
        {
            if (busy || loadingScope || settings == null) return;
            var scope = scopePicker.SelectedItem as MacroScope;
            if (scope == null) return;
            SaveCurrentSession();
            var session = new ChatSessionState { Scope = scope.Key, Provider = provider ?? settings.ProviderName };
            scopeSessions.Insert(0, session);
            ActivateSession(session, false);
            SaveCurrentSession();
        }

        /// <summary>Utilise la première question comme titre si la session conserve encore son titre par défaut.</summary>
        /// <param name="text">Texte de la question, réduit à une ligne et tronqué à 64 caractères.</param>
        private void RenameFromQuestion(string text)
        {
            if (currentSession == null || currentSession.Title != "Nouvelle conversation" || transcriptEntries.Any(x => x.Speaker == "Vous")) return;
            currentSession.Title = text.Replace("\r", " ").Replace("\n", " ");
            if (currentSession.Title.Length > 64) currentSession.Title = currentSession.Title.Substring(0, 61) + "…";
            sessionTitle.Text = currentSession.Title;
            chatTitleEditor.Text = currentSession.Title;
            RefreshHistory();
        }

        /// <summary>Applique le titre saisi à la session, le limite à 120 caractères et sauvegarde l’état.</summary>
        private void RenameCurrentChat()
        {
            if (busy || loadingScope || currentSession == null || string.IsNullOrWhiteSpace(chatTitleEditor.Text)) return;
            currentSession.Title = chatTitleEditor.Text.Trim();
            if (currentSession.Title.Length > 120) currentSession.Title = currentSession.Title.Substring(0, 120);
            sessionTitle.Text = currentSession.Title;
            SaveCurrentSession(); RefreshHistory();
        }

        /// <summary>Inverse l’état archivé de la session courante et ouvre une session neuve si elle vient d’être archivée.</summary>
        private void ToggleArchiveCurrentChat()
        {
            if (busy || loadingScope || currentSession == null) return;
            currentSession.Archived = !currentSession.Archived;
            SaveCurrentSession();
            if (currentSession.Archived) NewSession();
            else RefreshHistory();
        }

        /// <summary>Enregistre localement le texte de mémoire de la portée courante et actualise le contexte affiché.</summary>
        private void SaveProjectMemory()
        {
            var scope = scopePicker.SelectedItem as MacroScope;
            if (scope == null || busy || loadingScope) return;
            if (ChatSessionStore.IsTransientScope(scope.Key))
            {
                transientMemory[scope.Key] = projectMemory = memoryEditor.Text;
                RefreshContextChips();
                SetStatus(UiText.Get(scope.FirstSavedPath == null
                    ? "Temporary document: VBAi history and project notes stay in memory until the document is saved."
                    : "Local history is kept in memory because storage is unavailable."));
                return;
            }
            if (sessionStore == null) { SetStatus(UiText.Get("Memory requires an available SQLite history store.")); return; }
            try
            {
                sessionStore.SaveMemory(scope.Key, memoryEditor.Text);
                transientMemory.Remove(scope.Key);
                projectMemory = memoryEditor.Text;
                RefreshContextChips();
                SetStatus(UiText.Get("Document memory saved locally"));
            }
            catch (Exception ex) { SetStatus(UiText.Get("Memory not saved: ") + ex.Message); }
        }

        /// <summary>Validates cached conversation selection and refreshes tool bindings without host reads.</summary>
        /// <exception cref="InvalidOperationException">History is unavailable/loading or the selected scope is missing.</exception>
        /// <returns>macro scope produced by the operation for ensure cached scope on chat window.</returns>
        private MacroScope EnsureCachedScope()
        {
            if (sessionViewUnavailable) throw new InvalidOperationException(UiText.Get("History unavailable: ").TrimEnd());
            if (loadingScope) throw new InvalidOperationException("Conversation history is still loading.");
            var scope = scopePicker.SelectedItem as MacroScope;
            if (scopeSession == null) return null;
            if (scope == null) throw new InvalidOperationException(UiText.Get("The project for this conversation is closed or ambiguous."));
            if (tools != null)
            {
                tools.BoundProject = scope.Project;
                tools.SetReadAccess(currentSession?.ReadProjectGrants, currentSession?.SharedContextReadAllowed ?? false);
            }
            return scope;
        }

        /// <summary>Validates the selected conversation scope and its current host project.</summary>
        private void EnsureCurrentScope()
        {
            var scope = EnsureCachedScope();
            if (scopeSession == null) return;
            var response = ReadHost(scopeSession, new Request { Command = "list_projects" });
            if (!response.Ok) throw new InvalidOperationException(response.Error);
            var projects = json.DeserializeObject(json.Serialize(response.Data)) as object[];
            var matches = (projects ?? new object[0]).OfType<IDictionary<string, object>>()
                .Where(item => !Path.IsPathRooted(scope.Project)
                    ? Convert.ToString(item["Name"]) == scope.Project
                    : string.Equals(VbeProjectHostPath.FromFields(item), scope.Project, StringComparison.OrdinalIgnoreCase)).ToArray();
            if (matches.Length != 1) throw new InvalidOperationException(UiText.Get("The project for this conversation is closed or ambiguous."));
            if (ChatSessionStore.IsTransientScope(scope.Key) && (scope.Identity == null ||
                !scope.Identity.Matches(TryReadScopeProject(scopeSession, scope.Project))))
                throw new InvalidOperationException(UiText.Get("The project for this conversation is closed or ambiguous."));
            if (!scope.Key.StartsWith("temporary:", StringComparison.Ordinal))
            {
                string path = VbeProjectHostPath.FromFields(matches[0]);
                if (string.IsNullOrEmpty(path) || !string.Equals(Path.GetFullPath(path), scope.Key, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException(UiText.Get("The VBA document has changed. Reopen the chat to choose the current document."));
            }
        }

        /// <summary>Retire les appels d’outils incomplets du dernier tour et ajoute un avis de réponse interrompue.</summary>
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
