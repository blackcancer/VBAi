namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.IO;
    using System.Net;
    using System.Net.Http;
    using System.Reflection;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Web.Script.Serialization;
    using System.Windows.Forms;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>Vérifie l’historique, la persistance locale et la réparation des conversations.</summary>
    public sealed partial class ChatWindowStateTests
    {
        [STATestMethod, TestCategory("Unit")]
        public void WordDocumentScopesStayDistinctAndRejectChangedPathWithoutTransferringGrants()
        {
            using (var runtime = new RuntimeScope())
            {
                string secondPath = @"C:\Owned\Second.docm";
                runtime.Host = r => Response.Success(r.Command == "list_projects" ? (object)new[] {
                    new { Name = "Project", FileName = @"C:\Temp\~WRL0001.tmp", HostPath = @"C:\Owned\First.docm" },
                    new { Name = "Project", FileName = @"C:\Temp\~WRL0002.tmp", HostPath = secondPath }
                } : new { SelectedProject = "Project", SelectedProjectPath = @"C:\Temp\~WRL0002.tmp", SelectedHostPath = secondPath });
                using (var window = LoadedWindow(runtime.Session))
                {
                    var scopes = Get<ComboBox>(window, "scopePicker");
                    Assert.AreEqual(2, scopes.Items.Count);
                    Assert.AreEqual(1, scopes.SelectedIndex);
                    StringAssert.Contains(scopes.SelectedItem.ToString(), "Second.docm");
                    Call(window, "EnsureCurrentScope");
                    var original = Get<ChatSessionState>(window, "currentSession");
                    secondPath = @"C:\Owned\Renamed.docm";
                    Assert.ThrowsException<TargetInvocationException>(() => Call(window, "EnsureCurrentScope"));
                    Call(window, "RefreshAvailableScopes", runtime.Session);
                    Assert.AreEqual(-1, scopes.SelectedIndex);
                    Assert.AreSame(original, Get<ChatSessionState>(window, "currentSession"));
                }
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void ScopeReadIgnoresStaleResultsAndKeepsActionsBlockedUntilLatestScopeLoads()
        {
            using (var runtime = new RuntimeScope())
            using (var store = new ChatSessionStore(Path.Combine(runtime.Root, "scope-reader.db")))
            using (var window = ReadyCodexWindow(new ChatSessionState { Scope = "temporary:A" }))
            {
                Set(window, "sessionStore", store);
                var original = Get<ChatSessionState>(window, "currentSession");
                Get<List<ChatSessionState>>(window, "scopeSessions").Add(original);
                var b = AddScope(window, "temporary:B");
                var c = AddScope(window, "temporary:C");
                var scopes = Get<ComboBox>(window, "scopePicker");
                var first = new TaskCompletionSource<ChatSessionStore.ScopeSnapshot>();
                var latest = new TaskCompletionSource<ChatSessionStore.ScopeSnapshot>();
                var requested = new List<string>();
                window.ReadScope = (path, scope, includeSessions) => {
                    requested.Add(scope);
                    return scope == "temporary:B" ? first.Task : latest.Task;
                };
                try
                {
                    scopes.SelectedItem = b;
                    Call(window, "ChangeScope");
                    Task staleLoad = Get<Task>(window, "scopeLoad");
                    Assert.IsTrue(Get<bool>(window, "loadingScope"));
                    Assert.IsFalse(Get<TableLayoutPanel>(window, "rootLayout").Enabled);
                    Call(window, "NewSession", (object)null);
                    Assert.AreSame(original, Get<ChatSessionState>(window, "currentSession"));
                    Assert.ThrowsException<TargetInvocationException>(() => Call(window, "EnsureCurrentScope"));
                    scopes.SelectedItem = c;
                    Call(window, "ChangeScope");
                    first.SetResult(new ChatSessionStore.ScopeSnapshot {
                        Sessions = new List<ChatSessionState> { new ChatSessionState { Scope = "temporary:B", Title = "Stale" } }, Memory = "stale memory"
                    });
                    CompleteOnSta(staleLoad);
                    Assert.AreSame(original, Get<ChatSessionState>(window, "currentSession"));
                    Assert.IsTrue(Get<bool>(window, "loadingScope"));
                    CollectionAssert.AreEqual(new[] { "temporary:B", "temporary:C" }, requested);
                    var expected = new ChatSessionState { Scope = "temporary:C", Title = "Current" };
                    latest.SetResult(new ChatSessionStore.ScopeSnapshot { Sessions = new List<ChatSessionState> { expected }, Memory = "current memory" });
                    CompleteScopeLoad(window);
                    Assert.AreSame(expected, Get<ChatSessionState>(window, "currentSession"));
                    Assert.AreEqual("current memory", Get<TextBox>(window, "memoryEditor").Text);
                    Assert.IsFalse(Get<bool>(window, "loadingScope"));
                    Assert.IsTrue(Get<TableLayoutPanel>(window, "rootLayout").Enabled);
                }
                finally
                {
                    first.TrySetResult(new ChatSessionStore.ScopeSnapshot());
                    latest.TrySetResult(new ChatSessionStore.ScopeSnapshot());
                    CompleteScopeLoad(window);
                }
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void ScopeReadCompletingAfterDisposalDoesNotActivateAConversation()
        {
            using (var runtime = new RuntimeScope())
            using (var store = new ChatSessionStore(Path.Combine(runtime.Root, "scope-close.db")))
            using (var window = ReadyCodexWindow(new ChatSessionState { Scope = "temporary:A" }))
            {
                Set(window, "sessionStore", store);
                var original = Get<ChatSessionState>(window, "currentSession");
                var target = AddScope(window, "temporary:B");
                Get<ComboBox>(window, "scopePicker").SelectedItem = target;
                var read = new TaskCompletionSource<ChatSessionStore.ScopeSnapshot>();
                window.ReadScope = (path, scope, includeSessions) => read.Task;
                Call(window, "ChangeScope");
                Task loading = Get<Task>(window, "scopeLoad");
                window.Dispose();
                read.SetResult(new ChatSessionStore.ScopeSnapshot {
                    Sessions = new List<ChatSessionState> { new ChatSessionState { Scope = "temporary:B" } }, Memory = "ignored"
                });
                CompleteOnSta(loading);
                Assert.AreSame(original, Get<ChatSessionState>(window, "currentSession"));
                Assert.IsTrue(window.IsDisposed);
                Assert.IsFalse(Get<bool>(window, "loadingScope"));
            }
        }

        /// <summary>Restores unfinished activities as interrupted while retaining their content and permitting new live activity.</summary>
        [STATestMethod]
        public void RestoredActivitiesAreInterruptedWithoutInventingResultsOrDuration()
        {
            using (var runtime = new RuntimeScope())
            using (var window = LoadedWindow(runtime.Session))
            {
                var session = new ChatSessionState { Provider = "missing-provider", Entries = new List<ChatEntry>() };
                foreach (var kind in new[] { "reasoning", "commandExecution" })
                    session.Entries.Add(new ChatEntry { Speaker = kind == "reasoning" ? "Réflexion" : "Outil", StreamId = kind,
                        Activity = new CodexAgentActivity { Id = kind, Kind = kind, Title = "Saved activity", Detail = "Original partial output", Status = "inProgress" } });
                Call(window, "ActivateSession", session, false);
                foreach (var entry in session.Entries)
                {
                    Assert.AreEqual("interrupted", entry.Activity.Status); Assert.AreEqual("Original partial output", entry.Activity.Detail);
                    Assert.IsNull(entry.Activity.DurationMs);
                    using (var host = (ChatDesignerHost)Call(window, "RenderActivityStep", entry))
                    {
                        var view = (ChatActivityStepView)host.View;
                        Assert.AreEqual(UiText.Get("Interrupted"), view.state.Text); Assert.IsFalse(view.section.Expanded);
                    }
                }
                using (var host = (ChatDesignerHost)Call(window, "RenderActivityGroup", session.Entries[0], session.Entries))
                    Assert.IsFalse(((ChatActivityGroupView)host.View).section.Expanded);
                Call(window, "ReceiveAgentActivity", new CodexAgentActivity { Id = "new", Kind = "reasoning", Title = "New activity", Detail = "Current output", Status = "inProgress" });
                var live = Get<Dictionary<string, ChatEntry>>(window, "liveEntries")["new"];
                using (var host = (ChatDesignerHost)Call(window, "RenderActivityStep", live)) Assert.IsTrue(((ChatActivityStepView)host.View).section.Expanded);
            }
        }

        /// <summary>Filtre les sessions selon le texte et l’état archivé, puis place les sessions épinglées en premier.</summary>
        [TestMethod]
        [STATestMethod]
        public void HistoryFiltersArchivedSessionsAndOrdersPinnedFirst()
        {
            using (var window = Surfaces())
            {
                var sessions = Get<List<ChatSessionState>>(window, "scopeSessions");
                var normal = new ChatSessionState
                {
                    Title = "Alpha"
                };
                var pinned = new ChatSessionState
                {
                    Title = "Alpha pinned",
                    Pinned = true
                };
                var archived = new ChatSessionState
                {
                    Title = "Alpha archived",
                    Archived = true
                };
                sessions.Add(normal);
                sessions.Add(pinned);
                sessions.Add(archived);
                Get<TextBox>(window, "historySearch").Text = "Alpha";
                Call(window, "RefreshHistory");
                var list = Get<ListBox>(window, "sessionList");
                Assert.AreEqual(2, list.Items.Count);
                Assert.AreSame(pinned, list.Items[0]);
                Get<CheckBox>(window, "showArchived").Checked = true;
                Call(window, "RefreshHistory");
                Assert.AreEqual(3, list.Items.Count);
            }
        }

        /// <summary>Sauvegarde le titre, le brouillon et les messages de session sans stockage externe.</summary>
        [TestMethod]
        [STATestMethod]
        public void SessionTitleAndDraftAreSavedWithoutExternalStore()
        {
            using (var window = Surfaces())
            {
                var session = new ChatSessionState
                {
                    Scope = "temporary:test"
                };
                Set(window, "currentSession", session);
                Call(window, "RenameFromQuestion", "First line\nSecond line");
                Assert.AreEqual("First line Second line", session.Title);
                Assert.AreEqual(session.Title, Get<Label>(window, "sessionTitle").Text);
                var prompt = Get<object>(window, "prompt");
                prompt.GetType().GetProperty("Text").SetValue(prompt, "unsent draft", null);
                Get<List<object>>(window, "messages").Add(new Dictionary<string, object> { ["role"] = "user", ["content"] = "hello" });
                Call(window, "SaveCurrentSession");
                Assert.AreEqual("unsent draft", session.Draft);
                StringAssert.Contains(session.MessagesJson, "hello");
                Assert.AreEqual(0, session.DraftAttachments.Length);
                Set(window, "currentSession", null);
            }
        }

        /// <summary>Rogne le titre saisi manuellement et le limite à 120 caractères.</summary>
        [TestMethod]
        [STATestMethod]
        public void ManualRenameTrimsAndCapsTitle()
        {
            using (var window = Surfaces())
            {
                var session = new ChatSessionState
                {
                    Scope = "temporary:test"
                };
                Set(window, "currentSession", session);
                Get<TextBox>(window, "chatTitleEditor").Text = "  " + new string('a', 130) + "  ";
                Call(window, "RenameCurrentChat");
                Assert.AreEqual(120, session.Title.Length);
                Assert.AreEqual(session.Title, Get<Label>(window, "sessionTitle").Text);
                Set(window, "currentSession", null);
            }
        }

        /// <summary>Répare l’historique interrompu sans supprimer les messages antérieurs à l’appel d’outil incomplet.</summary>
        [TestMethod]
        [STATestMethod]
        public void InterruptedToolHistoryDropsOnlyUnfinishedAssistantTail()
        {
            using (var window = Surfaces())
            {
                var messages = Get<List<object>>(window, "messages");
                messages.Add(new Dictionary<string, object> { ["role"] = "system", ["content"] = "rules" });
                messages.Add(new Dictionary<string, object> { ["role"] = "user", ["content"] = "question" });
                messages.Add(new Dictionary<string, object> { ["role"] = "assistant", ["tool_calls"] = new object[] { new Dictionary<string, object> { ["id"] = "pending" } } });
                Call(window, "RepairInterruptedToolHistory");
                Assert.AreEqual(3, messages.Count);
                var repaired = new JavaScriptSerializer().Serialize(messages[2]);
                StringAssert.Contains(repaired, "assistant");
            }
        }

        /// <summary>Conserve les messages d’un appel d’outil dont la réponse est présente.</summary>
        [TestMethod]
        [STATestMethod]
        public void CompletedToolHistoryIsPreserved()
        {
            using (var window = Surfaces())
            {
                var messages = Get<List<object>>(window, "messages");
                messages.Add(new Dictionary<string, object> { ["role"] = "user", ["content"] = "question" });
                messages.Add(new Dictionary<string, object> { ["role"] = "assistant", ["tool_calls"] = new object[] { new Dictionary<string, object> { ["id"] = "done" } } });
                messages.Add(new Dictionary<string, object> { ["role"] = "tool", ["tool_call_id"] = "done" });
                Call(window, "RepairInterruptedToolHistory");
                Assert.AreEqual(3, messages.Count);
            }
        }
    }
}
namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Threading.Tasks;
    using System.Web.Script.Serialization;
    using System.Windows.Threading;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    /// <summary>Vérifie la découverte des sessions, leur activation et la validité de leur historique.</summary>
    public sealed partial class ChatWindowStateTests
    {
                /// <summary>Gère les projets enregistrés ou non, les réponses mal formées et les erreurs de stockage ou d’hôte.</summary>
[STATestMethod, TestCategory("Unit")]
        public void SessionDiscoveryHandlesSavedUnsavedEmptyMalformedAndFailingProjectResponses()
        {
            using (var runtime = new RuntimeScope())
            {
                foreach (var projects in new object[] { new object[0], "unexpected", new object[] { "skip", new { Name = "Unsaved", FileName = "" }, new { Name = "Relative", FileName = "relative.xlsm" }, new { Name = "Saved", FileName = @"C:\Temp\Saved.xlsm" } } })
                {
                    runtime.Host = r => Response.Success(r.Command == "list_projects" ? projects : (object)new { SelectedProject = "Unsaved" });
                    using (var window = LoadedWindow(runtime.Session))
                    {
                        var scopes = Get<System.Windows.Forms.ComboBox>(window, "scopePicker");
                        if (scopes.Items.Count == 0) { var timer = Get<DispatcherTimer>(window, "projectRetryTimer"); TimerTick(timer); runtime.Host = r => Response.Success(r.Command == "list_projects" ? (object)new[] { new { Name = "P", FileName = "" } } : new { }); TimerTick(timer); CompleteScopeLoad(window); Assert.AreEqual(1, scopes.Items.Count); }
                        else Assert.AreEqual(3, scopes.Items.Count);
                    }
                }
                ChatWindow.OpenHistory = p => { throw new IOException("store unavailable"); }; runtime.Host = r => Response.Failure("host unavailable");
                using (var window = LoadedWindow(runtime.Session)) { Assert.IsTrue(Get<bool>(window, "storageFailed")); var timer = Get<DispatcherTimer>(window, "projectRetryTimer"); TimerTick(timer); runtime.Host = r => { throw new IOException("retry failure"); }; TimerTick(timer); window.Dispose(); TimerTick(timer); }
            }
        }
                /// <summary>Vérifie la découverte, les événements de sélection et l’état valide lorsque le stockage est indisponible.</summary>
[STATestMethod, TestCategory("Unit")]
        public void SessionsDiscoveryEventsAndUnavailableStorageKeepValidState()
        {
            using (var runtime = new RuntimeScope())
            {
                foreach (var state in new[] { Response.Failure("state unavailable"), Response.Success((object)null), Response.Success(new { }), Response.Success(new { SelectedProject = "P" }), Response.Success(new { SelectedProject = "P", SelectedProjectPath = @"C:\Temp\P.xlsm" }) })
                {
                    runtime.Host = r => r.Command == "list_projects" ? Response.Success(new[] { new { Name = "P", FileName = @"C:\Temp\P.xlsm" } }) : state;
                    using (var window = LoadedWindow(runtime.Session))
                    {
                        var sessions = Get<System.Windows.Forms.ListBox>(window, "sessionList"); var first = Get<ChatSessionState>(window, "currentSession"); Call(window, "NewSession", (object)null); sessions.SelectedItem = first; Assert.AreSame(first, Get<ChatSessionState>(window, "currentSession")); sessions.SelectedIndex = -1; sessions.SelectedItem = first;
                        Set(window, "loadingSession", true); sessions.SelectedIndex = -1; sessions.SelectedItem = first; Set(window, "loadingSession", false); Set(window, "busy", true); sessions.SelectedIndex = -1; sessions.SelectedItem = first; Set(window, "busy", false);
                        Get<System.Windows.Forms.TextBox>(window, "historySearch").Text = "missing"; TimerTick(Get<DispatcherTimer>(window, "historySearchTimer")); Assert.AreEqual(0, sessions.Items.Count); Get<System.Windows.Forms.TextBox>(window, "historySearch").Text = "";
                    }
                }
                ChatWindow.OpenHistory = p => { throw new IOException("store unavailable"); }; runtime.Host = r => Response.Success(r.Command == "list_projects" ? (object)new[] { new { Name = "P", FileName = "" } } : new { SelectedProject = "P" });
                using (var window = LoadedWindow(runtime.Session)) { Call(window, "SaveProjectMemory"); StringAssert.Contains(Get<System.Windows.Forms.Label>(window, "status").Text, UiText.Get("Memory requires")); Call(window, "EnsureCurrentScope"); Get<System.Windows.Forms.ComboBox>(window, "scopePicker").SelectedIndex = -1; Call(window, "ChangeScope"); CompleteScopeLoad(window); var failure = Assert.ThrowsException<System.Reflection.TargetInvocationException>(() => Call(window, "EnsureCurrentScope")); Assert.IsInstanceOfType(failure.InnerException, typeof(InvalidOperationException)); }
            }
        }
                /// <summary>Active, restaure, renomme et archive des sessions puis persiste leur état localement.</summary>
[STATestMethod, TestCategory("Unit")]
        public void SessionsActivateRestoreRenameArchiveCacheAndPersistLocally()
        {
            using (var runtime = new RuntimeScope())
            using (var window = LoadedWindow(runtime.Session))
            {
                var original = Get<ChatSessionState>(window, "currentSession"); var store = Get<ChatSessionStore>(window, "sessionStore"); var scope = original.Scope;
                Get<System.Windows.Forms.TextBox>(window, "memoryEditor").Text = "memory"; Call(window, "SaveProjectMemory"); Assert.AreEqual("memory", store.ReadMemory(scope));
                var session = new ChatSessionState { Scope = scope, Provider = "missing-provider", Mode = ChatMode.Plan, Draft = null, DraftAttachments = null, DraftReferences = null, MessagesJson = "[{},7,{\"role\":\"assistant\",\"tool_calls\":[{\"id\":\"orphan\"}]}]", Entries = new List<ChatEntry> { new ChatEntry { Speaker = "Code", Change = new CodeChange { Project = "P", Module = "M", Before = "old", After = "new" } }, new ChatEntry { Speaker = "Assistant", Text = "stream complete", StreamId = "saved" } } };
                Call(window, "ActivateSession", session, true); Assert.AreSame(session, Get<ChatSessionState>(window, "currentSession")); Assert.AreEqual(2, Get<List<ChatEntry>>(window, "transcriptEntries").Count); Assert.IsTrue(Get<HashSet<string>>(window, "completedStreams").Contains("saved"));
                Call(window, "ScheduleSessionSave"); TimerTick(Get<DispatcherTimer>(window, "saveTimer"));
                Assert.IsTrue(Get<ChatPersistenceWorker>(window, "persistenceWorker").Flush(5000));
                Assert.AreEqual(2, store.List(scope).Count);
                Set(window, "busy", true); Call(window, "ActivateSession", original, true); Call(window, "NewSession", (object)null); Call(window, "ChangeScope"); CompleteScopeLoad(window); Call(window, "RenameCurrentChat"); Call(window, "ToggleArchiveCurrentChat"); Call(window, "SaveProjectMemory"); Set(window, "busy", false);
                Set(window, "loadingSession", true); Call(window, "ScheduleSessionSave"); Call(window, "SaveCurrentSession"); Call(window, "ChangeScope"); CompleteScopeLoad(window); Set(window, "loadingSession", false);
                Call(window, "RenameFromQuestion", new string('x', 100)); Assert.AreEqual(session.Title, Get<System.Windows.Forms.Label>(window, "sessionTitle").Text);
                session.Title = "Nouvelle conversation"; Get<List<ChatEntry>>(window, "transcriptEntries").Clear(); Call(window, "RenameFromQuestion", new string('x', 100)); Assert.AreEqual(62, session.Title.Length);
                Get<System.Windows.Forms.TextBox>(window, "chatTitleEditor").Text = " "; Call(window, "RenameCurrentChat"); Get<System.Windows.Forms.TextBox>(window, "chatTitleEditor").Text = "short"; Call(window, "RenameCurrentChat"); Assert.AreEqual("short", session.Title);
                session.Archived = true; Call(window, "ToggleArchiveCurrentChat"); Assert.IsFalse(session.Archived); Call(window, "ToggleArchiveCurrentChat"); Assert.IsTrue(session.Archived); Assert.AreNotSame(session, Get<ChatSessionState>(window, "currentSession"));
                Call(window, "NewSession", "Ollama"); Assert.AreEqual("Ollama", Get<ChatSessionState>(window, "currentSession").Provider);
                var scopes = Get<System.Windows.Forms.ComboBox>(window, "scopePicker"); var extra = AddScope(window, "temporary:extra"); scopes.SelectedItem = extra; Call(window, "ChangeScope"); CompleteScopeLoad(window); Assert.AreEqual("temporary:extra", Get<ChatSessionState>(window, "currentSession").Scope); scopes.SelectedIndex = 0; Call(window, "ChangeScope"); CompleteScopeLoad(window); Assert.AreEqual(scope, Get<ChatSessionState>(window, "currentSession").Scope);
                runtime.Host = r => Response.Failure("scope closed"); Assert.ThrowsException<System.Reflection.TargetInvocationException>(() => Call(window, "EnsureCurrentScope")); runtime.Host = r => Response.Success(new object[0]); Assert.ThrowsException<System.Reflection.TargetInvocationException>(() => Call(window, "EnsureCurrentScope"));
                var worker = Get<ChatPersistenceWorker>(window, "persistenceWorker");
                Assert.IsTrue(worker.Flush(5000)); worker.Dispose();
                Call(window, "SaveCurrentSession"); Assert.IsTrue(Get<bool>(window, "storageFailed"));
                store.Dispose(); Call(window, "SaveProjectMemory"); Call(window, "ChangeScope"); CompleteScopeLoad(window);
            }
        }
                /// <summary>Préserve les tours terminés et répare uniquement les tours utilisateur incomplets.</summary>
[STATestMethod, TestCategory("Unit")]
        public void InterruptedHistoryValidationPreservesCompletedAndRepairsOnlyPendingUserTurns()
        {
            using (var window = Surfaces())
            {
                var messages = Get<List<object>>(window, "messages");
                foreach (var rawCalls in new object[] { "invalid", new object[] { null, new { }, new Dictionary<string, object> { { "id", "pending" } } } })
                {
                    messages.Clear(); messages.Add(new object()); messages.Add(new Dictionary<string, object>()); messages.Add(new Dictionary<string, object> { { "role", "assistant" }, { "tool_calls", rawCalls } }); Call(window, "RepairInterruptedToolHistory"); Assert.AreEqual(3, messages.Count);
                    messages.Insert(0, new Dictionary<string, object> { { "role", "user" } }); messages.Add(new Dictionary<string, object> { { "role", "tool" } }); Call(window, "RepairInterruptedToolHistory"); Assert.IsTrue(messages.Count >= 2);
                }
                Set(window, "currentSession", null); Call(window, "RenameFromQuestion", "ignored"); Call(window, "RenameCurrentChat"); Call(window, "ToggleArchiveCurrentChat"); Call(window, "ScheduleSessionSave"); Call(window, "NewSession", (object)null); Call(window, "SaveProjectMemory");
            }
        }
                /// <summary>Refuse une identité de scope obsolète au moment de sélectionner une session.</summary>
[STATestMethod, TestCategory("Unit")]
        public void SessionsRejectStaleScopeIdentityAtTheSelectionBoundary()
        {
            using (var runtime = new RuntimeScope())
            using (var window = LoadedWindow(runtime.Session))
            {
                var picker = Get<System.Windows.Forms.ComboBox>(window, "scopePicker"); var scope = picker.SelectedItem; var type = scope.GetType(); var project = type.GetField("Project"); var key = type.GetField("Key"); var originalProject = project.GetValue(scope); var originalKey = key.GetValue(scope);
                try
                {
                    // Keep the persisted conversation identity while simulating a stale selector at the UI boundary.
                    key.SetValue(scope, @"C:\Temp\Previous.xlsm"); Assert.ThrowsException<System.Reflection.TargetInvocationException>(() => Call(window, "EnsureCurrentScope"));
                    project.SetValue(scope, ""); runtime.Host = r => Response.Success(new[] { new { Name = "P", FileName = "" } }); Assert.ThrowsException<System.Reflection.TargetInvocationException>(() => Call(window, "EnsureCurrentScope"));
                    runtime.Host = r => Response.Success((object)null); Assert.ThrowsException<System.Reflection.TargetInvocationException>(() => Call(window, "EnsureCurrentScope"));
                }
                finally { project.SetValue(scope, originalProject); key.SetValue(scope, originalKey); }
                using (var design = new ChatWindow())
                {
                    var list = Get<System.Windows.Forms.ListBox>(design, "sessionList");
                    try { Set(design, "sessionList", null); Call(design, "RefreshHistory"); Assert.IsFalse(Get<bool>(design, "loadingSession")); }
                    finally { Set(design, "sessionList", list); }
                    Get<System.Windows.Forms.TextBox>(design, "historySearch").Text = null; Assert.AreEqual("", Get<System.Windows.Forms.TextBox>(design, "historySearch").Text);
                }
            }
        }
    }
}

namespace VBAi.Tests.Unit
{
    public sealed partial class ChatWindowStateTests
    {
        [Microsoft.VisualStudio.TestTools.UnitTesting.STATestMethod]
        public void ScopeValidationWithoutCurrentSessionRevokesReadGrantsAndSharedAccess()
        {
            using (var runtime = new RuntimeScope())
            using (var window = LoadedWindow(runtime.Session))
            {
                var tools = Get<VBAi.LlmVbeTools>(window, "tools"); tools.SetReadAccess(new[] { "Foreign" }, true);
                Set(window, "currentSession", null); Call(window, "EnsureCurrentScope");
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(@"C:\Temp\P.xlsm", tools.BoundProject);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsException<System.InvalidOperationException>(() => tools.RequireProjectRead("Foreign"));
                var refused = new System.Web.Script.Serialization.JavaScriptSerializer().Deserialize<VBAi.Response>(tools.Invoke("code_panes", "{}"));
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsFalse(refused.Ok); Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsFalse(string.IsNullOrEmpty(refused.Error));
            }
        }
    }
}
