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
        public void CachedMetadataScopeRefreshesBusyUiSelectionWithoutReadingTheHost()
        {
            using (var runtime = new RuntimeScope())
            using (var window = LoadedWindow(runtime.Session))
            {
                var tools = Get<LlmVbeTools>(window, "tools");
                var originalSession = Get<ChatSessionState>(window, "currentSession");
                string originalBound = tools.BoundProject;
                object other = AddScope(window, @"C:\Owned\Other.xlsm");
                Set(window, "busy", true);
                Get<ComboBox>(window, "scopePicker").SelectedItem = other;
                Assert.AreEqual(originalBound, tools.BoundProject, "Busy scope switching leaves the old binding cached until guarded.");
                int reads = 0; runtime.Host = request => { reads++; throw new InvalidOperationException("No host read is permitted here."); };
                tools.ValidateCachedScope();
                Assert.AreEqual(@"C:\Owned\Other.xlsm", tools.BoundProject); Assert.AreEqual(0, reads);
                Assert.AreSame(originalSession, Get<ChatSessionState>(window, "currentSession"));
                Set(window, "busy", false);
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void CachedMetadataScopeRefusesUnavailableLoadingAndMissingSelectionWithoutHostReads()
        {
            using (var runtime = new RuntimeScope())
            using (var window = LoadedWindow(runtime.Session))
            {
                var tools = Get<LlmVbeTools>(window, "tools"); int reads = 0;
                runtime.Host = request => { reads++; throw new InvalidOperationException("No host read is permitted here."); };
                foreach (string flag in new[] { "sessionViewUnavailable", "loadingScope" })
                {
                    Set(window, flag, true);
                    Assert.ThrowsException<InvalidOperationException>(() => tools.ValidateCachedScope());
                    Set(window, flag, false);
                }
                Set(window, "busy", true); Get<ComboBox>(window, "scopePicker").SelectedIndex = -1;
                Assert.ThrowsException<InvalidOperationException>(() => tools.ValidateCachedScope());
                Assert.AreEqual(0, reads); Set(window, "busy", false);
            }
        }

        private static void UseLiveProjectCatalogue(RuntimeScope runtime)
        {
            runtime.Host = request => request.Command == "list_projects"
                ? Response.Success(runtime.Vbe.VBProjects.ConvertAll(project => new {
                    project.Name, HostPath = project.ThrowFileName ? null : project.FileName
                }).ToArray()) : Response.Success(new { SelectedProject = "P" });
        }

        public sealed class BrokenTransientCapture
        {
            public string Content { get { throw new IOException("Owned synthetic capture failure"); } }
        }

        [STATestMethod, TestCategory("Unit")]
        public void TemporaryHistoryAndNotesSurviveSwitchingAndCanBeDeletedWithoutStorage()
        {
            using (var runtime = new RuntimeScope())
            {
                runtime.Vbe.VBProjects[0].FileName = "";
                runtime.Vbe.VBProjects.Add(new VbeSessionTests.FakeProject { Name = "Q", FileName = "", Mode = 2 });
                UseLiveProjectCatalogue(runtime);
                ChatWindow.OpenHistory = path => { throw new IOException("Owned store unavailable"); };
                using (var window = LoadedWindow(runtime.Session))
                {
                    window.ReadScope = (path, scope, include) => { Assert.Fail("Temporary scopes cannot read SQLite."); return null; };
                    var first = Get<ChatSessionState>(window, "currentSession");
                    var picker = Get<ComboBox>(window, "scopePicker");
                    Get<System.Windows.Controls.TextBox>(window, "prompt").Text = "private draft";
                    Call(window, "AddEntry", new ChatEntry { Speaker = "You", Text = "private transcript" });
                    Call(window, "SaveCurrentSession");
                    Get<TextBox>(window, "memoryEditor").Text = "private notes"; Call(window, "SaveProjectMemory");
                    Call(window, "NewSession", (object)null);
                    picker.SelectedIndex = 1; CompleteScopeLoad(window);
                    var other = Get<ChatSessionState>(window, "currentSession");
                    Assert.AreEqual("", Get<TextBox>(window, "memoryEditor").Text);
                    picker.SelectedIndex = 0; CompleteScopeLoad(window);
                    Assert.AreEqual("private notes", Get<TextBox>(window, "memoryEditor").Text);
                    Get<ListBox>(window, "sessionList").SelectedItem = first;
                    Assert.AreEqual("private draft", Get<System.Windows.Controls.TextBox>(window, "prompt").Text);
                    Assert.AreEqual("private transcript", first.Entries[0].Text);
                    Assert.IsTrue(Get<Button>(window, "deleteSession").Enabled);
                    ChatWindow.ShowNotice = (owner, text, caption, buttons, icon) => DialogResult.Yes;
                    CompleteOnSta((Task)Call(window, "DeleteSelectedSessionAsync"));
                    Assert.AreNotSame(first, Get<ChatSessionState>(window, "currentSession"));
                    picker.SelectedIndex = 1; CompleteScopeLoad(window);
                    Assert.AreSame(other, Get<ChatSessionState>(window, "currentSession"));
                    runtime.Vbe.VBProjects.Clear(); Call(window, "RefreshAvailableScopes", runtime.Session);
                    window.Dispose();
                    Assert.IsFalse(Directory.Exists(runtime.Root), "Abandoned temporary projects must leave no history or recovery files.");
                }
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void TemporaryCaptureFailureKeepsTheLastSnapshotAndNeverWritesEmergencyRecovery()
        {
            using (var runtime = new RuntimeScope())
            {
                runtime.Vbe.VBProjects[0].FileName = ""; UseLiveProjectCatalogue(runtime);
                using (var window = LoadedWindow(runtime.Session))
                {
                    var original = Get<ChatSessionState>(window, "currentSession");
                    Get<System.Windows.Controls.TextBox>(window, "prompt").Text = "complete draft"; Call(window, "SaveCurrentSession");
                    Get<List<object>>(window, "messages").Add(new BrokenTransientCapture());
                    Get<System.Windows.Controls.TextBox>(window, "prompt").Text = "later private draft";
                    Call(window, "SaveCurrentSession");
                    Assert.AreEqual("complete draft", original.Draft);
                    Assert.IsTrue(Get<ChatPersistenceWorker>(window, "persistenceWorker").Flush(5000));
                    window.Dispose();
                    Assert.IsFalse(Directory.Exists(ChatWindow.HistoryPath() + ".recovery"));
                    using (var store = new ChatSessionStore(ChatWindow.HistoryPath())) Assert.IsFalse(store.HasSessions(original.Scope));
                }
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void ExactLiveFirstSavePromotesAllSessionIdsDraftsAndNotesAndSaveAsStaysIsolated()
        {
            using (var runtime = new RuntimeScope())
            {
                var project = runtime.Vbe.VBProjects[0]; project.FileName = ""; UseLiveProjectCatalogue(runtime);
                using (var window = LoadedWindow(runtime.Session))
                {
                    var first = Get<ChatSessionState>(window, "currentSession"); string transientKey = first.Scope;
                    var picker = Get<ComboBox>(window, "scopePicker"); object exactScope = picker.SelectedItem;
                    Get<System.Windows.Controls.TextBox>(window, "prompt").Text = "first draft"; Call(window, "SaveCurrentSession");
                    Get<TextBox>(window, "memoryEditor").Text = "retained note"; Call(window, "SaveProjectMemory");
                    Call(window, "NewSession", (object)null);
                    var current = Get<ChatSessionState>(window, "currentSession"); current.ReadProjectGrants = new[] { "Foreign" };
                    string originalThread = Get<CodexAppServerClient>(window, "codex").ThreadId;
                    current.CodexThreadId = originalThread; current.CodexThreadHome = ProviderSessionStorage.CodexHome; current.ResumeContext = "owned fixture context";
                    Get<System.Windows.Controls.TextBox>(window, "prompt").Text = "latest draft";
                    project.FileName = @"C:\Owned\FirstSave.xlsm";
                    Call(window, "RefreshAvailableScopes", runtime.Session);
                    Assert.AreSame(exactScope, picker.SelectedItem); Assert.AreSame(current, Get<ChatSessionState>(window, "currentSession"));
                    Assert.AreEqual(project.FileName.ToUpperInvariant(), first.Scope); Assert.AreEqual(first.Scope, current.Scope);
                    Assert.AreEqual(originalThread, current.CodexThreadId); Assert.AreEqual(ProviderSessionStorage.CodexHome, current.CodexThreadHome);
                    Assert.AreEqual("owned fixture context", current.ResumeContext);
                    Assert.IsTrue(Get<ChatPersistenceWorker>(window, "persistenceWorker").Flush(5000));
                    var store = Get<ChatSessionStore>(window, "sessionStore"); var persisted = store.List(first.Scope);
                    Assert.AreEqual(2, persisted.Count); Assert.IsTrue(persisted.Exists(item => item.Id == first.Id && item.Draft == "first draft"));
                    Assert.IsTrue(persisted.Exists(item => item.Id == current.Id && item.Draft == "latest draft"));
                    Assert.AreEqual("retained note", store.ReadMemory(first.Scope)); Assert.AreEqual(0, store.List(transientKey).Count);
                    Call(window, "EnsureCurrentScope");
                    project.FileName = @"C:\Owned\SaveAs.xlsm";
                    Call(window, "RefreshAvailableScopes", runtime.Session); Assert.AreEqual(-1, picker.SelectedIndex);
                    Assert.AreEqual(@"C:\OWNED\FIRSTSAVE.XLSM", current.Scope);
                    runtime.Transport.Requests.Clear(); runtime.Transport.NextThreadId = "fresh SaveAs thread";
                    picker.SelectedIndex = 0; CompleteScopeLoad(window);
                    var fresh = Get<ChatSessionState>(window, "currentSession");
                    Assert.AreNotEqual(current.Id, fresh.Id); Assert.AreEqual(0, fresh.ReadProjectGrants.Length);
                    Assert.AreEqual("fresh SaveAs thread", fresh.CodexThreadId); Assert.AreNotEqual(originalThread, fresh.CodexThreadId); Assert.IsNull(fresh.ResumeContext);
                    Assert.IsTrue(runtime.Transport.Requests.Exists(line => line.Contains("thread/start")));
                    Assert.IsFalse(runtime.Transport.Requests.Exists(line => line.Contains("thread/resume")));
                    Assert.AreEqual("", Get<TextBox>(window, "memoryEditor").Text);
                }
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void SameNameReplacementOrUnprovableIdentityCannotReuseOrPromotePrivateTemporaryHistory()
        {
            foreach (bool unavailable in new[] { false, true })
                using (var runtime = new RuntimeScope())
                {
                    runtime.Vbe.VBProjects[0].FileName = ""; UseLiveProjectCatalogue(runtime);
                    using (var window = LoadedWindow(runtime.Session))
                    {
                        var first = Get<ChatSessionState>(window, "currentSession");
                        Get<System.Windows.Controls.TextBox>(window, "prompt").Text = "private old object"; Call(window, "SaveCurrentSession");
                        if (unavailable) ChatWindow.ReadScopeProject = (session, selector) => { throw new IOException("Identity unavailable"); };
                        else { runtime.Vbe.VBProjects.Clear(); runtime.Vbe.VBProjects.Add(new VbeSessionTests.FakeProject { Name = "P", FileName = "", Mode = 2 }); }
                        Call(window, "RefreshAvailableScopes", runtime.Session);
                        var picker = Get<ComboBox>(window, "scopePicker"); Assert.AreEqual(-1, picker.SelectedIndex);
                        picker.SelectedIndex = 0; CompleteScopeLoad(window);
                        var fresh = Get<ChatSessionState>(window, "currentSession");
                        Assert.AreNotEqual(first.Scope, fresh.Scope); Assert.AreNotEqual(first.Id, fresh.Id);
                        Assert.AreEqual("", Get<System.Windows.Controls.TextBox>(window, "prompt").Text);
                        runtime.Vbe.VBProjects[0].FileName = @"C:\Owned\Replacement.xlsm";
                        Call(window, "RefreshAvailableScopes", runtime.Session);
                        Assert.IsTrue(ChatSessionStore.IsTransientScope(first.Scope));
                        Assert.IsFalse(Get<ChatSessionStore>(window, "sessionStore").List(@"C:\OWNED\REPLACEMENT.XLSM").Exists(item => item.Id == first.Id));
                    }
                }
        }

        [STATestMethod, TestCategory("Unit")]
        public void PromotionWithUnavailableStorageDefersPersistenceAndRetainsNotesWithoutDiskRecovery()
        {
            using (var runtime = new RuntimeScope())
            {
                var project = runtime.Vbe.VBProjects[0]; project.FileName = ""; UseLiveProjectCatalogue(runtime);
                ChatWindow.OpenHistory = path => { throw new IOException("History unavailable"); };
                using (var window = LoadedWindow(runtime.Session))
                {
                    var session = Get<ChatSessionState>(window, "currentSession");
                    string oldKey = session.Scope;
                    Get<TextBox>(window, "memoryEditor").Text = "keep despite no store"; Call(window, "SaveProjectMemory");
                    project.FileName = @"C:\Owned\OfflineSave.xlsm"; Call(window, "RefreshAvailableScopes", runtime.Session);
                    Assert.AreEqual(oldKey, session.Scope);
                    Call(window, "EnsureCurrentScope");
                    var notes = Get<Dictionary<string, string>>(window, "transientMemory");
                    Assert.AreEqual("keep despite no store", notes[oldKey]);
                    var other = AddScope(window, @"C:\Owned\Other.xlsm");
                    Get<ComboBox>(window, "scopePicker").SelectedItem = other; CompleteScopeLoad(window);
                    Get<ComboBox>(window, "scopePicker").SelectedIndex = 0; CompleteScopeLoad(window);
                    Assert.AreEqual("keep despite no store", Get<TextBox>(window, "memoryEditor").Text);
                    StringAssert.Contains(Get<Label>(window, "status").Text, UiText.Get("Local history is kept in memory because storage is unavailable."));
                    project.FileName = @"C:\Owned\OfflineSaveAs.xlsm"; Call(window, "RefreshAvailableScopes", runtime.Session);
                    Assert.AreEqual(-1, Get<ComboBox>(window, "scopePicker").SelectedIndex);
                    Assert.AreEqual(oldKey, session.Scope);
                    Assert.IsFalse(Directory.Exists(runtime.Root));
                }
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void SavedScopeCollisionRefusesTemporaryPromotionAndKeepsItsExistingHistory()
        {
            using (var runtime = new RuntimeScope())
            {
                var project = runtime.Vbe.VBProjects[0]; project.FileName = ""; UseLiveProjectCatalogue(runtime);
                using (var window = LoadedWindow(runtime.Session))
                {
                    var temporary = Get<ChatSessionState>(window, "currentSession");
                    var store = Get<ChatSessionStore>(window, "sessionStore");
                    var old = new ChatSessionState { Scope = @"C:\OWNED\COLLISION.XLSM", Draft = "existing saved history" }; store.Save(old);
                    project.FileName = @"C:\Owned\Collision.xlsm"; Call(window, "RefreshAvailableScopes", runtime.Session);
                    Assert.IsTrue(ChatSessionStore.IsTransientScope(temporary.Scope)); Assert.AreEqual(-1, Get<ComboBox>(window, "scopePicker").SelectedIndex);
                    Get<ComboBox>(window, "scopePicker").SelectedIndex = 0; CompleteScopeLoad(window);
                    Assert.AreEqual(old.Id, Get<ChatSessionState>(window, "currentSession").Id); Assert.AreEqual(1, store.List(old.Scope).Count);
                }
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void NotesOnlySavedScopeCollisionCannotOverwriteNotesOrInheritTemporaryProviderAuthority()
        {
            using (var runtime = new RuntimeScope())
            {
                var project = runtime.Vbe.VBProjects[0]; project.FileName = ""; UseLiveProjectCatalogue(runtime);
                using (var window = LoadedWindow(runtime.Session))
                {
                    var temporary = Get<ChatSessionState>(window, "currentSession");
                    temporary.CodexThreadId = "private original thread"; temporary.ReadProjectGrants = new[] { "Foreign" };
                    Get<TextBox>(window, "memoryEditor").Text = "temporary notes"; Call(window, "SaveProjectMemory");
                    var store = Get<ChatSessionStore>(window, "sessionStore"); string destination = @"C:\OWNED\NOTESONLY.XLSM";
                    store.SaveMemory(destination, "existing separate notes"); Assert.IsFalse(store.HasSessions(destination)); Assert.IsTrue(store.HasScopeData(destination));
                    project.FileName = @"C:\Owned\NotesOnly.xlsm"; Call(window, "RefreshAvailableScopes", runtime.Session);
                    Assert.IsTrue(ChatSessionStore.IsTransientScope(temporary.Scope));
                    runtime.Transport.Requests.Clear(); runtime.Transport.NextThreadId = "fresh notes-only thread";
                    Get<ComboBox>(window, "scopePicker").SelectedIndex = 0; CompleteScopeLoad(window);
                    var fresh = Get<ChatSessionState>(window, "currentSession");
                    Assert.AreNotEqual(temporary.Id, fresh.Id); Assert.AreEqual("fresh notes-only thread", fresh.CodexThreadId); Assert.AreEqual(0, fresh.ReadProjectGrants.Length);
                    Assert.IsNull(fresh.ResumeContext);
                    Assert.IsTrue(runtime.Transport.Requests.Exists(line => line.Contains("thread/start")));
                    Assert.IsFalse(runtime.Transport.Requests.Exists(line => line.Contains("thread/resume")));
                    Assert.AreEqual("existing separate notes", Get<TextBox>(window, "memoryEditor").Text);
                    Assert.AreEqual("existing separate notes", store.ReadMemory(destination));
                }
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void FailedSqlitePromotionPreservesNotesAcrossSwitchesUntilTheirDurableSave()
        {
            using (var runtime = new RuntimeScope())
            {
                var project = runtime.Vbe.VBProjects[0]; project.FileName = ""; UseLiveProjectCatalogue(runtime);
                using (var window = LoadedWindow(runtime.Session))
                {
                    var session = Get<ChatSessionState>(window, "currentSession");
                    Get<TextBox>(window, "memoryEditor").Text = "retain failed notes"; Call(window, "SaveProjectMemory");
                    var worker = Get<ChatPersistenceWorker>(window, "persistenceWorker"); Assert.IsTrue(worker.Flush(5000)); worker.Dispose(); Set(window, "persistenceWorker", null);
                    var store = Get<ChatSessionStore>(window, "sessionStore"); var native = store.StepNative;
                    int calls = 0;
                    store.StepNative = statement => ++calls == 4 ? throw new IOException("Owned SQLite note write failure") : native(statement);
                    try { project.FileName = @"C:\Owned\FailedSave.xlsm"; Call(window, "RefreshAvailableScopes", runtime.Session); }
                    finally { store.StepNative = native; }
                    Assert.IsTrue(ChatSessionStore.IsTransientScope(session.Scope));
                    Assert.AreEqual("retain failed notes", Get<Dictionary<string, string>>(window, "transientMemory")[session.Scope]);
                    Assert.AreEqual(0, store.List(project.FileName.ToUpperInvariant()).Count);
                    Call(window, "EnsureCurrentScope");
                    var other = AddScope(window, @"C:\Owned\Other.xlsm"); Get<ComboBox>(window, "scopePicker").SelectedItem = other; CompleteScopeLoad(window);
                    Get<ComboBox>(window, "scopePicker").SelectedIndex = 0; CompleteScopeLoad(window);
                    Assert.AreEqual("retain failed notes", Get<TextBox>(window, "memoryEditor").Text);
                    Call(window, "RefreshAvailableScopes", runtime.Session);
                    Assert.AreEqual(project.FileName.ToUpperInvariant(), session.Scope);
                    Assert.AreEqual("retain failed notes", store.ReadMemory(session.Scope));
                    Assert.IsFalse(Get<Dictionary<string, string>>(window, "transientMemory").ContainsKey(session.Scope));
                }
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void SaturatedSaveQueueCannotPreventAtomicInitialPromotionAndLaterWorkerSavesUseItsRevisions()
        {
            using (var runtime = new RuntimeScope())
            using (var inFlight = new ManualResetEventSlim())
            using (var release = new ManualResetEventSlim())
            {
                var project = runtime.Vbe.VBProjects[0]; project.FileName = ""; UseLiveProjectCatalogue(runtime);
                using (var window = LoadedWindow(runtime.Session))
                {
                    var first = Get<ChatSessionState>(window, "currentSession"); string temporaryKey = first.Scope;
                    Call(window, "NewSession", (object)null); var second = Get<ChatSessionState>(window, "currentSession");
                    var previous = Get<ChatPersistenceWorker>(window, "persistenceWorker"); Assert.IsTrue(previous.Flush(5000)); previous.Dispose();
                    int callbacks = 0;
                    using (var worker = new ChatPersistenceWorker(ChatWindow.HistoryPath(), (snapshot, version, error) => {
                        if (Interlocked.Increment(ref callbacks) == 1) { inFlight.Set(); release.Wait(10000); }
                    }))
                    {
                        Set(window, "persistenceWorker", worker);
                        var serializer = new JavaScriptSerializer();
                        try
                        {
                            var blocking = new ChatSessionState { Scope = "queue-fixture" }; worker.Enqueue(new ChatPersistenceWorker.Snapshot(blocking, serializer.Serialize(blocking)));
                            Assert.IsTrue(inFlight.Wait(5000));
                            for (int i = 0; i < 64; i++)
                            {
                                var filler = new ChatSessionState { Scope = "queue-fixture" }; worker.Enqueue(new ChatPersistenceWorker.Snapshot(filler, serializer.Serialize(filler)));
                            }
                            Get<System.Windows.Controls.TextBox>(window, "prompt").Text = "retained after saturation";
                            project.FileName = @"C:\Owned\QueueSave.xlsm"; Call(window, "RefreshAvailableScopes", runtime.Session);
                            Assert.IsFalse(Get<bool>(window, "storageFailed"));
                            var cache = Get<Dictionary<string, List<ChatSessionState>>>(window, "cachedScopes"); Assert.AreEqual(2, cache[second.Scope].Count);
                            Assert.IsFalse(cache.ContainsKey(temporaryKey));
                            string[] copies = Directory.GetFiles(worker.RecoveryDirectory, "capture-*.json"); Assert.AreEqual(0, copies.Length);
                            Assert.AreEqual(2, Get<ChatSessionStore>(window, "sessionStore").List(second.Scope).Count);
                        }
                        finally { release.Set(); }
                        Assert.IsTrue(worker.Flush(5000)); Call(window, "SaveCurrentSession"); Assert.IsTrue(worker.Flush(5000));
                        Assert.AreEqual(second.Id, Get<ChatSessionStore>(window, "sessionStore").List(second.Scope)[0].Id);
                        Assert.AreEqual("retained after saturation", second.Draft);
                        Assert.AreEqual(second.Scope, first.Scope);
                        Set(window, "persistenceWorker", null);
                    }
                }
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void UnverifiedPromotionCommitRetainsMemoryAndNeverAutomaticallyReplaysTheClaim()
        {
            using (var runtime = new RuntimeScope())
            {
                var project = runtime.Vbe.VBProjects[0]; project.FileName = ""; UseLiveProjectCatalogue(runtime);
                using (var window = LoadedWindow(runtime.Session))
                {
                    var session = Get<ChatSessionState>(window, "currentSession"); string originalScope = session.Scope;
                    Get<TextBox>(window, "memoryEditor").Text = "retained after uncertain acknowledgement"; Call(window, "SaveProjectMemory");
                    var store = Get<ChatSessionStore>(window, "sessionStore"); var native = store.StepNative; int steps = 0;
                    store.StepNative = statement => { int result = native(statement); if (++steps == 5) throw new IOException("Owned lost COMMIT acknowledgement"); return result; };
                    try { project.FileName = @"C:\Owned\UnverifiedSave.xlsm"; Call(window, "RefreshAvailableScopes", runtime.Session); }
                    finally { store.StepNative = native; }
                    Assert.AreEqual(originalScope, session.Scope);
                    Assert.AreEqual("retained after uncertain acknowledgement", Get<Dictionary<string, string>>(window, "transientMemory")[originalScope]);
                    StringAssert.Contains(Get<Label>(window, "status").Text, "unverified");
                    int repeatedSteps = 0; store.StepNative = statement => { repeatedSteps++; return native(statement); };
                    try { Call(window, "RefreshAvailableScopes", runtime.Session); }
                    finally { store.StepNative = native; }
                    Assert.AreEqual(0, repeatedSteps); Assert.AreEqual(originalScope, session.Scope);
                    Call(window, "EnsureCurrentScope");
                    Assert.AreEqual(1, store.List(project.FileName.ToUpperInvariant()).Count, "Retain the real committed database result separately from its missing acknowledgement.");
                    Assert.IsFalse(Directory.Exists(ChatWindow.HistoryPath() + ".recovery"));
                }
            }
        }
        [STATestMethod, TestCategory("Unit")]
        public void DeleteSelectedConversationRemovesPersistedAndCachedHistoryAndSelectsAnotherSession()
        {
            using (var runtime = new RuntimeScope())
            using (var window = LoadedWindow(runtime.Session))
            {
                var first = Get<ChatSessionState>(window, "currentSession");
                Call(window, "NewSession", (object)null);
                var deleted = Get<ChatSessionState>(window, "currentSession");
                var cache = Get<Dictionary<string, List<ChatSessionState>>>(window, "cachedScopes");
                cache[deleted.Scope] = new List<ChatSessionState> { first, deleted };
                var store = Get<ChatSessionStore>(window, "sessionStore");
                store.SaveMemory(deleted.Scope, "keep notes");
                int confirmations = 0;
                ChatWindow.ShowNotice = (owner, text, caption, buttons, icon) => {
                    confirmations++; StringAssert.Contains(text, deleted.DisplayTitle);
                    Assert.AreEqual(MessageBoxButtons.YesNo, buttons); return DialogResult.Yes;
                };
                window.Show();
                CompleteOnSta((Task)Call(window, "DeleteSelectedSessionAsync"));
                Assert.AreEqual(1, confirmations);
                Assert.AreSame(first, Get<ChatSessionState>(window, "currentSession"));
                Assert.IsFalse(Get<List<ChatSessionState>>(window, "scopeSessions").Contains(deleted));
                Assert.IsFalse(cache[deleted.Scope].Contains(deleted));
                Assert.IsTrue(Get<Panel>(window, "historyPanel").Visible);
                Assert.AreSame(first, Get<ListBox>(window, "sessionList").SelectedItem);
                Call(window, "SaveCurrentSession");
                Assert.IsTrue(Get<ChatPersistenceWorker>(window, "persistenceWorker").Flush(5000));
                Assert.AreEqual(1, store.List(deleted.Scope).Count);
                Assert.AreEqual(first.Id, store.List(deleted.Scope)[0].Id);
                Assert.AreEqual("keep notes", store.ReadMemory(deleted.Scope));
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void DeleteLastConversationCreatesAnEmptyReplacementAndArchivedConversationCanBeDeleted()
        {
            using (var runtime = new RuntimeScope())
            using (var window = LoadedWindow(runtime.Session))
            {
                var deleted = Get<ChatSessionState>(window, "currentSession");
                deleted.Archived = true;
                Get<CheckBox>(window, "showArchived").Checked = true;
                Call(window, "RefreshHistory");
                ChatWindow.ShowNotice = (owner, text, caption, buttons, icon) => DialogResult.Yes;
                CompleteOnSta((Task)Call(window, "DeleteSelectedSessionAsync"));
                var replacement = Get<ChatSessionState>(window, "currentSession");
                Assert.IsNotNull(replacement); Assert.AreNotEqual(deleted.Id, replacement.Id);
                Assert.AreEqual(deleted.Scope, replacement.Scope); Assert.IsFalse(replacement.Archived);
                Assert.AreEqual("", replacement.Draft);
                Assert.AreEqual(0, replacement.Entries.Count);
                Assert.IsTrue(Get<ChatPersistenceWorker>(window, "persistenceWorker").Flush(5000));
                var saved = Get<ChatSessionStore>(window, "sessionStore").List(deleted.Scope);
                Assert.AreEqual(1, saved.Count); Assert.AreEqual(replacement.Id, saved[0].Id);
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void DeleteCancellationBusyScopeLoadingAndMissingSelectionLeaveHistoryIntact()
        {
            using (var runtime = new RuntimeScope())
            using (var window = LoadedWindow(runtime.Session))
            {
                var original = Get<ChatSessionState>(window, "currentSession");
                int confirmations = 0;
                ChatWindow.ShowNotice = (owner, text, caption, buttons, icon) => { confirmations++; return DialogResult.No; };
                CompleteOnSta((Task)Call(window, "DeleteSelectedSessionAsync"));
                Assert.AreEqual(1, confirmations);
                Call(window, "SetBusy", true);
                Assert.IsFalse(Get<Button>(window, "deleteSession").Enabled);
                CompleteOnSta((Task)Call(window, "DeleteSelectedSessionAsync"));
                Call(window, "SetBusy", false);
                Assert.IsTrue(Get<Button>(window, "deleteSession").Enabled);
                foreach (string flag in new[] { "loadingScope", "loadingSession" })
                {
                    Set(window, flag, true);
                    CompleteOnSta((Task)Call(window, "DeleteSelectedSessionAsync"));
                    Set(window, flag, false);
                }
                Get<ListBox>(window, "sessionList").SelectedIndex = -1;
                Assert.IsFalse(Get<Button>(window, "deleteSession").Enabled);
                CompleteOnSta((Task)Call(window, "DeleteSelectedSessionAsync"));
                Assert.AreEqual(1, confirmations);
                Assert.AreSame(original, Get<ChatSessionState>(window, "currentSession"));
                Assert.IsTrue(Get<List<ChatSessionState>>(window, "scopeSessions").Contains(original));
                Call(window, "RefreshHistory");
                var store = Get<ChatSessionStore>(window, "sessionStore");
                Set(window, "sessionStore", null);
                try
                {
                    Call(window, "UpdateDeleteSessionButton");
                    Assert.IsFalse(Get<Button>(window, "deleteSession").Enabled);
                    CompleteOnSta((Task)Call(window, "DeleteSelectedSessionAsync"));
                    Assert.AreEqual(1, confirmations);
                }
                finally { Set(window, "sessionStore", store); }
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void DeleteStorageConflictRetainsCurrentDraftAndRestoresTheInterface()
        {
            using (var runtime = new RuntimeScope())
            using (var window = LoadedWindow(runtime.Session))
            {
                var original = Get<ChatSessionState>(window, "currentSession");
                Get<System.Windows.Controls.TextBox>(window, "prompt").Text = "keep local draft";
                Call(window, "SaveCurrentSession");
                Assert.IsTrue(Get<ChatPersistenceWorker>(window, "persistenceWorker").Flush(5000));
                using (var other = new ChatSessionStore(ChatWindow.HistoryPath()))
                {
                    var winner = other.List(original.Scope)[0]; winner.Draft = "other host"; other.Save(winner);
                }
                ChatWindow.ShowNotice = (owner, text, caption, buttons, icon) => DialogResult.Yes;
                CompleteOnSta((Task)Call(window, "DeleteSelectedSessionAsync"));
                Assert.AreSame(original, Get<ChatSessionState>(window, "currentSession"));
                Assert.AreEqual("keep local draft", Get<System.Windows.Controls.TextBox>(window, "prompt").Text);
                Assert.IsTrue(Get<TableLayoutPanel>(window, "rootLayout").Enabled);
                Assert.IsFalse(Get<bool>(window, "loadingScope")); Assert.IsFalse(Get<bool>(window, "loadingSession"));
                StringAssert.StartsWith(Get<Label>(window, "status").Text, UiText.Get("Conversation not deleted: "));
                Assert.AreEqual("other host", Get<ChatSessionStore>(window, "sessionStore").List(original.Scope)[0].Draft);
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void DeletePendingBehindAnInFlightSaveBlocksSessionChangesAndLateDraftSaves()
        {
            using (var runtime = new RuntimeScope())
            using (var window = LoadedWindow(runtime.Session))
            using (var inFlight = new ManualResetEventSlim())
            using (var release = new ManualResetEventSlim())
            {
                var deleted = Get<ChatSessionState>(window, "currentSession");
                var previous = Get<ChatPersistenceWorker>(window, "persistenceWorker");
                Assert.IsTrue(previous.Flush(5000)); previous.Dispose();
                int writes = 0;
                using (var worker = new ChatPersistenceWorker(ChatWindow.HistoryPath(), (snapshot, version, error) => {
                    if (Interlocked.Increment(ref writes) == 1) { inFlight.Set(); release.Wait(5000); }
                }))
                {
                    Set(window, "persistenceWorker", worker);
                    Task deletion = null;
                    var nextScope = AddScope(window, "temporary:next");
                    try
                    {
                        Call(window, "SaveCurrentSession"); Assert.IsTrue(inFlight.Wait(5000));
                        ChatWindow.ShowNotice = (owner, text, caption, buttons, icon) => DialogResult.Yes;
                        deletion = (Task)Call(window, "DeleteSelectedSessionAsync");
                        Assert.IsFalse(deletion.IsCompleted);
                        Assert.IsTrue(Get<bool>(window, "loadingScope"));
                        Assert.IsFalse(Get<TableLayoutPanel>(window, "rootLayout").Enabled);
                        Assert.IsFalse(Get<Button>(window, "deleteSession").Enabled);
                        Call(window, "NewSession", (object)null); Call(window, "SaveCurrentSession");
                        Assert.AreSame(deleted, Get<ChatSessionState>(window, "currentSession"));
                        Assert.ThrowsException<TargetInvocationException>(() => Call(window, "EnsureCurrentScope"));
                        Get<ComboBox>(window, "scopePicker").SelectedItem = nextScope;
                        Assert.AreSame(deleted, Get<ChatSessionState>(window, "currentSession"));
                    }
                    finally { release.Set(); }
                    CompleteOnSta(deletion);
                    CompleteScopeLoad(window);
                    Assert.IsTrue(worker.Flush(5000));
                    Assert.AreNotSame(deleted, Get<ChatSessionState>(window, "currentSession"));
                    Assert.AreEqual("temporary:next", Get<ChatSessionState>(window, "currentSession").Scope);
                    Assert.IsTrue(Get<TableLayoutPanel>(window, "rootLayout").Enabled);
                    var saved = Get<ChatSessionStore>(window, "sessionStore").List(deleted.Scope);
                    Assert.AreEqual(1, saved.Count); Assert.AreNotEqual(deleted.Id, saved[0].Id);
                }
                Set(window, "persistenceWorker", null);
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void DeleteReportsDatabaseSuccessWhenReplacementHistoryCannotBeRendered()
        {
            using (var runtime = new RuntimeScope())
            using (var window = LoadedWindow(runtime.Session))
            {
                var deleted = Get<ChatSessionState>(window, "currentSession");
                var corrupt = new ChatSessionState { Scope = deleted.Scope, Title = "Unreadable fixture", MessagesJson = "{" };
                Get<List<ChatSessionState>>(window, "scopeSessions").Add(corrupt);
                Call(window, "RefreshHistory");
                ChatWindow.ShowNotice = (owner, text, caption, buttons, icon) => DialogResult.Yes;
                CompleteOnSta((Task)Call(window, "DeleteSelectedSessionAsync"));
                Assert.IsNull(Get<ChatSessionState>(window, "currentSession"));
                Assert.AreEqual(-1, Get<ListBox>(window, "sessionList").SelectedIndex);
                Assert.IsFalse(Get<Button>(window, "send").Enabled);
                Call(window, "UpdateBudgetControls"); Assert.IsFalse(Get<Button>(window, "send").Enabled);
                StringAssert.StartsWith(Get<Label>(window, "status").Text, UiText.Get("Conversation deleted from local history"));
                StringAssert.Contains(Get<Label>(window, "status").Text, UiText.Get("History unavailable: "));
                Assert.ThrowsException<TargetInvocationException>(() => Call(window, "EnsureCurrentScope"));
                Assert.AreEqual(0, Get<ChatSessionStore>(window, "sessionStore").List(deleted.Scope).Count);
                Call(window, "NewSession", (object)null);
                Assert.IsNotNull(Get<ChatSessionState>(window, "currentSession"));
                Assert.IsFalse(Get<bool>(window, "sessionViewUnavailable"));
                Assert.IsTrue(Get<Button>(window, "send").Enabled);
            }
        }

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
                var b = AddScope(window, @"C:\Owned\ScopeB.xlsm");
                var c = AddScope(window, @"C:\Owned\ScopeC.xlsm");
                var scopes = Get<ComboBox>(window, "scopePicker");
                var first = new TaskCompletionSource<ChatSessionStore.ScopeSnapshot>();
                var latest = new TaskCompletionSource<ChatSessionStore.ScopeSnapshot>();
                var requested = new List<string>();
                window.ReadScope = (path, scope, includeSessions) => {
                    requested.Add(scope);
                    return scope == @"C:\Owned\ScopeB.xlsm" ? first.Task : latest.Task;
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
                        Sessions = new List<ChatSessionState> { new ChatSessionState { Scope = @"C:\Owned\ScopeB.xlsm", Title = "Stale" } }, Memory = "stale memory"
                    });
                    CompleteOnSta(staleLoad);
                    Assert.AreSame(original, Get<ChatSessionState>(window, "currentSession"));
                    Assert.IsTrue(Get<bool>(window, "loadingScope"));
                    CollectionAssert.AreEqual(new[] { @"C:\Owned\ScopeB.xlsm", @"C:\Owned\ScopeC.xlsm" }, requested);
                    var expected = new ChatSessionState { Scope = @"C:\Owned\ScopeC.xlsm", Title = "Current" };
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
                var target = AddScope(window, @"C:\Owned\ScopeB.xlsm");
                Get<ComboBox>(window, "scopePicker").SelectedItem = target;
                var read = new TaskCompletionSource<ChatSessionStore.ScopeSnapshot>();
                window.ReadScope = (path, scope, includeSessions) => read.Task;
                Call(window, "ChangeScope");
                Task loading = Get<Task>(window, "scopeLoad");
                window.Dispose();
                read.SetResult(new ChatSessionStore.ScopeSnapshot {
                    Sessions = new List<ChatSessionState> { new ChatSessionState { Scope = @"C:\Owned\ScopeB.xlsm" } }, Memory = "ignored"
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
                runtime.Vbe.VBProjects[0].FileName = "";
                using (var window = LoadedWindow(runtime.Session))
                {
                    Get<System.Windows.Forms.TextBox>(window, "memoryEditor").Text = "available without SQLite";
                    Call(window, "SaveProjectMemory");
                    StringAssert.Contains(Get<System.Windows.Forms.Label>(window, "status").Text, UiText.Get("Temporary document: VBAi history and project notes stay in memory until the document is saved."));
                    Assert.AreEqual("available without SQLite", Get<Dictionary<string, string>>(window, "transientMemory")[Get<ChatSessionState>(window, "currentSession").Scope]);
                    Call(window, "EnsureCurrentScope");
                    Get<System.Windows.Forms.ComboBox>(window, "scopePicker").SelectedIndex = -1; Call(window, "ChangeScope"); CompleteScopeLoad(window);
                    var failure = Assert.ThrowsException<System.Reflection.TargetInvocationException>(() => Call(window, "EnsureCurrentScope"));
                    Assert.IsInstanceOfType(failure.InnerException, typeof(InvalidOperationException));
                }
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
