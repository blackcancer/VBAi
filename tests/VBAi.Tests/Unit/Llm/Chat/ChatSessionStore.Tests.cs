namespace VBAi.Tests.Unit
{
    using System;
    using System.IO;
    using System.Linq;
    using System.Web.Script.Serialization;
    using VBAi;
    using VBAi.Tests.Infrastructure;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class ChatSessionStoreTests
    {
        private static ChatSessionStore.PromotionRow Promotion(string scope, string text)
        {
            var session = new ChatSessionState { Scope = scope, Draft = text };
            return new ChatSessionStore.PromotionRow { Id = session.Id, Title = session.Title, Payload = new JavaScriptSerializer().Serialize(session) };
        }

        [TestMethod]
        public void PromotionClaimsDestinationAtomicallyAcrossTwoOwnedSqliteConnections()
        {
            using (var scope = new LlmBoundaryScope())
            using (var ready = new System.Threading.ManualResetEventSlim())
            using (var claiming = new System.Threading.ManualResetEventSlim())
            using (var competing = new System.Threading.ManualResetEventSlim())
            {
                string path = Path.Combine(scope.Root, "claim.db"), key = @"C:\OWNED\ATOMIC.XLSM";
                using (var first = new ChatSessionStore(path))
                {
                    var loser = Promotion(key, "competing host");
                    var second = System.Threading.Tasks.Task.Run(() => {
                        using (var store = new ChatSessionStore(path))
                        {
                            ready.Set(); Assert.IsTrue(claiming.Wait(5000)); competing.Set();
                            return store.PromoteEmptyScope(key, new[] { loser }, "competing notes");
                        }
                    });
                    Assert.IsTrue(ready.Wait(5000));
                    var native = first.StepNative; int steps = 0;
                    first.StepNative = statement => {
                        int result = native(statement);
                        if (++steps == 1) { claiming.Set(); Assert.IsTrue(competing.Wait(5000)); Assert.IsFalse(second.Wait(50), "A competing claim must wait for the first BEGIN IMMEDIATE transaction."); }
                        return result;
                    };
                    var winner = Promotion(key, "winning host");
                    try { Assert.IsNotNull(first.PromoteEmptyScope(key, new[] { winner }, "winning notes")); }
                    finally { first.StepNative = native; claiming.Set(); }
                    Assert.IsNull(second.GetAwaiter().GetResult());
                    Assert.AreEqual(winner.Id, first.List(key).Single().Id); Assert.AreEqual("winning notes", first.ReadMemory(key));
                }
            }
        }

        [TestMethod]
        public void PromotionRollsBackEveryInitialRowAndNotesWhenAnyWriteFails()
        {
            using (var scope = new LlmBoundaryScope())
            using (var store = new ChatSessionStore(Path.Combine(scope.Root, "rollback.db")))
            {
                string key = @"C:\OWNED\ROLLBACK.XLSM"; var native = store.StepNative; int steps = 0;
                var rows = new[] { Promotion(key, "first"), Promotion(key, "second") };
                store.StepNative = statement => ++steps == 5 ? throw new IOException("Owned note failure after two inserts") : native(statement);
                try { Assert.ThrowsException<IOException>(() => store.PromoteEmptyScope(key, rows, "notes")); }
                finally { store.StepNative = native; }
                Assert.AreEqual(0, store.List(key).Count); Assert.AreEqual("", store.ReadMemory(key));
                Assert.IsNotNull(store.PromoteEmptyScope(key, rows, "notes"), "A confirmed rollback permits a fresh claim.");
                Assert.AreEqual(2, store.List(key).Count); Assert.AreEqual("notes", store.ReadMemory(key));
            }
        }
        [TestMethod]
        public void TemporaryScopesNeverReadOrWriteSqliteIncludingNotesAndBookmarks()
        {
            using (var scope = new LlmBoundaryScope())
            using (var store = new ChatSessionStore(Path.Combine(scope.Root, "transient.db")))
            {
                store.StepNative = statement => { Assert.Fail("A temporary scope reached SQLite."); return 0; };
                var session = new ChatSessionState { Scope = "temporary:" + Guid.NewGuid().ToString("N"), Draft = "private" };
                store.Save(session); Assert.IsNull(session.StorageVersion);
                Assert.IsNull(store.SavePayload(session.Id, session.Scope, "title", "invalid json", null));
                store.SaveMemory(session.Scope, "notes"); store.SaveBookmark(session.Scope, null);
                store.Delete(session.Id, session.Scope, null);
                Assert.AreEqual(0, store.List(session.Scope).Count); Assert.IsFalse(store.HasSessions(session.Scope));
                Assert.IsFalse(store.HasScopeData(session.Scope));
                Assert.AreEqual("", store.ReadMemory(session.Scope)); Assert.AreEqual(0, store.ListBookmarks(session.Scope).Count);
                Assert.IsFalse(store.RemoveBookmark(session.Scope, "bookmark"));
                var loaded = ChatSessionStore.ReadScopeAsync(Path.Combine(scope.Root, "does-not-exist.db"), session.Scope, true).GetAwaiter().GetResult();
                Assert.AreEqual(0, loaded.Sessions.Count); Assert.AreEqual("", loaded.Memory);
            }
        }

        [TestMethod]
        public void DeleteRemovesOnlyTheMatchingConversationAndPreservesMemoryAndOtherScopes()
        {
            using (var scope = new LlmBoundaryScope())
            using (var store = new ChatSessionStore(Path.Combine(scope.Root, "delete.db")))
            {
                var first = new ChatSessionState { Scope = "first" };
                var sibling = new ChatSessionState { Scope = "first" };
                var other = new ChatSessionState { Scope = "other" };
                store.Save(first); store.Save(sibling); store.Save(other); store.SaveMemory("first", "retained memory");
                store.Delete(first.Id, "wrong scope", first.StorageVersion);
                Assert.AreEqual(2, store.List("first").Count);
                store.Delete(first.Id, first.Scope, first.StorageVersion);
                Assert.AreEqual(sibling.Id, store.List("first").Single().Id);
                Assert.AreEqual(other.Id, store.List("other").Single().Id);
                Assert.AreEqual("retained memory", store.ReadMemory("first"));
                store.Delete(first.Id, first.Scope, first.StorageVersion);
                store.Delete(new ChatSessionState().Id, first.Scope, null);
            }
        }

        [TestMethod]
        public void DeleteRejectsStaleOrMissingRevisionAndDeletedSnapshotsCannotRecreateTheRow()
        {
            using (var scope = new LlmBoundaryScope())
            using (var store = new ChatSessionStore(Path.Combine(scope.Root, "delete-conflict.db")))
            {
                var session = new ChatSessionState { Scope = "fixture" };
                store.Save(session);
                string staleVersion = session.StorageVersion;
                session.Title = "other host"; store.Save(session);
                Assert.ThrowsException<IOException>(() => store.Delete(session.Id, session.Scope, staleVersion));
                Assert.ThrowsException<IOException>(() => store.Delete(session.Id, session.Scope, null));
                Assert.AreEqual("other host", store.List(session.Scope).Single().Title);
                store.Delete(session.Id, session.Scope, session.StorageVersion);
                Assert.ThrowsException<IOException>(() => store.Save(session));
                Assert.AreEqual(0, store.List(session.Scope).Count);
                Assert.ThrowsException<ArgumentException>(() => store.Delete(null, session.Scope, session.StorageVersion));
            }
        }

        [TestMethod]
        public void NativeStorePersistsRichConversationDraftsAndKeepsScopesIsolated()
        {
            using(var scope=new LlmBoundaryScope())
            {
                string path=Path.Combine(scope.Root,"history.db");
                var session=new ChatSessionState {Scope="é-scope",Title="O'Brien",Pinned=true,Archived=true,Mode=ChatMode.Plan,
                    Provider="Fixture",Model="fixture-model",CodexThreadId="thread",CodexThreadHome=ProviderSessionStorage.CodexHome,CodexDeveloperInstructionsHash="instructions-sha",ResumeContext="resume",MessagesJson="[]",Draft="draft",
                    DraftAttachments=new[] {new ChatAttachment {Label="L",Text="T",Project="P",Module="M",Sha256="sha",StartLine=8}},
                    DraftReferences=new[] {new VbeChatReference {Project="P",Module="M",Kind="Module"}}};
                session.Entries.Add(new ChatEntry {Speaker="Assistant",Text="été",StreamId="stream",TurnId="turn",AttachedMemory="memory",
                    References=session.DraftReferences,Attachments=session.DraftAttachments,Change=new CodeChange("P","M","old","before-sha","new","after-sha",1)});
                using(var store=new ChatSessionStore(path))
                {
                    store.Save(session);store.SaveMemory(session.Scope,null);Assert.AreEqual("",store.ReadMemory(session.Scope));
                    store.SaveMemory(session.Scope,"mémoire");Assert.AreEqual("",store.ReadMemory("missing"));
                }
                using(var store=new ChatSessionStore(path))
                {
                    var read=store.List(session.Scope).Single();Assert.AreEqual(session.Id,read.Id);Assert.AreEqual("O'Brien",read.Title);
                    Assert.IsTrue(read.Pinned);Assert.IsTrue(read.Archived);Assert.AreEqual(ChatMode.Plan,read.Mode);
                    Assert.AreEqual("Fixture",read.Provider);Assert.AreEqual("fixture-model",read.Model);Assert.AreEqual("thread",read.CodexThreadId);
                    Assert.AreEqual(ProviderSessionStorage.CodexHome,read.CodexThreadHome);Assert.AreEqual("instructions-sha",read.CodexDeveloperInstructionsHash);Assert.AreEqual("resume",read.ResumeContext);Assert.AreEqual("[]",read.MessagesJson);Assert.AreEqual("draft",read.Draft);
                    Assert.AreEqual("é-scope",read.Scope);Assert.AreEqual("sha",read.DraftAttachments[0].Sha256);
                    Assert.AreEqual("stream",read.Entries[0].StreamId);Assert.AreEqual("turn",read.Entries[0].TurnId);Assert.AreEqual("été",read.Entries[0].Text);
                    Assert.AreEqual("before-sha",read.Entries[0].Change.BeforeSha256);Assert.AreEqual("after-sha",read.Entries[0].Change.AfterSha256);
                    Assert.AreEqual(1,read.Entries[0].Change.AfterLineCount);Assert.AreEqual("mémoire",store.ReadMemory(session.Scope));
                    Assert.AreEqual(0,store.List("other").Count);
                    Assert.AreEqual("★ O'Brien",read.ToString());Assert.AreEqual("O'Brien",read.DisplayTitle);
                    var empty=new ChatSessionState();Assert.AreEqual("New conversation",empty.DisplayTitle);Assert.AreEqual("New conversation",empty.ToString());
                    store.Dispose();store.Dispose();
                }
            }
        }

        [TestMethod]
        public void NativeRowsIgnoreNullOrMismatchedPayloadsAndReportMalformedJson()
        {
            using(var scope=new LlmBoundaryScope())
            using(var store=new ChatSessionStore(Path.Combine(scope.Root,"history.db")))
            {
                Sql(store,"INSERT INTO chat_sessions VALUES (?1,?2,?3,?4,?5)","null","S","T","1","null");
                Sql(store,"INSERT INTO chat_sessions VALUES (?1,?2,?3,?4,?5)","mismatch","S","T","2",new JavaScriptSerializer().Serialize(new ChatSessionState {Scope="other"}));
                Assert.AreEqual(0,store.List("S").Count);
                Sql(store,"INSERT INTO chat_sessions VALUES (?1,?2,?3,?4,?5)","invalid","S","T","3","{invalid}");
                Assert.ThrowsException<ArgumentException>(()=>store.List("S"));
                Sql(store,"DROP TABLE project_memory");Sql(store,"CREATE TABLE project_memory(scope TEXT PRIMARY KEY,content TEXT)");
                Sql(store,"INSERT INTO project_memory(scope,content) VALUES (?1,NULL)","legacy");Assert.AreEqual("",store.ReadMemory("legacy"));
            }
        }

        [TestMethod]
        public void NativePrepareBindConstraintAndInitializationFailuresReleaseOwnedHandles()
        {
            using(var scope=new LlmBoundaryScope())
            {
                Assert.ThrowsException<IOException>(()=>new ChatSessionStore(scope.Root));
                string path=Path.Combine(scope.Root,"history.db");
                using(var store=new ChatSessionStore(path))
                {
                    Assert.ThrowsException<IOException>(()=>Sql(store,"THIS IS INVALID SQL"));
                    Assert.ThrowsException<IOException>(()=>Call(store,"Prepare","SELECT 1",new[] {"unexpected binding"}));
                    var statement=(IDisposable)Call(store,"Prepare","SELECT ?1",new string[] {null});statement.Dispose();statement.Dispose();
                    Sql(store,"CREATE TABLE fixture_unique(id TEXT PRIMARY KEY)");Sql(store,"INSERT INTO fixture_unique VALUES ('same')");
                    Assert.ThrowsException<IOException>(()=>Sql(store,"INSERT INTO fixture_unique VALUES ('same')"));
                    store.SaveMemory("healthy","after-errors");Assert.AreEqual("after-errors",store.ReadMemory("healthy"));
                    Sql(store,"DROP TABLE chat_sessions");Sql(store,"CREATE TABLE chat_sessions(id TEXT)");
                }
                var initialization=Assert.ThrowsException<IOException>(()=>new ChatSessionStore(path));
                StringAssert.Contains(initialization.Message,"scope");
                File.Delete(path);Assert.IsFalse(File.Exists(path));
            }
        }
    }
}
