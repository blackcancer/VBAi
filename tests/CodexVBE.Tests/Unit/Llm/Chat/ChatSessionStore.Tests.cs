namespace CodexVBE.Tests.Unit
{
    using System;
    using System.IO;
    using System.Linq;
    using System.Web.Script.Serialization;
    using CodexVBE;
    using CodexVBE.Tests.Infrastructure;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class ChatSessionStoreTests
    {
        [TestMethod]
        public void NativeStorePersistsRichConversationDraftsAndKeepsScopesIsolated()
        {
            using(var scope=new LlmBoundaryScope())
            {
                string path=Path.Combine(scope.Root,"history.db");
                var session=new ChatSessionState {Scope="é-scope",Title="O'Brien",Pinned=true,Archived=true,Mode=ChatMode.Plan,
                    Provider="Fixture",Model="fixture-model",CodexThreadId="thread",ResumeContext="resume",MessagesJson="[]",Draft="draft",
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
                    Assert.AreEqual("resume",read.ResumeContext);Assert.AreEqual("[]",read.MessagesJson);Assert.AreEqual("draft",read.Draft);
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
