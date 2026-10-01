using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Scenarios
{
    [TestClass, TestCategory("Scenario"), DoNotParallelize]
    public sealed class ChatSessionStorePromotionTests
    {
        private static ChatSessionStore.PromotionRow Row(string scope, string draft)
        {
            var session = new ChatSessionState { Scope = scope, Draft = draft };
            return new ChatSessionStore.PromotionRow {
                Id = session.Id, Title = session.Title, Payload = new JavaScriptSerializer().Serialize(session)
            };
        }

        [TestMethod]
        public void InitialClaimCommitsTheWholeConversationBatchAndNotes()
        {
            using (var boundary = new LlmBoundaryScope())
            using (var store = new ChatSessionStore(Path.Combine(boundary.Root, "batch.db")))
            {
                string key = @"C:\OWNED\BATCH.XLSM";
                var rows = new[] { Row(key, "first draft"), Row(key, "second draft") };
                var versions = store.PromoteEmptyScope(key, rows, "project notes");
                Assert.IsNotNull(versions);
                Assert.AreEqual(2, versions.Count);
                Assert.IsTrue(rows.All(row => !string.IsNullOrEmpty(versions[row.Id])));
                var saved = store.List(key);
                Assert.AreEqual(2, saved.Count);
                Assert.IsTrue(saved.Any(item => item.Id == rows[0].Id && item.Draft == "first draft" && item.StorageVersion == versions[item.Id]));
                Assert.IsTrue(saved.Any(item => item.Id == rows[1].Id && item.Draft == "second draft" && item.StorageVersion == versions[item.Id]));
                Assert.AreEqual("project notes", store.ReadMemory(key));
                Assert.IsNull(store.PromoteEmptyScope(key, new[] { Row(key, "late claim") }, "late notes"));
                Assert.AreEqual(2, store.List(key).Count);
                Assert.AreEqual("project notes", store.ReadMemory(key));
            }
        }

        [TestMethod]
        public void NotesOrBookmarksAloneOccupyTheDestinationAndPreserveItsData()
        {
            using (var boundary = new LlmBoundaryScope())
            using (var store = new ChatSessionStore(Path.Combine(boundary.Root, "occupied.db")))
            {
                string notesKey = @"C:\OWNED\NOTES.XLSM";
                store.SaveMemory(notesKey, "existing notes from another document");
                Assert.IsFalse(store.HasSessions(notesKey));
                Assert.IsNull(store.PromoteEmptyScope(notesKey, new[] { Row(notesKey, "private draft") }, "replacement notes"));
                Assert.AreEqual("existing notes from another document", store.ReadMemory(notesKey));
                Assert.AreEqual(0, store.List(notesKey).Count);

                string bookmarkKey = @"C:\OWNED\BOOKMARK.XLSM";
                store.SaveBookmark(bookmarkKey, new CodeBookmark { Name = "existing", Module = "Module1", Sha256 = "sha", Line = 1, Column = 1 });
                Assert.IsFalse(store.HasSessions(bookmarkKey));
                Assert.IsNull(store.PromoteEmptyScope(bookmarkKey, new[] { Row(bookmarkKey, "private draft") }, "new notes"));
                Assert.AreEqual("existing", store.ListBookmarks(bookmarkKey).Single().Name);
                Assert.AreEqual(0, store.List(bookmarkKey).Count);
                Assert.AreEqual("", store.ReadMemory(bookmarkKey));
            }
        }

        [TestMethod]
        public void FailedInsertOrCommitRollsBackAllRowsAndNotes()
        {
            using (var boundary = new LlmBoundaryScope())
            using (var store = new ChatSessionStore(Path.Combine(boundary.Root, "rollback.db")))
            {
                var nativeStep = store.StepNative;
                foreach (int failingStep in new[] { 3, 6 }) // first INSERT, then COMMIT before SQLite executes it
                {
                    string key = @"C:\OWNED\ROLLBACK" + failingStep + ".XLSM";
                    var rows = new[] { Row(key, "first"), Row(key, "second") };
                    int steps = 0;
                    store.StepNative = statement => ++steps == failingStep
                        ? throw new IOException("Owned promotion write failure") : nativeStep(statement);
                    try { Assert.ThrowsException<IOException>(() => store.PromoteEmptyScope(key, rows, "notes")); }
                    finally { store.StepNative = nativeStep; }
                    Assert.AreEqual(0, store.List(key).Count);
                    Assert.AreEqual("", store.ReadMemory(key));
                    Assert.IsFalse(store.HasScopeData(key));
                    Assert.IsNotNull(store.PromoteEmptyScope(key, rows, "notes"), "A confirmed rollback permits a fresh claim.");
                    Assert.AreEqual(2, store.List(key).Count);
                    Assert.AreEqual("notes", store.ReadMemory(key));
                }
            }
        }

        [TestMethod]
        public void CompetingConnectionCannotWriteWhileInitialClaimHoldsTheSqliteWriterLock()
        {
            using (var boundary = new LlmBoundaryScope())
            using (var ready = new ManualResetEventSlim())
            using (var acquired = new ManualResetEventSlim())
            using (var attempting = new ManualResetEventSlim())
            using (var first = new ChatSessionStore(Path.Combine(boundary.Root, "concurrent.db")))
            {
                string path = first.DatabasePath, key = @"C:\OWNED\CONCURRENT.XLSM";
                var contender = Task.Run(() => {
                    using (var second = new ChatSessionStore(path))
                    {
                        ready.Set();
                        if (!acquired.Wait(2000)) throw new TimeoutException("The first claim never acquired its writer lock.");
                        attempting.Set();
                        second.SaveMemory(key, "competing notes");
                    }
                });
                Assert.IsTrue(ready.Wait(2000));
                var nativeStep = first.StepNative;
                int steps = 0;
                first.StepNative = statement => {
                    int result = nativeStep(statement);
                    if (++steps == 1)
                    {
                        acquired.Set();
                        Assert.IsTrue(attempting.Wait(2000));
                        Assert.IsFalse(contender.Wait(80), "The second connection wrote while BEGIN IMMEDIATE held the writer lock.");
                    }
                    return result;
                };
                var winner = Row(key, "claimed first");
                try { Assert.IsNotNull(first.PromoteEmptyScope(key, new[] { winner }, "initial notes")); }
                finally { first.StepNative = nativeStep; acquired.Set(); }
                Assert.IsTrue(contender.Wait(2000));
                Assert.AreEqual(winner.Id, first.List(key).Single().Id);
                Assert.AreEqual("competing notes", first.ReadMemory(key));
            }
        }
    }
}
