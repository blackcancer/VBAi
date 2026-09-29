using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Web.Script.Serialization;
using VBAi;
using VBAi.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class ChatPersistenceWorkerTests
    {
        private static readonly JavaScriptSerializer Json = new JavaScriptSerializer();

        private static ChatPersistenceWorker.Snapshot Capture(ChatSessionState session) =>
            new ChatPersistenceWorker.Snapshot(session, Json.Serialize(session));

        private static void Sql(ChatSessionStore store, string command)
        {
            typeof(ChatSessionStore).GetMethod("Execute", BindingFlags.Instance | BindingFlags.NonPublic)
                .Invoke(store, new object[] { command, new string[0] });
        }

        [TestMethod]
        public void SnapshotRetainsTheUiValuesCapturedBeforeLaterMutation()
        {
            using (var scope = new LlmBoundaryScope())
            {
                string path = Path.Combine(scope.Root, "chat.db");
                var session = new ChatSessionState { Scope = "fixture", Title = "before", Draft = "before draft" };
                var snapshot = Capture(session);
                session.Title = "after"; session.Draft = "after draft";
                using (var worker = new ChatPersistenceWorker(path, null))
                {
                    worker.Enqueue(snapshot);
                    Assert.IsTrue(worker.Flush(5000));
                }
                using (var store = new ChatSessionStore(path))
                {
                    var saved = store.List("fixture").Single();
                    Assert.AreEqual("before", saved.Title);
                    Assert.AreEqual("before draft", saved.Draft);
                    Assert.AreEqual("after draft", session.Draft);
                }
            }
        }

        [TestMethod]
        public void SameWriterCoalescesQueuedSnapshotsAndChainsSqliteVersions()
        {
            using (var scope = new LlmBoundaryScope())
            using (var firstCompleted = new ManualResetEventSlim())
            using (var releaseFirst = new ManualResetEventSlim())
            {
                string path = Path.Combine(scope.Root, "chat.db");
                var session = new ChatSessionState { Scope = "fixture", Draft = "first" };
                var completions = new ConcurrentQueue<Tuple<string, string, Exception>>();
                int count = 0;
                using (var worker = new ChatPersistenceWorker(path, (snapshot, version, error) =>
                {
                    completions.Enqueue(Tuple.Create(snapshot.Payload, version, error));
                    if (Interlocked.Increment(ref count) == 1)
                    {
                        firstCompleted.Set();
                        releaseFirst.Wait(5000);
                    }
                }))
                {
                    try
                    {
                        worker.Enqueue(Capture(session));
                        Assert.IsTrue(firstCompleted.Wait(5000));
                        session.Draft = "superseded"; worker.Enqueue(Capture(session));
                        session.Draft = "latest"; worker.Enqueue(Capture(session));
                    }
                    finally { releaseFirst.Set(); }
                    Assert.IsTrue(worker.Flush(5000));
                }
                var results = completions.ToArray();
                Assert.AreEqual(2, results.Length, "Only the latest queued snapshot follows the in-flight write.");
                Assert.IsTrue(results.All(result => result.Item3 == null));
                Assert.AreNotEqual(results[0].Item2, results[1].Item2);
                using (var store = new ChatSessionStore(path))
                    Assert.AreEqual("latest", store.List("fixture").Single().Draft);
            }
        }

        [TestMethod]
        public void ConflictingWorkersRetainRecoveryAndNeverOverwriteTheWinningConversation()
        {
            using (var scope = new LlmBoundaryScope())
            {
                string path = Path.Combine(scope.Root, "chat.db");
                using (var store = new ChatSessionStore(path))
                    store.Save(new ChatSessionState { Scope = "fixture", Draft = "base" });
                ChatSessionState first, second;
                using (var store = new ChatSessionStore(path))
                {
                    first = store.List("fixture").Single();
                    second = store.List("fixture").Single();
                }
                first.Draft = "winner"; second.Draft = "loser";
                var failures = new ConcurrentQueue<Exception>();
                using (var winningWorker = new ChatPersistenceWorker(path, (snapshot, version, error) => failures.Enqueue(error)))
                using (var losingWorker = new ChatPersistenceWorker(path, (snapshot, version, error) => failures.Enqueue(error)))
                {
                    winningWorker.Enqueue(Capture(first));
                    Assert.IsTrue(winningWorker.Flush(5000));
                    losingWorker.Enqueue(Capture(second));
                    Assert.IsTrue(losingWorker.Flush(5000));
                    string recovery = Path.Combine(losingWorker.RecoveryDirectory, second.WriterId + ".json");
                    Assert.IsTrue(File.Exists(recovery));
                    StringAssert.Contains(File.ReadAllText(recovery), "loser");
                    second.Draft = "later loser";
                    losingWorker.Enqueue(Capture(second));
                    Assert.IsTrue(losingWorker.Flush(5000));
                    Assert.IsTrue(File.Exists(recovery));
                    StringAssert.Contains(File.ReadAllText(recovery), "later loser");
                    using (var store = new ChatSessionStore(path))
                        Assert.AreEqual("winner", store.List("fixture").Single().Draft);
                }
                var outcomes = failures.ToArray();
                Assert.AreEqual(3, outcomes.Length);
                Assert.IsNull(outcomes[0]);
                Assert.IsInstanceOfType(outcomes[1], typeof(IOException));
                Assert.IsInstanceOfType(outcomes[2], typeof(IOException));
            }
        }

        [TestMethod]
        public void LockedSqliteWriteKeepsEnqueueFastFlushBoundedAndRecoveryUntilCommit()
        {
            using (var scope = new LlmBoundaryScope())
            {
                string path = Path.Combine(scope.Root, "chat.db");
                var session = new ChatSessionState { Scope = "fixture", Draft = "first" };
                using (var locker = new ChatSessionStore(path))
                using (var worker = new ChatPersistenceWorker(path, null))
                {
                    Sql(locker, "BEGIN EXCLUSIVE");
                    try
                    {
                        var watch = Stopwatch.StartNew();
                        worker.Enqueue(Capture(session));
                        Assert.IsTrue(SpinWait.SpinUntil(() => File.Exists(Path.Combine(worker.RecoveryDirectory, session.WriterId + ".json")), 500));
                        Assert.IsFalse(worker.Flush(25));
                        session.Draft = "latest"; worker.Enqueue(Capture(session));
                        Assert.IsTrue(watch.ElapsedMilliseconds < 1000, "The UI path must not wait for SQLite's busy timeout.");
                        Assert.IsTrue(File.Exists(Path.Combine(worker.RecoveryDirectory, session.WriterId + ".json")));
                    }
                    finally { Sql(locker, "COMMIT"); }
                    Assert.IsTrue(worker.Flush(5000));
                    Assert.IsFalse(File.Exists(Path.Combine(worker.RecoveryDirectory, session.WriterId + ".json")));
                    Assert.AreEqual("latest", locker.List("fixture").Single().Draft);
                }
            }
        }

        [TestMethod]
        public void TimedOutCloseKeepsTheLatestPendingSnapshotInASeparateRecoveryFile()
        {
            using (var scope = new LlmBoundaryScope())
            using (var firstCompleted = new ManualResetEventSlim())
            using (var releaseFirst = new ManualResetEventSlim())
            {
                string path = Path.Combine(scope.Root, "chat.db");
                var session = new ChatSessionState { Scope = "fixture", Draft = "older" };
                using (var worker = new ChatPersistenceWorker(path, (snapshot, version, error) =>
                {
                    firstCompleted.Set();
                    releaseFirst.Wait();
                }))
                {
                    string closingFile = null;
                    try
                    {
                        worker.Enqueue(Capture(session));
                        Assert.IsTrue(firstCompleted.Wait(5000), "The first write must reach its callback before close.");
                        session.Draft = "superseded"; worker.Enqueue(Capture(session));
                        session.Draft = "latest pending"; worker.Enqueue(Capture(session));
                        worker.Dispose();
                        string[] closing = Directory.GetFiles(worker.RecoveryDirectory, "closing-*.json");
                        Assert.AreEqual(1, closing.Length);
                        closingFile = closing[0];
                        var record = Json.DeserializeObject(File.ReadAllText(closingFile)) as Dictionary<string, object>;
                        Assert.IsNotNull(record);
                        var recovered = Json.Deserialize<ChatSessionState>((string)record["Payload"]);
                        Assert.AreEqual("latest pending", recovered.Draft);
                        using (var store = new ChatSessionStore(path))
                            Assert.AreEqual("older", store.List("fixture").Single().Draft);
                    }
                    finally { releaseFirst.Set(); }
                    Assert.IsTrue(worker.Flush(5000));
                    Assert.IsTrue(File.Exists(closingFile), "A later worker cleanup must not delete the shutdown copy.");
                    Assert.IsFalse(File.Exists(Path.Combine(worker.RecoveryDirectory, session.WriterId + ".json")));
                    using (var store = new ChatSessionStore(path))
                        Assert.AreEqual("latest pending", store.List("fixture").Single().Draft);
                }
            }
        }
    }
}
