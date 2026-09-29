using System;
using System.IO;
using System.Linq;
using VBAi;
using VBAi.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ChatSessionConcurrencyTests
    {
        [TestMethod]
        public async System.Threading.Tasks.Task ScopeReaderLoadsIndependentObjectsAndDoesNotCreateMissingHistory()
        {
            using (var scope = new LlmBoundaryScope())
            {
                string path = Path.Combine(scope.Root, "reader.db");
                var seed = new ChatSessionState { Scope = "fixture", Draft = "saved" };
                using (var store = new ChatSessionStore(path))
                {
                    store.Save(seed);
                    store.SaveMemory("fixture", "project notes");
                }
                var loaded = await ChatSessionStore.ReadScopeAsync(path, "fixture", true);
                Assert.AreEqual("saved", loaded.Sessions.Single().Draft);
                Assert.AreEqual(seed.StorageVersion, loaded.Sessions.Single().StorageVersion);
                Assert.AreNotSame(seed, loaded.Sessions.Single());
                Assert.AreEqual("project notes", loaded.Memory);
                var memoryOnly = await ChatSessionStore.ReadScopeAsync(path, "fixture", false);
                Assert.IsNull(memoryOnly.Sessions);
                Assert.AreEqual("project notes", memoryOnly.Memory);
                string absent = Path.Combine(scope.Root, "missing-reader.db");
                await Assert.ThrowsExceptionAsync<IOException>(() => ChatSessionStore.ReadScopeAsync(absent, "fixture", true));
                Assert.IsFalse(File.Exists(absent), "A history read must never create an empty replacement database.");
            }
        }

        [TestMethod]
        public void TwoConnectionsCannotSilentlyReplaceTheSameConversation()
        {
            using (var scope = new LlmBoundaryScope())
            using (var first = new ChatSessionStore(Path.Combine(scope.Root, "history.db")))
            using (var second = new ChatSessionStore(Path.Combine(scope.Root, "history.db")))
            {
                var seed = new ChatSessionState { Scope = "fixture", Draft = "base" };
                first.Save(seed);
                var a = first.List("fixture").Single();
                var b = second.List("fixture").Single();
                a.Draft = "first host"; first.Save(a);
                b.Draft = "second host";
                Assert.ThrowsException<IOException>(() => second.Save(b));
                Assert.AreEqual("second host", b.Draft, "The unsuccessful writer retains its draft in memory.");
                Assert.AreEqual("first host", second.List("fixture").Single().Draft);
                a.Draft = "next edit"; first.Save(a);
                Assert.AreEqual("next edit", second.List("fixture").Single().Draft);
                Assert.ThrowsException<IOException>(() => second.Save(new ChatSessionState { Id = seed.Id, Scope = "fixture" }));
                second.Save(new ChatSessionState { Scope = "fixture", Draft = "independent" });
                Assert.AreEqual(2, first.List("fixture").Count);
            }
        }
    }
}
