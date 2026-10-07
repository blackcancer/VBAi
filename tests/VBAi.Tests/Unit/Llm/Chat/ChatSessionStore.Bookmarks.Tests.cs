using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Linq;

namespace VBAi.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class ChatBookmarkStoreTests
    {
        [TestMethod]
        public void BookmarkTransactionsEnforceCapacityAllowUpdatesAndRollBackFailedWrites()
        {
            string root = Path.Combine(Path.GetTempPath(), "CodexBookmark-" + Guid.NewGuid().ToString("N"));
            try
            {
                using (var store = new ChatSessionStore(Path.Combine(root, "bookmarks.db")))
                {
                    Assert.AreEqual(0, store.ListBookmarks("macro").Count);
                    Assert.IsFalse(store.RemoveBookmark("macro", "missing"));
                    for (int i = 0; i < 200; i++) store.SaveBookmark("macro", new CodeBookmark { Name = "mark" + i, Module = "M", Sha256 = "sha", Line = 1, Column = 1 });
                    Assert.ThrowsException<InvalidOperationException>(() => store.SaveBookmark("macro", new CodeBookmark { Name = "new" }));
                    Assert.AreEqual(200, store.ListBookmarks("MACRO").Count);
                    store.SaveBookmark("MACRO", new CodeBookmark { Name = "MARK0", Module = "Updated" });
                    Assert.AreEqual("Updated", store.ListBookmarks("macro").Single(x => x.Name == "MARK0").Module);
                    Assert.IsTrue(store.RemoveBookmark("macro", "mark0"));
                    Assert.IsFalse(store.RemoveBookmark("macro", "mark0"));
                    Assert.ThrowsException<NullReferenceException>(() => store.SaveBookmark("macro", new CodeBookmark()));
                    Assert.AreEqual(199, store.ListBookmarks("macro").Count);
                    var native = store.StepNative; int calls = 0;
                    store.StepNative = statement => ++calls == 2 ? 101 : native(statement);
                    Assert.IsFalse(store.RemoveBookmark("macro", "missing"));
                    store.StepNative = native;
                    Assert.AreEqual(199, store.ListBookmarks("macro").Count);
                }
            }
            finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
        }
    }
}
