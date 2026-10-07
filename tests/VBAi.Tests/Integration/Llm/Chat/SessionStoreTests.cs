using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;

namespace VBAi.Tests.Integration
{
    /// <summary>Vérifie la persistance SQLite des conversations et mémoires de session.</summary>
    [TestClass]
    [TestCategory("LocalIntegration")]
    public sealed class SessionStoreTests
    {
        /// <summary>Code et empreinte renvoyés par le lecteur de test des signets.</summary>
        public sealed class BookmarkCodeSnapshot { public string Code { get; set; } public string Sha256 { get; set; } }
        [TestMethod]
        public void BookmarksPersistByMacroAndRefuseStaleSourceAfterNewSession()
        {
            string root = Path.Combine(Path.GetTempPath(), "VBAi-VSTest", Guid.NewGuid().ToString("N"));
            string database = Path.Combine(root, "bookmarks.sqlite"), macro = Path.Combine(root, "été.xlsm"), other = Path.Combine(root, "other.xlsm");
            string sha = "original"; Request selected = null;
            Func<Request, Response> execute = r =>
            {
                if (r.Command == "read_module") return Response.Success(new BookmarkCodeSnapshot { Code = "abc", Sha256 = sha });
                selected = r; return Response.Success(new { Selected = true });
            };
            try
            {
                var first = new VbeNavigationHistory(null, execute, database);
                first.Bookmark(new Request { Project = macro, Action = "add", Query = "L'été", Module = "Main", ExpectedSha256 = sha, StartLine = 1, StartColumn = 2 });
                var second = new VbeNavigationHistory(null, execute, database);
                dynamic list = second.Bookmark(new Request { Project = macro.ToUpperInvariant(), Action = "list" });
                Assert.AreEqual(1, list.Bookmarks.Count); Assert.AreEqual("L'été", (string)list.Bookmarks[0].Name);
                Assert.AreEqual(0, ((dynamic)second.Bookmark(new Request { Project = other, Action = "list" })).Bookmarks.Count);
                second.Bookmark(new Request { Project = macro, Action = "go", Query = "L'ÉTÉ" });
                Assert.AreEqual(macro, selected.Project); Assert.AreEqual(2, selected.StartColumn);
                sha = "changed"; selected = null;
                Assert.ThrowsException<InvalidOperationException>(() => second.Bookmark(new Request { Project = macro, Action = "go", Query = "L'été" }));
                Assert.IsNull(selected);
                Assert.IsTrue((bool)((dynamic)second.Bookmark(new Request { Project = macro, Action = "remove", Query = "L'été" })).Removed);
                Assert.IsFalse((bool)((dynamic)second.Bookmark(new Request { Project = macro, Action = "remove", Query = "L'été" })).Removed);
            }
            finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
        }

        /// <summary>Conserve les textes Unicode et isole les portées après fermeture puis réouverture du stockage.</summary>
        [TestMethod]
        public void SqlitePersistsUnicodeAndSeparatesScopesAcrossReopen()
        {
            var root = Path.Combine(Path.GetTempPath(), "VBAi-VSTest", Guid.NewGuid().ToString("N"));
            var path = Path.Combine(root, "sessions.sqlite");
            try
            {
                using (var store = new ChatSessionStore(path))
                {
                    store.Save(new ChatSessionState { Scope = "Classeur Été.xlsm", Title = "L'été d'O'Brien" });
                    store.Save(new ChatSessionState { Scope = "Autre.xlsm", Title = "Isolé" });
                    store.SaveMemory("Classeur Été.xlsm", "mémoire: 'été'");
                }
                using (var store = new ChatSessionStore(path))
                {
                    var sessions = store.List("Classeur Été.xlsm");
                    Assert.AreEqual(1, sessions.Count);
                    Assert.AreEqual("L'été d'O'Brien", sessions[0].Title);
                    Assert.AreEqual("mémoire: 'été'", store.ReadMemory("Classeur Été.xlsm"));
                    Assert.AreEqual("", store.ReadMemory("Autre.xlsm"));
                }
            }
            finally
            {
                if (Directory.Exists(root)) Directory.Delete(root, true);
            }
        }
    }
}
