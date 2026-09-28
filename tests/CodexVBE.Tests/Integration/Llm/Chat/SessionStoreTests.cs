using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using CodexVBE;

namespace CodexVBE.Tests.Integration
{
    /// <summary>Vérifie la persistance SQLite des conversations et mémoires de session.</summary>
    [TestClass]
    [TestCategory("LocalIntegration")]
    public sealed class SessionStoreTests
    {
        /// <summary>Conserve les textes Unicode et isole les portées après fermeture puis réouverture du stockage.</summary>
        [TestMethod]
        public void SqlitePersistsUnicodeAndSeparatesScopesAcrossReopen()
        {
            var root = Path.Combine(Path.GetTempPath(), "CodexVBE-VSTest", Guid.NewGuid().ToString("N"));
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
