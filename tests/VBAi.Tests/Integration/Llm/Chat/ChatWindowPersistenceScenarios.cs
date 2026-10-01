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

    /// <summary>Vérifie la persistance des conversations, brouillons et mémoires par portée.</summary>
    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class ChatWindowStateTests
    {
        /// <summary>Enregistre puis relit le brouillon, les messages Unicode et les entrées de la session courante.</summary>
        [TestMethod]
        [STATestMethod]
        public void CurrentSessionPersistsDraftMessagesAndEntriesInIsolatedStore()
        {
            string path = Path.Combine(Path.GetTempPath(), "VBAi-Chat-" + Guid.NewGuid().ToString("N"), "chat.db");
            try
            {
                using (var store = new ChatSessionStore(path))
                using (var window = Surfaces())
                {
                    var session = new ChatSessionState
                    {
                        Scope = @"C:\OWNED\PERSISTENCE.XLSM",
                        Title = "Essai éè"
                    };
                    Set(window, "sessionStore", store);
                    Set(window, "currentSession", session);
                    Question(window, "Brouillon éè");
                    Get<List<object>>(window, "messages").Add(new Dictionary<string, object> { ["role"] = "user", ["content"] = "Question éè" });
                    Call(window, "AddTranscriptMessage", "Assistant", "Réponse éè");
                    Call(window, "SaveCurrentSession");
                    var restored = store.List(session.Scope);
                    Assert.AreEqual(1, restored.Count);
                    Assert.AreEqual("Essai éè", restored[0].Title);
                    Assert.AreEqual("Brouillon éè", restored[0].Draft);
                    StringAssert.Contains(restored[0].MessagesJson, "Question éè");
                    Assert.AreEqual("Réponse éè", restored[0].Entries[0].Text);
                    Set(window, "currentSession", null);
                    Set(window, "sessionStore", null);
                }
            }
            finally
            {
                string directory = Path.GetDirectoryName(path);
                if (Directory.Exists(directory))
                {
                    foreach (string file in Directory.GetFiles(directory))
                        File.Delete(file);
                    Directory.Delete(directory, false);
                }
            }
        }

        /// <summary>Charge la conversation et la mémoire d’une portée puis restaure la session mise en cache.</summary>
        [TestMethod]
        [STATestMethod]
        public void ScopeSwitchLoadsSavedConversationAndMemoryThenRestoresCachedScope()
        {
            string path = Path.Combine(Path.GetTempPath(), "VBAi-Scope-" + Guid.NewGuid().ToString("N"), "chat.db");
            try
            {
                using (var store = new ChatSessionStore(path))
                using (var window = ReadyCodexWindow(new ChatSessionState { Scope = @"C:\OWNED\SCOPEA.XLSM", Title = "A" }))
                {
                    window.ModelCatalogueOverride = provider => Task.FromResult(new LlmModelOption[0]);
                    var original = Get<ChatSessionState>(window, "currentSession");
                    Get<List<ChatSessionState>>(window, "scopeSessions").Add(original);
                    var scopes = Get<ComboBox>(window, "scopePicker");
                    AddScope(window, @"C:\OWNED\SCOPEA.XLSM");
                    AddScope(window, @"C:\OWNED\SCOPEB.XLSM");
                    store.Save(new ChatSessionState { Scope = @"C:\OWNED\SCOPEB.XLSM", Title = "B", Draft = "Brouillon B" });
                    store.SaveMemory(@"C:\OWNED\SCOPEB.XLSM", "Mémoire B");
                    Set(window, "sessionStore", store);
                    scopes.SelectedIndex = 1;
                    Call(window, "ChangeScope"); CompleteScopeLoad(window);
                    Assert.AreEqual("B", Get<ChatSessionState>(window, "currentSession").Title);
                    Assert.AreEqual("Brouillon B", Get<object>(window, "prompt").GetType().GetProperty("Text").GetValue(Get<object>(window, "prompt"), null));
                    Assert.AreEqual("Mémoire B", Get<TextBox>(window, "memoryEditor").Text);
                    scopes.SelectedIndex = 0;
                    Call(window, "ChangeScope"); CompleteScopeLoad(window);
                    Assert.AreSame(original, Get<ChatSessionState>(window, "currentSession"));
                    Set(window, "currentSession", null);
                    Set(window, "sessionStore", null);
                }
            }
            finally
            {
                string directory = Path.GetDirectoryName(path);
                if (Directory.Exists(directory))
                {
                    foreach (string file in Directory.GetFiles(directory))
                        File.Delete(file);
                    Directory.Delete(directory, false);
                }
            }
        }
    }
}
