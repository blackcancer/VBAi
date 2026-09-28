using System.Collections.Generic;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    public sealed partial class ChatWindowStateTests
    {
        [TestMethod, TestCategory("Unit")]
        public void StoredSessionsWithoutPrivacyMarkerAreIdentifiedAsLegacy()
        {
            var legacy = ChatSessionStore.DecodeSession("{\"Scope\":\"A\",\"MessagesJson\":\"[]\"}");
            Assert.AreEqual(0, legacy.ReadAccessPolicyVersion);
            var current = new ChatSessionState { Scope = "A", ReadProjectGrants = new[] { "B" }, BudgetPaused = true, PausedTurnId = "turn" };
            var restored = ChatSessionStore.DecodeSession(new JavaScriptSerializer().Serialize(current));
            Assert.AreEqual(1, restored.ReadAccessPolicyVersion);
            CollectionAssert.AreEqual(current.ReadProjectGrants, restored.ReadProjectGrants);
            Assert.IsTrue(restored.BudgetPaused);
            Assert.AreEqual("turn", restored.PausedTurnId);
        }

        [STATestMethod, TestCategory("Unit")]
        public void LegacyProviderContextIsClearedWhileLocalTranscriptRemainsAvailable()
        {
            var state = new ChatSessionState { Scope = "A", ReadAccessPolicyVersion = 0,
                CodexThreadId = "legacy-thread", ResumeContext = "PRIVATE-B", BudgetPaused = true,
                MessagesJson = "[{\"role\":\"assistant\",\"content\":\"PRIVATE-B\"}]" };
            using (var window = ReadyCodexWindow(state))
            {
                Call(window, "AddEntry", new ChatEntry { Speaker = "Assistant", Text = "PRIVATE-B" });
                Get<List<object>>(window, "messages").Add(new { role = "assistant", content = "PRIVATE-B" });
                Call(window, "MigrateProviderPrivacy");
                Assert.AreEqual(0, Get<List<object>>(window, "messages").Count);
                Assert.AreEqual("[]", state.MessagesJson);
                Assert.IsNull(state.CodexThreadId);
                Assert.IsNull(state.ResumeContext);
                Assert.IsFalse(state.BudgetPaused);
                Assert.AreEqual(1, state.ReadAccessPolicyVersion);
                var entries = Get<List<ChatEntry>>(window, "transcriptEntries");
                Assert.AreEqual("PRIVATE-B", entries[0].Text);
                Assert.AreEqual(2, entries.Count);
                Call(window, "MigrateProviderPrivacy");
                Assert.AreEqual(2, entries.Count, "Migration should happen only once.");
                Assert.AreEqual(2, state.ProviderHistoryStartIndex);
                Call(window, "ForkChat", entries[0]);
                Assert.AreSame(state, Get<ChatSessionState>(window, "currentSession"));
                var recent = new ChatEntry { Speaker = "Assistant", Text = "ALLOWED-A" };
                Call(window, "AddEntry", recent);
                Call(window, "ForkChat", recent);
                var fork = Get<ChatSessionState>(window, "currentSession");
                Assert.AreNotSame(state, fork);
                Assert.IsFalse(fork.MessagesJson.Contains("PRIVATE-B"));
                Assert.IsFalse(fork.ResumeContext.Contains("PRIVATE-B"));
                StringAssert.Contains(fork.ResumeContext, "ALLOWED-A");
            }
        }
    }
}
