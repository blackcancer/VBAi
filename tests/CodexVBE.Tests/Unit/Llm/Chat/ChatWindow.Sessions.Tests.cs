namespace CodexVBE.Tests.Unit
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
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class ChatWindowStateTests
    {
        [TestMethod]
        [STATestMethod]
        public void HistoryFiltersArchivedSessionsAndOrdersPinnedFirst()
        {
            using (var window = Surfaces())
            {
                var sessions = Get<List<ChatSessionState>>(window, "scopeSessions");
                var normal = new ChatSessionState
                {
                    Title = "Alpha"
                };
                var pinned = new ChatSessionState
                {
                    Title = "Alpha pinned",
                    Pinned = true
                };
                var archived = new ChatSessionState
                {
                    Title = "Alpha archived",
                    Archived = true
                };
                sessions.Add(normal);
                sessions.Add(pinned);
                sessions.Add(archived);
                Get<TextBox>(window, "historySearch").Text = "Alpha";
                Call(window, "RefreshHistory");
                var list = Get<ListBox>(window, "sessionList");
                Assert.AreEqual(2, list.Items.Count);
                Assert.AreSame(pinned, list.Items[0]);
                Get<CheckBox>(window, "showArchived").Checked = true;
                Call(window, "RefreshHistory");
                Assert.AreEqual(3, list.Items.Count);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void SessionTitleAndDraftAreSavedWithoutExternalStore()
        {
            using (var window = Surfaces())
            {
                var session = new ChatSessionState
                {
                    Scope = "temporary:test"
                };
                Set(window, "currentSession", session);
                Call(window, "RenameFromQuestion", "First line\nSecond line");
                Assert.AreEqual("First line Second line", session.Title);
                Assert.AreEqual(session.Title, Get<Label>(window, "sessionTitle").Text);
                var prompt = Get<object>(window, "prompt");
                prompt.GetType().GetProperty("Text").SetValue(prompt, "unsent draft", null);
                Get<List<object>>(window, "messages").Add(new Dictionary<string, object> { ["role"] = "user", ["content"] = "hello" });
                Call(window, "SaveCurrentSession");
                Assert.AreEqual("unsent draft", session.Draft);
                StringAssert.Contains(session.MessagesJson, "hello");
                Assert.AreEqual(0, session.DraftAttachments.Length);
                Set(window, "currentSession", null);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void ManualRenameTrimsAndCapsTitle()
        {
            using (var window = Surfaces())
            {
                var session = new ChatSessionState
                {
                    Scope = "temporary:test"
                };
                Set(window, "currentSession", session);
                Get<TextBox>(window, "chatTitleEditor").Text = "  " + new string ('a', 130) + "  ";
                Call(window, "RenameCurrentChat");
                Assert.AreEqual(120, session.Title.Length);
                Assert.AreEqual(session.Title, Get<Label>(window, "sessionTitle").Text);
                Set(window, "currentSession", null);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void InterruptedToolHistoryDropsOnlyUnfinishedAssistantTail()
        {
            using (var window = Surfaces())
            {
                var messages = Get<List<object>>(window, "messages");
                messages.Add(new Dictionary<string, object> { ["role"] = "system", ["content"] = "rules" });
                messages.Add(new Dictionary<string, object> { ["role"] = "user", ["content"] = "question" });
                messages.Add(new Dictionary<string, object> { ["role"] = "assistant", ["tool_calls"] = new object[] { new Dictionary<string, object> { ["id"] = "pending" } } });
                Call(window, "RepairInterruptedToolHistory");
                Assert.AreEqual(3, messages.Count);
                var repaired = new JavaScriptSerializer().Serialize(messages[2]);
                StringAssert.Contains(repaired, "assistant");
            }
        }

        [TestMethod]
        [STATestMethod]
        public void CompletedToolHistoryIsPreserved()
        {
            using (var window = Surfaces())
            {
                var messages = Get<List<object>>(window, "messages");
                messages.Add(new Dictionary<string, object> { ["role"] = "user", ["content"] = "question" });
                messages.Add(new Dictionary<string, object> { ["role"] = "assistant", ["tool_calls"] = new object[] { new Dictionary<string, object> { ["id"] = "done" } } });
                messages.Add(new Dictionary<string, object> { ["role"] = "tool", ["tool_call_id"] = "done" });
                Call(window, "RepairInterruptedToolHistory");
                Assert.AreEqual(3, messages.Count);
            }
        }
    }
}
