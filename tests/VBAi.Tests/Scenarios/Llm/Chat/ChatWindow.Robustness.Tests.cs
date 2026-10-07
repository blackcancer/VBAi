using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;
using System.Windows.Threading;

namespace VBAi.Tests.Unit
{
    public sealed partial class ChatWindowStateTests
    {
        [STATestMethod, TestCategory("Unit")]
        public void SerializationFailurePreservesSavedHistoryAndDoesNotInterruptCleanup()
        {
            using (var runtime = new RuntimeScope())
            using (var window = LoadedWindow(runtime.Session))
            {
                Call(window, "SaveCurrentSession");
                var session = Get<ChatSessionState>(window, "currentSession");
                string saved = session.MessagesJson;
                var timer = Get<DispatcherTimer>(window, "saveTimer");
                Get<JavaScriptSerializer>(window, "json").MaxJsonLength = 64;
                Get<List<object>>(window, "messages").Add(new { role = "user", content = new string('x', 256) });
                Call(window, "SaveCurrentSession");
                Assert.IsTrue(Get<bool>(window, "storageFailed"));
                Assert.AreEqual(saved, session.MessagesJson);
                timer.Start();
                window.Dispose();
                Assert.IsFalse(timer.IsEnabled);
                Assert.IsNull(Get<ChatSessionStore>(window, "sessionStore"));
                Assert.IsTrue(Get<bool>(window, "runtimeDisposed"));
                using (var check = new ChatSessionStore(System.IO.Path.Combine(runtime.Root, "chat.db")))
                    Assert.AreEqual(saved, check.List(session.Scope).Single(s => s.Id == session.Id).MessagesJson);
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void StreamingCoalescesFragmentsAndFlushesBeforeFinalizationAndSave()
        {
            using (var window = Surfaces())
            {
                var session = new ChatSessionState { Scope = "temporary:stream" };
                Set(window, "currentSession", session);
                for (int i = 0; i < 2000; i++) Call(window, "ReceiveChatUpdate", "final", "stream", "x", false);
                var entry = Get<List<ChatEntry>>(window, "transcriptEntries").Single();
                Assert.AreEqual("", entry.Text, "Rendering is deferred until the bounded flush.");
                Assert.AreEqual(1, Get<Dictionary<ChatEntry, StringBuilder>>(window, "pendingStreamText").Count);
                Call(window, "SaveCurrentSession");
                Assert.AreEqual(new string('x', 2000), session.Entries.Single().Text);
                Assert.IsFalse(Get<DispatcherTimer>(window, "streamRenderTimer").IsEnabled);
                Call(window, "ReceiveChatUpdate", "final", "stream", "y", false);
                Call(window, "ReceiveChatUpdate", "final", "stream", null, true);
                Assert.AreEqual(new string('x', 2000) + "y", entry.Text);
                Call(window, "CompleteAssistantResponse", entry.Text);
                Assert.AreEqual(1, Get<List<ChatEntry>>(window, "transcriptEntries").Count);
                Set(window, "currentSession", null);
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void FollowLatestKeepsOnlyOnePendingScrollAndCancelsItOnSessionClear()
        {
            using (var window = Surfaces())
            {
                Call(window, "AddEntry", new ChatEntry { Speaker = "Assistant", Text = "fixture" });
                var operation = Get<DispatcherOperation>(window, "pendingFollow");
                for (int i = 0; i < 100; i++) Call(window, "FollowLatest");
                Assert.AreSame(operation, Get<DispatcherOperation>(window, "pendingFollow"));
                Call(window, "ClearTranscript");
                Assert.AreEqual(DispatcherOperationStatus.Aborted, operation.Status);
                Assert.IsNull(Get<DispatcherOperation>(window, "pendingFollow"));
            }
        }
    }
}
