using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class ChatWindowStateTests
    {
        private const BindingFlags Fields = BindingFlags.Instance | BindingFlags.NonPublic;
        private const BindingFlags Methods = BindingFlags.Instance | BindingFlags.Static | BindingFlags.NonPublic;

        private static T Get<T>(ChatWindow window, string field)
        {
            var info = typeof(ChatWindow).GetField(field, Fields);
            Assert.IsNotNull(info, "Missing ChatWindow field " + field);
            return (T)info.GetValue(window);
        }

        private static void Set(ChatWindow window, string field, object value)
        {
            var info = typeof(ChatWindow).GetField(field, Fields);
            Assert.IsNotNull(info, "Missing ChatWindow field " + field);
            info.SetValue(window, value);
        }

        private static object Call(ChatWindow window, string method, params object[] args)
        {
            var info = typeof(ChatWindow).GetMethod(method, Methods);
            Assert.IsNotNull(info, "Missing ChatWindow method " + method);
            return info.Invoke(window, args);
        }

        private static ChatWindow Surfaces()
        {
            var window = new ChatWindow();
            Call(window, "InitializeShell");
            Call(window, "InitializeComposer", new object[] { null });
            Call(window, "InitializeTranscript");
            return window;
        }

        [TestMethod]
        [STATestMethod]
        public void ConstructorAndLocalSurfacesCreateComposerTranscriptAndModes()
        {
            using (var window = Surfaces())
            {
                Assert.AreEqual(3, Get<ComboBox>(window, "modePicker").Items.Count);
                Assert.AreEqual(ChatMode.Agent, Get<ComboBox>(window, "modePicker").SelectedItem);
                Assert.IsNotNull(Get<object>(window, "prompt"));
                Assert.IsNotNull(Get<object>(window, "conversationScroll"));
                Assert.IsFalse(Get<Panel>(window, "historyPanel").Visible);
                Assert.IsFalse(Get<ProgressBar>(window, "activityBar").Visible);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void ProviderEffortChoicesFollowModelAndBusyState()
        {
            using (var window = Surfaces())
            {
                Set(window, "currentSession", new ChatSessionState { Effort = "high" });
                var codexModel = new LlmModelOption("gpt-test", "Test", false, "medium", new[] {
                    new LlmEffortOption("medium", "Medium"), new LlmEffortOption("high", "High") });
                Call(window, "UpdateEfforts", LlmProvider.All[0], codexModel);
                var effort = Get<ComboBox>(window, "effortPicker");
                Assert.AreEqual(2, effort.Items.Count);
                Assert.AreEqual("high", ((LlmEffortOption)effort.SelectedItem).Id);
                Assert.IsTrue(effort.Enabled);
                string sendBefore = Get<Button>(window, "send").Text;
                Call(window, "SetBusy", true);
                Assert.IsFalse(effort.Enabled);
                Assert.AreNotEqual(sendBefore, Get<Button>(window, "send").Text);
                Call(window, "SetBusy", false);
                Assert.AreNotEqual("", Get<Button>(window, "send").Text);
                Call(window, "UpdateEfforts", LlmProvider.All[1], codexModel);
                Assert.AreEqual(0, effort.Items.Count);
                Assert.IsFalse(effort.Enabled);
                Set(window, "currentSession", null);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void HistoryFiltersArchivedSessionsAndOrdersPinnedFirst()
        {
            using (var window = Surfaces())
            {
                var sessions = Get<List<ChatSessionState>>(window, "scopeSessions");
                var normal = new ChatSessionState { Title = "Alpha" };
                var pinned = new ChatSessionState { Title = "Alpha pinned", Pinned = true };
                var archived = new ChatSessionState { Title = "Alpha archived", Archived = true };
                sessions.Add(normal); sessions.Add(pinned); sessions.Add(archived);
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
                var session = new ChatSessionState { Scope = "temporary:test" };
                Set(window, "currentSession", session);
                Call(window, "RenameFromQuestion", "First line\nSecond line");
                Assert.AreEqual("First line Second line", session.Title);
                Assert.AreEqual(session.Title, Get<Label>(window, "sessionTitle").Text);
                var prompt = Get<object>(window, "prompt");
                prompt.GetType().GetProperty("Text").SetValue(prompt, "unsent draft", null);
                Get<List<object>>(window, "messages").Add(new Dictionary<string, object> {
                    ["role"] = "user", ["content"] = "hello" });
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
                var session = new ChatSessionState { Scope = "temporary:test" };
                Set(window, "currentSession", session);
                Get<TextBox>(window, "chatTitleEditor").Text = "  " + new string('a', 130) + "  ";
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
                messages.Add(new Dictionary<string, object> { ["role"] = "assistant", ["tool_calls"] =
                    new object[] { new Dictionary<string, object> { ["id"] = "pending" } } });
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
                messages.Add(new Dictionary<string, object> { ["role"] = "assistant", ["tool_calls"] =
                    new object[] { new Dictionary<string, object> { ["id"] = "done" } } });
                messages.Add(new Dictionary<string, object> { ["role"] = "tool", ["tool_call_id"] = "done" });
                Call(window, "RepairInterruptedToolHistory");
                Assert.AreEqual(3, messages.Count);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void TranscriptStreamsUpdateInPlaceAndAvoidDuplicateFinalMessage()
        {
            using (var window = Surfaces())
            {
                Call(window, "ReceiveChatUpdate", "summary", "empty", " ", false);
                var entries = Get<List<ChatEntry>>(window, "transcriptEntries");
                Assert.AreEqual(0, entries.Count);
                Call(window, "ReceiveChatUpdate", "final", "answer", "Hel", false);
                Call(window, "ReceiveChatUpdate", "final", "answer", "lo", true);
                Assert.AreEqual(1, entries.Count);
                Assert.AreEqual("lo", entries[0].Text);
                Call(window, "CompleteAssistantResponse", "lo");
                Assert.AreEqual(1, entries.Count);
                Call(window, "CompleteAssistantResponse", "different");
                Assert.AreEqual(2, entries.Count);
                Call(window, "ClearTranscript");
                Assert.AreEqual(0, entries.Count);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void ComposerTokenBoundaryAndDraftContextLimitAreEnforced()
        {
            using (var window = Surfaces())
            {
                var token = "#Project.Module";
                Assert.AreEqual(true, Call(window, "ContainsToken", "Use " + token + " now", token));
                Assert.AreEqual(false, Call(window, "ContainsToken", "x" + token + "suffix", token));
                Assert.AreEqual(true, Call(window, "IsReferenceChar", '_'));
                Assert.AreEqual(false, Call(window, "IsReferenceChar", '-'));
                var attachments = Get<List<ChatAttachment>>(window, "draftAttachments");
                attachments.Add(new ChatAttachment { Label = "selection", Text = "code" });
                var prepared = (ChatAttachment[])Call(window, "PrepareAttachments", "question");
                Assert.AreEqual(1, prepared.Length);
                Assert.AreEqual("selection", prepared[0].Label);
                attachments[0].Text = new string('x', 48001);
                var error = Assert.ThrowsException<TargetInvocationException>(() =>
                    Call(window, "PrepareAttachments", "question"));
                Assert.IsInstanceOfType(error.InnerException, typeof(InvalidOperationException));
            }
        }

        [TestMethod]
        [STATestMethod]
        public void EditorActionSeedsCommandWhileBusyStateBlocksIt()
        {
            using (var window = Surfaces())
            {
                Call(window, "SetBusy", true);
                window.PrepareEditorAction("/corriger");
                var prompt = Get<object>(window, "prompt");
                Assert.AreEqual("", prompt.GetType().GetProperty("Text").GetValue(prompt, null));
                Call(window, "SetBusy", false);
                window.PrepareEditorAction("/corriger");
                Assert.AreEqual("/corriger ", prompt.GetType().GetProperty("Text").GetValue(prompt, null));
                Assert.AreEqual(ChatMode.Agent, Get<ComboBox>(window, "modePicker").SelectedItem);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void CurrentReferencesKeepLatestExactTokenAndIgnoreEmbeddedMatches()
        {
            using (var window = Surfaces())
            {
                var selected = Get<List<VbeChatReference>>(window, "selectedReferences");
                selected.Add(new VbeChatReference { Project = "P", Module = "M", Kind = "Module", Sha256 = "old" });
                selected.Add(new VbeChatReference { Project = "P", Module = "M", Kind = "Module", Sha256 = "new" });
                var exact = (VbeChatReference[])Call(window, "CurrentReferences", "Use #P.M for this change");
                Assert.AreEqual(1, exact.Length);
                Assert.AreEqual("new", exact[0].Sha256);
                var embedded = (VbeChatReference[])Call(window, "CurrentReferences", "prefix#P.Msuffix");
                Assert.AreEqual(0, embedded.Length);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void ContextChipsRepresentMemoryAndDraftAttachmentsAndCanRemoveDraft()
        {
            using (var window = Surfaces())
            {
                Set(window, "projectMemory", "local note");
                Get<CheckBox>(window, "attachMemory").Checked = true;
                var attachments = Get<List<ChatAttachment>>(window, "draftAttachments");
                attachments.Add(new ChatAttachment { Label = "Selected code", Text = "Sub A()" });
                Call(window, "RefreshContextChips");
                var chips = Get<FlowLayoutPanel>(window, "contextChips");
                Assert.AreEqual(2, chips.Controls.Count);
                typeof(Control).GetMethod("OnClick", BindingFlags.Instance | BindingFlags.NonPublic)
                    .Invoke(chips.Controls[1], new object[] { EventArgs.Empty });
                Assert.AreEqual(0, attachments.Count);
                Assert.AreEqual(1, chips.Controls.Count);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void ModeSelectionUpdatesSessionOnlyWhenIdle()
        {
            using (var window = Surfaces())
            {
                var session = new ChatSessionState { Mode = ChatMode.Agent };
                Set(window, "currentSession", session);
                Get<ComboBox>(window, "modePicker").SelectedItem = ChatMode.Plan;
                Call(window, "ModePicker_SelectedIndexChanged", null, EventArgs.Empty);
                Assert.AreEqual(ChatMode.Plan, session.Mode);
                Call(window, "SetBusy", true);
                Get<ComboBox>(window, "modePicker").SelectedItem = ChatMode.Discussion;
                Call(window, "ModePicker_SelectedIndexChanged", null, EventArgs.Empty);
                Assert.AreEqual(ChatMode.Plan, session.Mode);
                Set(window, "currentSession", null);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void MarkdownTranscriptAndContextPreviewRetainStructuredEntries()
        {
            using (var window = Surfaces())
            {
                Call(window, "AddTranscriptMessage", "Assistant", "# Heading\n- item\n```vba\nDebug.Print 1\n```");
                Assert.AreEqual(1, Get<List<ChatEntry>>(window, "transcriptEntries").Count);
                Assert.AreEqual(1, ((IDictionary)Get<object>(window, "entryViews")).Count);
                Get<List<ChatAttachment>>(window, "draftAttachments").Add(
                    new ChatAttachment { Label = "Selection", Text = "VBA code" });
                Call(window, "RefreshContextPreview");
                Assert.AreEqual(2, Get<FlowLayoutPanel>(window, "contextPreview").Controls.Count);
            }
        }
    }
}
