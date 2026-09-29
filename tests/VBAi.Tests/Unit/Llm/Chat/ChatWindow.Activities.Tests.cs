namespace VBAi.Tests.Unit
{
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Linq;
    using System.Windows;
    using System.Windows.Controls;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class ChatWindowStateTests
    {
        /// <summary>Realized activity metadata and new siblings keep existing native controls and explicit disclosure choices.</summary>
        [STATestMethod, TestCategory("Unit")]
        public void RealizedActivityMetadataKeepsControlsSelectionAndDisclosureChoices()
        {
            using (var window = Surfaces())
            {
                Call(window, "ReceiveAgentActivity", new CodexAgentActivity { Id = "first", Kind = "reasoning", Title = "Plan", Detail = "select this text", Status = "inProgress" });
                var owner = Get<List<ChatEntry>>(window, "transcriptEntries")[0];
                using (var host = (ChatDesignerHost)Call(window, "RenderEntry", owner))
                {
                    Get<Dictionary<ChatEntry, FrameworkElement>>(window, "entryViews")[owner] = host;
                    var group = (ChatActivityGroupView)host.View;
                    var step = (ChatActivityStepView)group.section.body.Controls[0];
                    var text = step.detail.content;
                    text.Select(2, 5);
                    // These are explicit choices; status updates must not turn them into automatic defaults.
                    group.section.Expanded = false;
                    step.section.Expanded = false;
                    int collectionChanges = 0;
                    Get<ObservableCollection<object>>(window, "visibleEntries").CollectionChanged += (sender, args) => collectionChanges++;
                    Call(window, "ReceiveAgentActivity", new CodexAgentActivity { Id = "first", Kind = "reasoning", Title = "Updated", Detail = "select this text plus output", Status = "completed", DurationMs = 125 });
                    Assert.AreEqual(0, collectionChanges);
                    Assert.AreSame(step, group.section.body.Controls[0]);
                    Assert.AreSame(text, step.detail.content);
                    Assert.AreEqual(2, text.SelectionStart); Assert.AreEqual(5, text.SelectionLength);
                    Assert.AreEqual("Updated", step.section.Title);
                    Assert.IsFalse(step.section.Expanded); Assert.IsFalse(group.section.Expanded);
                    Call(window, "ReceiveAgentActivity", new CodexAgentActivity { Id = "second", Kind = "commandExecution", Title = "Run", Detail = "output", Status = "inProgress" });
                    Assert.AreEqual(2, group.section.body.Controls.Count);
                    Assert.AreSame(step, group.section.body.Controls[0]);
                    Assert.IsFalse(step.IsDisposed); Assert.AreEqual(0, collectionChanges);
                    Assert.IsFalse(group.section.Expanded);
                    Assert.AreEqual(2, text.SelectionStart); Assert.AreEqual(5, text.SelectionLength);
                    Get<Dictionary<ChatEntry, FrameworkElement>>(window, "entryViews").Remove(owner);
                }
            }
        }

        /// <summary>Automatic expansion does not become a stored user preference when realized metadata changes.</summary>
        [STATestMethod, TestCategory("Unit")]
        public void AutomaticActivityExpansionIsNotRecordedAsUserChoice()
        {
            using (var window = Surfaces())
            {
                Call(window, "ReceiveAgentActivity", new CodexAgentActivity { Id = "activity", Kind = "reasoning", Detail = "result", Status = "completed" });
                var owner = Get<List<ChatEntry>>(window, "transcriptEntries")[0];
                using (var host = (ChatDesignerHost)Call(window, "RenderEntry", owner))
                {
                    Get<Dictionary<ChatEntry, FrameworkElement>>(window, "entryViews")[owner] = host;
                    var group = (ChatActivityGroupView)host.View;
                    var step = (ChatActivityStepView)group.section.body.Controls[0];
                    foreach (string status in new[] { "inProgress", "completed" })
                    {
                        Call(window, "ReceiveAgentActivity", new CodexAgentActivity { Id = "activity", Kind = "reasoning", Detail = "result", Status = status });
                        Assert.AreEqual(status == "inProgress", step.section.Expanded);
                        Assert.AreEqual(status == "inProgress", group.section.Expanded);
                        Assert.AreEqual(0, Get<HashSet<ChatEntry>>(window, "expandedActivitySteps").Count);
                        Assert.AreEqual(0, Get<HashSet<ChatEntry>>(window, "collapsedActivitySteps").Count);
                        Assert.AreEqual(0, Get<HashSet<ChatEntry>>(window, "expandedActivityGroups").Count);
                        Assert.AreEqual(0, Get<HashSet<ChatEntry>>(window, "collapsedActivityGroups").Count);
                    }
                    Get<Dictionary<ChatEntry, FrameworkElement>>(window, "entryViews").Remove(owner);
                }
            }
        }

        /// <summary>Rebuilds absent or disposed activity views and updates metadata without reopening completed streams.</summary>
        [STATestMethod]
        public void ActivityFragmentsRebuildMissingViewsAndPreserveTerminalMetadata()
        {
            foreach (string changed in new[] { "missing", "disposed", "kind", "title", "status", "duration" })
            using (var window = Surfaces())
            {
                Call(window, "ReceiveAgentActivity", new CodexAgentActivity { Id = "stream", Kind = "commandExecution", Title = "Run", Detail = "first", Status = "inProgress" });
                var entry = Get<List<ChatEntry>>(window, "transcriptEntries")[0];
                using (var host = (ChatDesignerHost)Call(window, "RenderEntry", entry))
                {
                    var texts = Get<Dictionary<string, System.Windows.Forms.RichTextBox>>(window, "liveTexts");
                    if (changed == "missing") texts.Remove("stream");
                    if (changed == "disposed") texts["stream"].Dispose();
                    int changes = 0; Get<ObservableCollection<object>>(window, "visibleEntries").CollectionChanged += (s, e) => changes++;
                    Call(window, "ReceiveAgentActivity", new CodexAgentActivity { Id = "stream", Kind = changed == "kind" ? "reasoning" : "commandExecution", Title = changed == "title" ? "Updated" : "Run", Detail = " next", Append = true, Status = changed == "status" ? "completed" : "inProgress", DurationMs = changed == "duration" ? (long?)75 : null });
                    Assert.AreEqual(2, changes, changed); Assert.AreEqual("first next", entry.Activity.Detail);
                }
            }
            using (var window = Surfaces())
            {
                Set(window, "busy", true);
                foreach (string state in new[] { "live", "completed", "no-stream", "native-terminal" })
                {
                    var entry = new ChatEntry { Speaker = "Outil", Text = "Legacy output", StreamId = state == "no-stream" ? null : state };
                    if (state == "native-terminal") entry.Activity = new CodexAgentActivity { Status = "completed", Detail = "Native result" };
                    if (state == "completed") Get<HashSet<string>>(window, "completedStreams").Add(state);
                    using (var host = (ChatDesignerHost)Call(window, "RenderActivityGroup", entry, new List<ChatEntry> { entry }))
                        Assert.AreEqual(state == "live", ((ChatActivityGroupView)host.View).section.Expanded);
                }
                Set(window, "busy", false);
                foreach (var title in new[] { (string)null, "", "Visible title" })
                {
                    var entry = new ChatEntry { Speaker = "Outil", Activity = new CodexAgentActivity { Kind = "commandExecution", Title = title, Status = "completed", Detail = "result" } };
                    using (var host = (ChatDesignerHost)Call(window, "RenderActivityGroup", entry, new List<ChatEntry> { entry }))
                    { Assert.AreEqual(title == "Visible title" ? UiText.Get("Agent activity") + " · 1 · Visible title" : UiText.Get("Agent activity") + " · 1", ((ChatActivityGroupView)host.View).section.Title); }
                }
            }
        }

        /// <summary>Updates hundreds of text fragments without recycling controls belonging to earlier steps.</summary>
        [STATestMethod]
        public void AppendOnlyActivityFragmentsKeepRealizedStepControlsAndExpansion()
        {
            using (var window = Surfaces())
            {
                Call(window, "ReceiveAgentActivity", new CodexAgentActivity { Id = "previous", Kind = "commandExecution", Detail = "saved", Status = "completed" });
                Call(window, "ReceiveAgentActivity", new CodexAgentActivity { Id = "stream", Kind = "reasoning", Title = "Reasoning", Detail = "start", Status = "inProgress" });
                var entries = Get<List<ChatEntry>>(window, "transcriptEntries");
                using (var host = (ChatDesignerHost)Call(window, "RenderEntry", entries[0]))
                {
                    var group = (ChatActivityGroupView)host.View;
                    var controls = group.section.body.Controls.OfType<ChatActivityStepView>().ToArray();
                    controls[0].section.Expanded = true;
                    var live = Get<Dictionary<string, System.Windows.Forms.RichTextBox>>(window, "liveTexts");
                    var original = live["stream"];
                    original.Select(1, 3);
                    int textChanges = 0;
                    original.TextChanged += (s, e) => textChanges++;
                    int changes = 0; Get<ObservableCollection<object>>(window, "visibleEntries").CollectionChanged += (sender, args) => changes++;
                    for (int i = 0; i < 300; i++)
                        Call(window, "ReceiveAgentActivity", new CodexAgentActivity { Id = "stream", Kind = "reasoning", Detail = ".", Status = "inProgress", Append = true });
                    Assert.AreEqual(0, changes); Assert.AreSame(original, live["stream"]);
                    Assert.AreEqual(0, textChanges, "Fragments must not reflow the native text control individually.");
                    Call(window, "FlushStreamText");
                    Assert.AreEqual(1, textChanges);
                    Assert.AreEqual(1, original.SelectionStart); Assert.AreEqual(3, original.SelectionLength);
                    Assert.AreEqual("start" + new string('.', 300), original.Text);
                    Assert.IsTrue(controls[0].section.Expanded); Assert.IsFalse(controls[0].IsDisposed);
                    Assert.AreSame(controls[0], group.section.body.Controls[0]);
                    Call(window, "ReceiveAgentActivity", new CodexAgentActivity { Id = "stream", Kind = "reasoning", Title = "Reasoning", Detail = original.Text, Status = "completed" });
                    Assert.AreEqual(2, changes); Assert.AreEqual("completed", entries[1].Activity.Status);
                }
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void ActivityFlushUsesFinalMetadataAndClearsPendingWorkAcrossSessions()
        {
            using (var window = Surfaces())
            {
                Call(window, "ReceiveAgentActivity", new CodexAgentActivity { Id = "stream", Kind = "reasoning", Title = "Thinking", Detail = "initial", Status = "inProgress" });
                var entry = Get<List<ChatEntry>>(window, "transcriptEntries").Single();
                using (var first = (ChatDesignerHost)Call(window, "RenderEntry", entry))
                {
                    Call(window, "ReceiveAgentActivity", new CodexAgentActivity { Id = "stream", Kind = "reasoning", Detail = " pending", Append = true, Status = "inProgress" });
                    Assert.AreEqual(1, Get<HashSet<ChatEntry>>(window, "pendingActivityText").Count);
                    Call(window, "ReceiveAgentActivity", new CodexAgentActivity { Id = "stream", Kind = "reasoning", Detail = "authoritative final", Status = "completed" });
                    using (var final = (ChatDesignerHost)Call(window, "RenderEntry", entry))
                    {
                        Call(window, "FlushStreamText");
                        Assert.AreEqual("authoritative final", Get<Dictionary<string, System.Windows.Forms.RichTextBox>>(window, "liveTexts")["stream"].Text);
                        Assert.AreEqual(0, Get<HashSet<ChatEntry>>(window, "pendingActivityText").Count);
                        Call(window, "ReceiveAgentActivity", new CodexAgentActivity { Id = "stream", Kind = "reasoning", Detail = " late", Append = true, Status = "inProgress" });
                        Call(window, "ClearTranscript");
                        Assert.AreEqual(0, Get<HashSet<ChatEntry>>(window, "pendingActivityText").Count);
                        Assert.IsFalse(Get<System.Windows.Threading.DispatcherTimer>(window, "streamRenderTimer").IsEnabled);
                        Assert.AreEqual("completed", entry.Activity.Status);
                    }
                }
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void NativeTimelineUpdatesEachStepOnceAndRetainsItsDetailsAcrossRecycling()
        {
            using (var window = Surfaces())
            {
                Call(window, "ClearTranscript");
                Call(window, "ReceiveAgentActivity", new CodexAgentActivity { Id = "r:0", Kind = "reasoning", Title = UiText.Get("Reasoning · summary"), Detail = "Inspecter le module avant modification.", Status = "completed" });
                Call(window, "ReceiveAgentActivity", new CodexAgentActivity { Id = "t", Kind = "dynamicToolCall", Title = "read_module", Detail = "Classeur1 · Module1", Status = "inProgress" });
                Call(window, "ReceiveAgentActivity", new CodexAgentActivity { Id = "c", Kind = "commandExecution", Title = "dotnet test", Detail = "passed", Status = "completed", DurationMs = 1250 });
                Call(window, "ReceiveAgentActivity", new CodexAgentActivity { Id = "t", Kind = "dynamicToolCall", Title = "read_module", Detail = "Classeur1 · Module1", Status = "completed", DurationMs = 30 });
                var entries = Get<List<ChatEntry>>(window, "transcriptEntries");
                Assert.AreEqual(3, entries.Count); Assert.AreEqual(1, Get<ObservableCollection<object>>(window, "visibleEntries").Count);
                var group = ((ChatActivityGroupView)((ChatDesignerHost)Call(window, "RenderEntry", entries[0])).View).section;
                var steps = group.body.Controls.OfType<ChatActivityStepView>().ToArray();
                Assert.AreEqual(3, steps.Length); Assert.IsFalse(steps[1].section.Expanded);
                steps[1].section.Expanded = true;
                Call(window, "ReceiveAgentActivity", new CodexAgentActivity { Id = "t", Kind = "dynamicToolCall", Title = "read_module", Detail = "Changed result", Status = "failed" });
                group = ((ChatActivityGroupView)((ChatDesignerHost)Call(window, "RenderEntry", entries[0])).View).section;
                steps = group.body.Controls.OfType<ChatActivityStepView>().ToArray();
                Assert.IsTrue(steps[1].section.Expanded); Assert.AreEqual("Changed result", steps[1].detail.content.Text);
                StringAssert.Contains(steps[1].state.Text, UiText.Get("Failed"));
                Call(window, "ReceiveAgentActivity", new CodexAgentActivity { Id = "c", Kind = "commandExecution", Detail = "\nnext", Append = true, Status = "inProgress" });
                Assert.AreEqual("dotnet test", entries[2].Activity.Title); Assert.AreEqual("passed\nnext", entries[2].Activity.Detail);
                Assert.AreEqual("completed", entries[2].Activity.Status, "A late output delta cannot reopen a completed action.");
                foreach (var state in new[] { "declined", "interrupted", "completed" })
                {
                    Call(window, "ReceiveAgentActivity", new CodexAgentActivity { Id = "c", Kind = "commandExecution", Detail = "result", Status = state });
                    Assert.IsNotNull(Call(window, "RenderActivityStep", entries[2]));
                }
                Call(window, "ReceiveAgentActivity", new CodexAgentActivity { Kind = "reasoning", Id = "empty", Detail = " " });
                Call(window, "ReceiveAgentActivity", new CodexAgentActivity()); Call(window, "ReceiveAgentActivity", (object)null);
                Assert.AreEqual(3, entries.Count);
                Call(window, "RefreshTranscriptWindow", 0); Assert.AreEqual(1, Get<ObservableCollection<object>>(window, "visibleEntries").Count);
                group = ((ChatActivityGroupView)((ChatDesignerHost)Call(window, "RenderEntry", entries[0])).View).section; group.Expanded = true;
                foreach (var step in group.body.Controls.OfType<ChatActivityStepView>()) step.section.Expanded = true;
                string preview = System.Environment.GetEnvironmentVariable("VBAi_ACTIVITY_PREVIEW");
                if (!string.IsNullOrEmpty(preview))
                {
                    group.Width = 500; group.PerformLayout();
                    using (var bitmap = new System.Drawing.Bitmap(group.Width, group.Height)) {
                        group.DrawToBitmap(bitmap, new System.Drawing.Rectangle(0,0,group.Width,group.Height));
                        System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(preview)); bitmap.Save(preview, System.Drawing.Imaging.ImageFormat.Png);
                    }
                }
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void ActivitiesGroupBetweenMessagesAndPreserveStreamsRawHistoryAndExpansion()
        {
            using (var window = Surfaces())
            {
                Call(window, "ClearTranscript");
                Set(window, "activeTurnId", "turn-1");
                Call(window, "AddTranscriptMessage", "Vous", "Question");
                Call(window, "ReceiveChatUpdate", "tool", "tool-a", "Lecture", false);
                Call(window, "ReceiveChatUpdate", "summary", "reason-a", "Analyse", false);
                Call(window, "ReceiveChatUpdate", "tool", "tool-b", "Écriture", true);
                var raw = Get<List<ChatEntry>>(window, "transcriptEntries");
                var visible = Get<ObservableCollection<object>>(window, "visibleEntries");
                Assert.AreEqual(4, raw.Count); Assert.AreEqual(2, visible.Count);
                var owner = raw[1]; Assert.AreSame(owner, visible[1]);
                var group = ((ChatActivityGroupView)((ChatDesignerHost)Call(window, "RenderEntry", owner)).View).section;
                Assert.IsFalse(group.Expanded); StringAssert.EndsWith(group.Title, " · 3");
                Assert.AreEqual(3, group.body.Controls.OfType<ChatTextContentView>().Count());
                group.Expanded = true;
                Call(window, "ReceiveChatUpdate", "tool", "tool-a", " terminée", false);
                TimerTick(Get<System.Windows.Threading.DispatcherTimer>(window, "streamRenderTimer"));
                Assert.AreEqual("Lecture terminée", Get<Dictionary<string, System.Windows.Forms.RichTextBox>>(window, "liveTexts")["tool-a"].Text);
                Call(window, "ReceiveChatUpdate", "summary", "reason-a", "Analyse terminée", true);
                group = ((ChatActivityGroupView)((ChatDesignerHost)Call(window, "RenderEntry", owner)).View).section;
                Assert.IsTrue(group.Expanded);
                Assert.AreEqual("Analyse terminée", Get<Dictionary<string, System.Windows.Forms.RichTextBox>>(window, "liveTexts")["reason-a"].Text);
                Call(window, "AddTranscriptMessage", "Assistant", "Réponse intermédiaire");
                Call(window, "ReceiveChatUpdate", "tool", "tool-c", "Suite", true);
                Assert.AreEqual(4, visible.Count); Assert.AreEqual(6, raw.Count);
                Set(window, "activeTurnId", "turn-2");
                Call(window, "ReceiveChatUpdate", "tool", "tool-d", "Tour suivant", true);
                Assert.AreEqual(5, visible.Count);
                Call(window, "RefreshTranscriptWindow", 0);
                Assert.AreEqual(5, visible.Count); Assert.AreEqual(7, raw.Count);
                Assert.IsTrue((((ChatActivityGroupView)((ChatDesignerHost)Call(window, "RenderEntry", owner)).View).section).Expanded);
                Call(window, "ReleaseActivityTexts", owner);
                Assert.IsFalse(Get<Dictionary<string, System.Windows.Forms.RichTextBox>>(window, "liveTexts").ContainsKey("tool-a"));
                Call(window, "ClearTranscript");
                Assert.AreEqual(0, Get<Dictionary<ChatEntry, List<ChatEntry>>>(window, "activityGroups").Count);
                Assert.AreEqual(0, Get<HashSet<ChatEntry>>(window, "expandedActivityGroups").Count);
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void ActivityPagingRetainsWholeGroupsAndSessionLoadingRebuildsTheProjection()
        {
            using (var window = Surfaces())
            {
                Call(window, "ClearTranscript"); Set(window, "loadingSession", true);
                var entries = new[] {
                    new ChatEntry { Speaker = "Vous", Text = "Question", TurnId = "a" },
                    new ChatEntry { Speaker = "Outil", Text = "Premier", TurnId = "a" },
                    new ChatEntry { Speaker = "Réflexion", Text = "Résumé", TurnId = "a" },
                    new ChatEntry { Speaker = "Outil", Text = "Dernier", TurnId = "a" },
                    new ChatEntry { Speaker = "Assistant", Text = "Réponse", TurnId = "a" }
                };
                foreach (var entry in entries) Call(window, "AddEntry", entry);
                Set(window, "loadingSession", false);
                Call(window, "RefreshTranscriptWindow", 3);
                Assert.AreEqual(1, Get<int>(window, "firstLoadedEntry"));
                var visible = Get<ObservableCollection<object>>(window, "visibleEntries");
                Assert.AreEqual(3, visible.Count); Assert.AreSame(entries[1], visible[1]);
                var item = new TranscriptItem { DataContext = entries[1] };
                Call(window, "RealizeEntry", item);
                Assert.IsInstanceOfType(item.Content, typeof(ChatDesignerHost));
                Assert.AreEqual(3, ((ChatActivityGroupView)((ChatDesignerHost)item.Content).View).section.body.Controls.OfType<ChatTextContentView>().Count());
                Call(window, "RefreshTranscriptWindow", 0);
                Assert.AreEqual(3, visible.Count); Assert.AreEqual(5, Get<List<ChatEntry>>(window, "transcriptEntries").Count);
                Call(window, "ReleaseActivityTexts", entries[4]);
            }
        }
    }
}

namespace VBAi.Tests.Unit
{
    using System.Collections.Generic;
    using System.Linq;
    using System.Windows.Controls;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class ChatWindowStateTests
    {
        [STATestMethod, TestCategory("Unit")]
        public void NativeActivityUpgradesLegacyStreamsAndRendersEveryTerminalAndRunningState()
        {
            using (var window = Surfaces())
            {
                Call(window, "ReceiveChatUpdate", "tool", "legacy", "previous text", false);
                var entry = Get<Dictionary<string, ChatEntry>>(window, "liveEntries")["legacy"];
                Assert.IsNull(entry.Activity);
                Call(window, "ReceiveAgentActivity", new CodexAgentActivity { Id = "legacy", Kind = "dynamicToolCall", Title = " ", Detail = "native detail", Append = true, Status = "inProgress" });
                Assert.IsNull(entry.Activity.Title);
                Assert.AreEqual("native detail", entry.Activity.Detail);
                Assert.AreEqual("inProgress", entry.Activity.Status);
                Assert.IsNull(entry.Activity.DurationMs);
                Call(window, "ReceiveAgentActivity", new CodexAgentActivity { Id = "legacy", Kind = "dynamicToolCall", Title = new string('t', 120), Detail = " continuation", Append = true, Status = "inProgress", DurationMs = 1234 });
                Assert.AreEqual("native detail continuation", entry.Activity.Detail);
                Assert.AreEqual(1234L, entry.Activity.DurationMs);
                var group = (ChatDesignerHost)Call(window, "RenderActivityGroup", entry, new List<ChatEntry> { entry });
                StringAssert.EndsWith(((ChatActivityGroupView)group.View).section.Title, "…");
                group.Dispose();
                foreach (var state in new[] { "inProgress", "failed", "declined", "completed", "interrupted", "cancelled" })
                {
                    entry.Activity.Status = state;
                    var step = (ChatDesignerHost)Call(window, "RenderActivityStep", entry);
                    var card = (ChatActivityStepView)step.View;
                    var label = card.state.Text;
                    StringAssert.StartsWith(label, UiText.Get(state == "inProgress" ? "In progress" : state == "failed" ? "Failed" : state == "declined" ? "Declined" : state == "completed" ? "Completed" : state == "interrupted" ? "Interrupted" : "Cancelled"));
                    StringAssert.Contains(label, " s");
                    Assert.AreEqual("native detail continuation", card.detail.content.Text);
                    step.Dispose();
                }
            }
        }
    }
}
namespace VBAi.Tests.Unit
{
    using System.Collections.Generic;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    public sealed partial class ChatWindowStateTests
    {
        [STATestMethod, TestCategory("Unit")]
        public void CollapsingNativeActivityGroupAndStepRemovesTheirPersistedExpansion()
        {
            using(var window=Surfaces()) {
                var entry=new ChatEntry { Speaker="Outil", Activity=new CodexAgentActivity { Kind="commandExecution",Title="command",Detail="details",Status="completed" } };
                using(var host=(ChatDesignerHost)Call(window,"RenderActivityGroup",entry,new List<ChatEntry>{entry})) {
                    var group=((ChatActivityGroupView)host.View).section;
                    var step=(ChatActivityStepView)group.body.Controls[0];
                    group.Expanded=true; step.section.Expanded=true;
                    Assert.IsTrue(Get<HashSet<ChatEntry>>(window,"expandedActivityGroups").Contains(entry));
                    Assert.IsTrue(Get<HashSet<ChatEntry>>(window,"expandedActivitySteps").Contains(entry));
                    group.Expanded=false; step.section.Expanded=false;
                    Assert.IsFalse(Get<HashSet<ChatEntry>>(window,"expandedActivityGroups").Contains(entry));
                    Assert.IsFalse(Get<HashSet<ChatEntry>>(window,"expandedActivitySteps").Contains(entry));
                }
                using(var recreated=(ChatDesignerHost)Call(window,"RenderActivityGroup",entry,new List<ChatEntry>{entry})) {
                    var group=((ChatActivityGroupView)recreated.View).section;
                    Assert.IsFalse(group.Expanded); Assert.IsFalse(((ChatActivityStepView)group.body.Controls[0]).section.Expanded);
                }
            }
        }
    }
}
