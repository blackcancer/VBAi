namespace CodexVBE.Tests.Unit
{
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Linq;
    using System.Windows;
    using System.Windows.Controls;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class ChatWindowStateTests
    {
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
                string preview = System.Environment.GetEnvironmentVariable("CODEXVBE_ACTIVITY_PREVIEW");
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

namespace CodexVBE.Tests.Unit
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
                foreach (var state in new[] { "inProgress", "failed", "declined", "completed", "interrupted" })
                {
                    entry.Activity.Status = state;
                    var step = (ChatDesignerHost)Call(window, "RenderActivityStep", entry);
                    var card = (ChatActivityStepView)step.View;
                    var label = card.state.Text;
                    StringAssert.StartsWith(label, UiText.Get(state == "inProgress" ? "In progress" : state == "failed" ? "Failed" : state == "declined" ? "Declined" : state == "completed" ? "Completed" : "Cancelled"));
                    StringAssert.Contains(label, " s");
                    Assert.AreEqual("native detail continuation", card.detail.content.Text);
                    step.Dispose();
                }
            }
        }
    }
}
namespace CodexVBE.Tests.Unit
{
    using System.Collections.Generic;
    using CodexVBE;
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
