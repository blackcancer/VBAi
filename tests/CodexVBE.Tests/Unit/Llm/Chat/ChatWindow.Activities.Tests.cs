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
                var group = (Expander)Call(window, "RenderEntry", entries[0]);
                var steps = ((StackPanel)group.Content).Children.OfType<Expander>().ToArray();
                Assert.AreEqual(3, steps.Length); Assert.IsFalse(steps[1].IsExpanded);
                steps[1].IsExpanded = true;
                Call(window, "ReceiveAgentActivity", new CodexAgentActivity { Id = "t", Kind = "dynamicToolCall", Title = "read_module", Detail = "Changed result", Status = "failed" });
                group = (Expander)Call(window, "RenderEntry", entries[0]);
                steps = ((StackPanel)group.Content).Children.OfType<Expander>().ToArray();
                Assert.IsTrue(steps[1].IsExpanded); Assert.AreEqual("Changed result", ((TextBox)steps[1].Content).Text);
                Assert.IsTrue(((DockPanel)steps[1].Header).Children.OfType<TextBlock>().Any(b => b.Text.Contains(UiText.Get("Failed"))));
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
                group = (Expander)Call(window, "RenderEntry", entries[0]); group.IsExpanded = true;
                foreach (var step in ((StackPanel)group.Content).Children.OfType<Expander>()) step.IsExpanded = true;
                string preview = System.Environment.GetEnvironmentVariable("CODEXVBE_ACTIVITY_PREVIEW");
                if (!string.IsNullOrEmpty(preview))
                {
                    group.Measure(new Size(500, double.PositiveInfinity)); group.Arrange(new Rect(0, 0, 500, group.DesiredSize.Height)); group.UpdateLayout();
                    var bitmap = new System.Windows.Media.Imaging.RenderTargetBitmap(500, (int)System.Math.Ceiling(group.ActualHeight), 96, 96, System.Windows.Media.PixelFormats.Pbgra32); bitmap.Render(group);
                    var png = new System.Windows.Media.Imaging.PngBitmapEncoder(); png.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(bitmap));
                    System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(preview)); using (var stream = System.IO.File.Create(preview)) png.Save(stream);
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
                var group = (Expander)Call(window, "RenderEntry", owner);
                Assert.IsFalse(group.IsExpanded); StringAssert.EndsWith((string)group.Header, " · 3");
                Assert.AreEqual(3, ((StackPanel)group.Content).Children.OfType<TextBox>().Count());
                group.IsExpanded = true;
                Call(window, "ReceiveChatUpdate", "tool", "tool-a", " terminée", false);
                Assert.AreEqual("Lecture terminée", Get<Dictionary<string, TextBox>>(window, "liveTexts")["tool-a"].Text);
                Call(window, "ReceiveChatUpdate", "summary", "reason-a", "Analyse terminée", true);
                group = (Expander)Call(window, "RenderEntry", owner);
                Assert.IsTrue(group.IsExpanded);
                Assert.AreEqual("Analyse terminée", Get<Dictionary<string, TextBox>>(window, "liveTexts")["reason-a"].Text);
                Call(window, "AddTranscriptMessage", "Assistant", "Réponse intermédiaire");
                Call(window, "ReceiveChatUpdate", "tool", "tool-c", "Suite", true);
                Assert.AreEqual(4, visible.Count); Assert.AreEqual(6, raw.Count);
                Set(window, "activeTurnId", "turn-2");
                Call(window, "ReceiveChatUpdate", "tool", "tool-d", "Tour suivant", true);
                Assert.AreEqual(5, visible.Count);
                Call(window, "RefreshTranscriptWindow", 0);
                Assert.AreEqual(5, visible.Count); Assert.AreEqual(7, raw.Count);
                Assert.IsTrue(((Expander)Call(window, "RenderEntry", owner)).IsExpanded);
                Call(window, "ReleaseActivityTexts", owner);
                Assert.IsFalse(Get<Dictionary<string, TextBox>>(window, "liveTexts").ContainsKey("tool-a"));
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
                Assert.IsInstanceOfType(item.Content, typeof(Expander));
                Assert.AreEqual(3, ((StackPanel)((Expander)item.Content).Content).Children.OfType<TextBox>().Count());
                Call(window, "RefreshTranscriptWindow", 0);
                Assert.AreEqual(3, visible.Count); Assert.AreEqual(5, Get<List<ChatEntry>>(window, "transcriptEntries").Count);
                Call(window, "ReleaseActivityTexts", entries[4]);
            }
        }
    }
}
