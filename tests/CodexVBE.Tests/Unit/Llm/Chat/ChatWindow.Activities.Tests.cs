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
