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

    /// <summary>Vérifie la mise à jour incrémentale et le rendu des entrées du transcript.</summary>
    public sealed partial class ChatWindowStateTests
    {
        /// <summary>Met à jour les flux du transcript en place et évite de dupliquer le message final.</summary>
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

        /// <summary>Complète en place les résumés et entrées d’outil transmis en flux.</summary>
        [TestMethod]
        [STATestMethod]
        public void StreamingSummaryAndToolEntriesCompleteInPlace()
        {
            using (var window = Surfaces())
            {
                Call(window, "ReceiveChatUpdate", "summary", "summary-1", "Réflexion ", false);
                Call(window, "ReceiveChatUpdate", "summary", "summary-1", "terminée", true);
                Call(window, "ReceiveChatUpdate", "tool", "tool-1", "Lecture", false);
                Call(window, "ReceiveChatUpdate", "tool", "tool-1", null, true);
                var entries = Get<List<ChatEntry>>(window, "transcriptEntries");
                Assert.AreEqual(2, entries.Count);
                Assert.AreEqual("Réflexion", entries[0].Speaker);
                Assert.AreEqual("terminée", entries[0].Text);
                Assert.AreEqual("Outil", entries[1].Speaker);
                Assert.AreEqual("Lecture", entries[1].Text);
                Assert.AreEqual(2, Get<HashSet<string>>(window, "completedStreams").Count);
            }
        }
    }
}
namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Documents;
    using System.Windows.Media;
    using CodexVBE;
    using CodexVBE.Tests.Infrastructure;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    /// <summary>Vérifie le rendu de chaque type d’entrée et la réutilisation des vues du transcript.</summary>
    public sealed partial class ChatWindowStateTests
    {
                /// <summary>Rend les différents types d’entrées, recycle leurs vues et préserve les actions de chaque carte.</summary>
[STATestMethod, TestCategory("Unit")]
        public void TranscriptRendersEveryEntryKindAndRecyclesViewsWithoutLosingState()
        {
            using (var runtime = new RuntimeScope())
            using (var window = new ChatWindow(runtime.Session))
            {
                CompleteOnSta((System.Threading.Tasks.Task)Call(window, "LoadModelsAsync"));
                var change = new CodeChange(@"C:\Temp\P.xlsm", "M", "old", "oldsha", "new", "newsha", 1) { TurnId = "turn" }; Get<List<CodeChange>>(window, "codeChanges").Add(change);
                var attachment = new ChatAttachment { Label = "Selection", Text = "code", Project = "P", Module = "M", Sha256 = "sha", StartLine = 1 };
                foreach (var speaker in new[] { "Vous", "Assistant", "Erreur", "Réflexion", "Outil", "Intervention", "Vérification" })
                    foreach (var streamed in new[] { false, true })
                    {
                        var entry = new ChatEntry { Speaker = speaker, Text = "message", TurnId = "turn", StreamId = streamed ? speaker : null, AttachedMemory = "note", Attachments = new[] { attachment, new ChatAttachment { Label = "plain", Text = "plain" } }, References = new[] { new VbeChatReference { Project = "P" } } };
                        Call(window, "AddEntry", entry); var view = (FrameworkElement)Call(window, "RenderEntry", entry); Assert.IsNotNull(view);
                        foreach (var copy in Descendants(view).OfType<Button>().Where(b => Convert.ToString(b.Content) == UiText.Get("Copy"))) { string copied = null; ChatWindow.WriteClipboard = t => copied = t; WpfClick(copy); Assert.AreEqual("message", copied); ChatWindow.WriteClipboard = t => { throw new InvalidOperationException("clipboard busy"); }; WpfClick(copy); StringAssert.Contains(Get<System.Windows.Forms.Label>(window, "status").Text, "clipboard busy"); }
                        if (speaker == "Vérification") { var fix = Descendants(view).OfType<Button>().Single(b => Convert.ToString(b.Content) == UiText.Get("Prepare a fix")); Set(window, "busy", true); WpfClick(fix); Set(window, "busy", false); WpfClick(fix); StringAssert.StartsWith(Get<TextBox>(window, "prompt").Text, "/corriger "); }
                        foreach (var link in Descendants(view).OfType<Button>().Where(b => Convert.ToString(b.Content) == "#P" || Convert.ToString(b.Content) == UiText.Get("Open in the VBE"))) WpfClick(link);
                    }
                var simple = new ChatEntry { Speaker = "Assistant", Text = null }; var container = new TranscriptItem { DataContext = simple }; container.SetValue(TranscriptItem.RenderProperty, new Action<TranscriptItem>(i => Call(window, "RealizeEntry", i))); container.SetValue(TranscriptItem.ReleaseProperty, new Action<TranscriptItem>(i => Call(window, "ReleaseEntry", i))); container.RaiseEvent(new RoutedEventArgs(FrameworkElement.LoadedEvent)); Assert.IsNotNull(container.Content); container.RaiseEvent(new RoutedEventArgs(FrameworkElement.UnloadedEvent)); Assert.AreEqual(0, Get<Dictionary<ChatEntry, FrameworkElement>>(window, "entryViews").Count);
                Call(window, "ReleaseEntry", new TranscriptItem()); Call(window, "RealizeEntry", new TranscriptItem { DataContext = new TextBlock { Text = "welcome" } }); Call(window, "RealizeEntry", new TranscriptItem { DataContext = new object() });
                var card = (FrameworkElement)Call(window, "RenderEntry", new ChatEntry { Change = change }); var restore = Get<Dictionary<CodeChange, Button>>(window, "rollbackButtons")[change]; Set(window, "busy", true); WpfClick(restore); Set(window, "busy", false); WpfClick(restore); Assert.IsTrue(change.Restored);
                foreach (var state in new[] { 0, 1, 2 }) { change.Restored = state == 2; change.RestoredHunks.Clear(); if (state == 1) change.RestoredHunks.Add(0); Call(window, "RefreshCodeChangeCards"); Assert.AreEqual(!change.Restored, restore.IsEnabled); }
                foreach (var button in Descendants(card).OfType<Button>().Where(b => Convert.ToString(b.Content) == UiText.Get("Undo a block…") || Convert.ToString(b.Content) == UiText.Get("Undo turn"))) WpfClick(button);
                Call(window, "RenderChange", change); Call(window, "AddCodeChangeCard", change); var already = Get<Dictionary<CodeChange, Button>>(window, "rollbackButtons")[change]; WpfClick(already); var host = runtime.Host; runtime.Host = r => Response.Failure("scope unavailable"); WpfClick(already); StringAssert.Contains(Get<System.Windows.Forms.Label>(window, "status").Text, "scope unavailable"); runtime.Host = host;
                var emptyTurn = (FrameworkElement)Call(window, "RenderEntry", new ChatEntry { Speaker = "Intervention", TurnId = "absent", Text = "empty" }); Assert.IsNotNull(emptyTurn);
                Get<HashSet<string>>(window, "completedStreams").Add("done"); Assert.IsNotNull(Call(window, "RenderEntry", new ChatEntry { Speaker = "Assistant", Text = "complete", StreamId = "done", References = new VbeChatReference[0], Attachments = new ChatAttachment[0] }));
                LocalizationScope.Set("ar-SA"); var code = (TextBox)Call(window, "SelectableText", null, true); Assert.AreEqual(FlowDirection.LeftToRight, code.FlowDirection); var plain = (TextBox)Call(window, "SelectableText", "rtl", false); Assert.AreEqual(FlowDirection.RightToLeft, plain.FlowDirection);
            }
        }
        /// <summary>Vérifie la pagination, les mises à jour en flux et la conservation de l’historique lors des changements de thème.</summary>
[STATestMethod, TestCategory("Unit")]
        public void TranscriptPagingStreamingWelcomeAndThemeChangesPreserveVisibleHistory()
        {
            using (var runtime = new RuntimeScope())
            using (var window = new ChatWindow(runtime.Session))
            {
                Call(window, "ClearTranscript"); Call(window, "ShowWelcome"); var visible = Get<System.Collections.ObjectModel.ObservableCollection<object>>(window, "visibleEntries"); var welcome = (StackPanel)visible.Last(); foreach (var b in welcome.Children.OfType<Button>()) { WpfClick(b); Assert.IsTrue(Get<TextBox>(window, "prompt").Text.EndsWith(" ")); }
                for (int i = 0; i < 170; i++) Call(window, "AddEntry", new ChatEntry { Speaker = "Assistant", Text = "line " + i }); Call(window, "RefreshTranscriptWindow", 90); Assert.AreEqual(81, visible.Count);
                var marker = new TranscriptItem { DataContext = Get<object>(window, "earlierEntries") }; Call(window, "RealizeEntry", marker); WpfClick((Button)marker.Content); Assert.AreEqual(10, Get<int>(window, "firstLoadedEntry"));
                Set(window, "followConversation", false); Call(window, "FollowLatest"); Set(window, "followConversation", true); Call(window, "FollowLatest"); window.Show(); System.Windows.Forms.Application.DoEvents();
                UiInvoke.Call(typeof(UiTheme), "PreferencesChanged", null, null, new Microsoft.Win32.UserPreferenceChangedEventArgs(Microsoft.Win32.UserPreferenceCategory.Color)); System.Windows.Forms.Application.DoEvents();
                foreach (var kind in new[] { "summary", "tool", "final" }) { Call(window, "ReceiveChatUpdate", kind, kind, "delta", false); var entry = Get<Dictionary<string, ChatEntry>>(window, "liveEntries")[kind]; Call(window, "RenderEntry", entry); Call(window, "ReceiveChatUpdate", kind, kind, null, false); Call(window, "ReceiveChatUpdate", kind, kind, "", true); Assert.AreEqual("delta", entry.Text); }
                Call(window, "RefreshTranscriptWindow", Get<List<ChatEntry>>(window, "transcriptEntries").Count); Call(window, "ReceiveChatUpdate", "final", "final", "replacement", true); Assert.AreEqual("replacement", Get<Dictionary<string, ChatEntry>>(window, "liveEntries")["final"].Text);
                Call(window, "ClearTranscript"); Call(window, "FollowLatest"); Call(window, "ShowWelcome"); window.Close(); Call(window, "ReceiveChatUpdate", "final", "disposed", "ignored", false); Call(window, "Append", "Assistant", "ignored"); Call(window, "SetStatus", "ignored");
            }
        }
    }
}
namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Media;
    using System.Threading.Tasks;
    using CodexVBE;
    using CodexVBE.Tests.Infrastructure;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    /// <summary>Vérifie le routage des vues recyclées et les actions de suivi du transcript.</summary>
    public sealed partial class ChatWindowStateTests
    {
        /// <summary>Exerce le chargement et le déchargement des vues, les événements de défilement et le suivi différé.</summary>
[STATestMethod, TestCategory("Unit")]
        public void TranscriptReleaseRoutingAndPendingFollowActionsHandleEveryLifecycle()
        {
            using (var runtime = new RuntimeScope())
            {
                LocalizationScope.Set("ar-SA");
                using (var window = new ChatWindow(runtime.Session))
                {
                    var list = Get<ListBox>(window, "conversationItems"); Assert.AreEqual(FlowDirection.RightToLeft, list.FlowDirection); window.Show(); list.ApplyTemplate(); System.Windows.Forms.Application.DoEvents(); var scroll = Visual<ScrollViewer>(list); Assert.IsNotNull(scroll);
                    RaiseScroll(window, list, 1, 0); foreach (var vertical in new[] { 0d, 1d }) foreach (var extent in new[] { 0d, 1d }) RaiseScroll(window, scroll, vertical, extent);
                    Call(window, "CopyText", (object)null); Call(window, "ClearTranscript");
                    foreach (var entry in new[] { new ChatEntry { Speaker = "Assistant", Text = "live", StreamId = "release" }, new ChatEntry { Change = new CodeChange(@"C:\Temp\P.xlsm", "M", "old", "before", "new", "after", 1) } })
                    {
                        Call(window, "AddEntry", entry);
                        var first = new TranscriptItem { DataContext = entry }; first.SetValue(TranscriptItem.RenderProperty, new Action<TranscriptItem>(i => Call(window, "RealizeEntry", i))); first.SetValue(TranscriptItem.ReleaseProperty, new Action<TranscriptItem>(i => Call(window, "ReleaseEntry", i))); first.RaiseEvent(new RoutedEventArgs(FrameworkElement.LoadedEvent));
                        var second = new TranscriptItem { DataContext = entry }; second.SetValue(TranscriptItem.RenderProperty, new Action<TranscriptItem>(i => Call(window, "RealizeEntry", i))); second.SetValue(TranscriptItem.ReleaseProperty, new Action<TranscriptItem>(i => Call(window, "ReleaseEntry", i))); second.RaiseEvent(new RoutedEventArgs(FrameworkElement.LoadedEvent)); first.RaiseEvent(new RoutedEventArgs(FrameworkElement.UnloadedEvent)); Assert.IsTrue(Get<Dictionary<ChatEntry, FrameworkElement>>(window, "entryViews").ContainsKey(entry)); second.RaiseEvent(new RoutedEventArgs(FrameworkElement.UnloadedEvent)); Assert.IsFalse(Get<Dictionary<ChatEntry, FrameworkElement>>(window, "entryViews").ContainsKey(entry)); Call(window, "ReleaseEntry", second);
                    }
                    Set(window, "followConversation", true); Call(window, "FollowLatest"); Set(window, "followConversation", false); System.Windows.Forms.Application.DoEvents(); Set(window, "followConversation", true); Call(window, "FollowLatest"); Call(window, "ClearTranscript"); System.Windows.Forms.Application.DoEvents();
                    var marker = new TranscriptItem { DataContext = Get<object>(window, "earlierEntries") }; Call(window, "RealizeEntry", marker); WpfClick((Button)marker.Content);
                    Call(window, "AddEntry", new ChatEntry { Speaker = "Assistant", Text = "nonempty" }); Call(window, "ShowWelcome");
                }
                using (var design = new ChatWindow())
                {
                    Call(design, "AddEntry", new ChatEntry { Speaker = "Assistant", Text = "ignored" }); Call(design, "ReceiveChatUpdate", "final", "ignored", "text", false); Call(design, "ShowWelcome"); Call(design, "FollowLatest");
                    Call(design, "InitializeComposer", runtime.Session); Call(design, "InitializeTranscript"); var change = new CodeChange { Project = "P", Module = "M", Before = "old", After = "new" }; var card = (FrameworkElement)Call(design, "RenderChange", change); WpfClick(Get<Dictionary<CodeChange, Button>>(design, "rollbackButtons")[change]);
                    UiInvoke.Call(typeof(UiTheme), "PreferencesChanged", null, null, new Microsoft.Win32.UserPreferenceChangedEventArgs(Microsoft.Win32.UserPreferenceCategory.Color));
                    var queuedTheme = (Action)typeof(UiTheme).GetField("Changed", System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic).GetValue(null);
                    var blocks = Descendants(card).OfType<Button>().Single(b => Convert.ToString(b.Content) == UiText.Get("Undo a block…")); WpfClick(blocks); var menu = blocks.ContextMenu; foreach (var item in menu.Items.OfType<MenuItem>()) item.RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent)); menu.IsOpen = false; design.Dispose(); queuedTheme();
                }
            }
        }
    }
}
