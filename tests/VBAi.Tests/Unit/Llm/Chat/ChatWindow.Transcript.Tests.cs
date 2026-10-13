namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System.Collections.Generic;
    using VBAi;

    /// <summary>Vérifie la mise à jour incrémentale et le rendu des entrées du transcript.</summary>
    public sealed partial class ChatWindowStateTests
    {
        /// <summary>The real message renderer must not create a link from a future prompt or conflicting attachments.</summary>
        [STATestMethod]
        [DataRow("future-only")]
        [DataRow("explicit-empty")]
        [DataRow("conflicting")]
        [DataRow("corrupt")]
        public void RenderedHistoricalMessageDoesNotBorrowLaterOrAmbiguousReferences(string scenario)
        {
            using (var window = Surfaces())
            {
                var first = new VbeChatReference { Project = "P", Module = "M", Name = "Calculate", Sha256 = "old" };
                var second = new VbeChatReference { Project = "P", Module = "M", Name = "Calculate", Sha256 = "new" };
                var entry = new ChatEntry { Speaker = "Assistant", TurnId = "old", Text = first.Token };
                if (scenario == "explicit-empty") entry.References = new VbeChatReference[0];
                if (scenario == "conflicting") entry.References = new[] { first, second };
                if (scenario == "corrupt") entry.References = new[] { null, new VbeChatReference() };
                var entries = Get<List<ChatEntry>>(window, "transcriptEntries");
                entries.Add(entry);
                entries.Add(new ChatEntry { Speaker = "Vous", TurnId = "new", References = new[] { second } });
                using (var host = (ChatDesignerHost)Call(window, "RenderEntry", entry))
                {
                    var markdown = ((ChatMessageView)host.View).message.MarkdownView;
                    Assert.AreEqual(entry.Text, markdown.SourceText);
                    Assert.AreEqual(1, markdown.ContentPanel.Controls.Count);
                    var block = (Blackcancer.ChatUI.Views.ChatTextView)markdown.ContentPanel.Controls[0];
                    StringAssert.Contains(block.Content.Text, first.Token);
                    var actions = (System.Collections.IList)typeof(Blackcancer.ChatUI.Views.ChatTextView)
                        .GetField("actions", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).GetValue(block);
                    Assert.AreEqual(0, actions.Count, scenario);
                }
            }
        }

        /// <summary>Stored messages own their attachment snapshots even when the producer refreshes a module hash.</summary>
        [STATestMethod, DataRow(false), DataRow(true)]
        public void CapturedTranscriptReferencesStayDetachedFromTheirProducer(bool procedure)
        {
            using (var window = Surfaces())
            {
                var reference = new VbeChatReference { Project = "P", Module = "M", Name = procedure ? "Calculate" : null, Sha256 = "old", StartLine = 2, EndLine = 6 };
                string token = reference.Token;
                var references = new[] { reference };
                var entry = new ChatEntry { Speaker = "Vous", Text = token, References = references };
                Call(window, "AddEntry", entry);
                var stored = Get<List<ChatEntry>>(window, "transcriptEntries")[0].References[0];
                Assert.AreNotSame(reference, stored); Assert.AreNotSame(references, entry.References);
                reference.Project = "Changed"; reference.Sha256 = "new"; reference.StartLine = 90;
                references[0] = new VbeChatReference { Project = "Replacement" };
                Assert.AreEqual(token, stored.Token); Assert.AreEqual("old", stored.Sha256); Assert.AreEqual(2, stored.StartLine);
            }
        }

        /// <summary>Met à jour les flux du transcript en place et évite de dupliquer le message final.</summary>
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
namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Windows;
    using System.Windows.Controls;
    using VBAi;
    using VBAi.Tests.Infrastructure;
    /// <summary>Vérifie le rendu de chaque type d’entrée et la réutilisation des vues du transcript.</summary>
    public sealed partial class ChatWindowStateTests
    {
        /// <summary>Rend les différents types d’entrées, recycle leurs vues et préserve les actions de chaque carte.</summary>
        [STATestMethod, TestCategory("Unit")]
        public void TranscriptRendersEveryEntryKindAndRecyclesViewsWithoutLosingState()
        {
            using (var runtime = new RuntimeScope())
            using (var window = LoadedWindow(runtime.Session))
            {
                CompleteOnSta((System.Threading.Tasks.Task)Call(window, "LoadModelsAsync"));
                var change = new CodeChange(@"C:\Temp\P.xlsm", "M", "old", "oldsha", "new", "newsha", 1) { TurnId = "turn" }; Get<List<CodeChange>>(window, "codeChanges").Add(change);
                var attachment = new ChatAttachment { Label = "Selection", Text = "code", Project = "P", Module = "M", Sha256 = "sha", StartLine = 1 };
                foreach (var speaker in new[] { "Vous", "Assistant", "Erreur", "Réflexion", "Outil", "Intervention", "Vérification" })
                    foreach (var streamed in new[] { false, true })
                    {
                        var entry = new ChatEntry { Speaker = speaker, Text = "message", TurnId = "turn", StreamId = streamed ? speaker : null, AttachedMemory = "note", Attachments = new[] { attachment, new ChatAttachment { Label = "plain", Text = "plain" } }, References = new[] { new VbeChatReference { Project = "P" } } };
                        Call(window, "AddEntry", entry); var view = (FrameworkElement)Call(window, "RenderEntry", entry); Assert.IsNotNull(view);
                        foreach (var copy in NativeDescendants(view).OfType<System.Windows.Forms.Button>().Where(b => b.Text == UiText.Get("Copy"))) { string copied = null; ChatWindow.WriteClipboard = t => copied = t; WpfClick(copy); Assert.AreEqual("message", copied); ChatWindow.WriteClipboard = t => { throw new InvalidOperationException("clipboard busy"); }; WpfClick(copy); StringAssert.Contains(Get<System.Windows.Forms.Label>(window, "status").Text, "clipboard busy"); }
                        if (speaker == "Vérification") { var fix = NativeDescendants(view).OfType<System.Windows.Forms.Button>().Single(b => b.Text == UiText.Get("Prepare a fix")); Set(window, "busy", true); WpfClick(fix); Set(window, "busy", false); WpfClick(fix); StringAssert.StartsWith(Get<TextBox>(window, "prompt").Text, "/corriger "); }
                        foreach (var link in NativeDescendants(view).OfType<System.Windows.Forms.Button>().Where(b => b.Text == "#P" || b.Text == UiText.Get("Open in the VBE"))) WpfClick(link);
                    }
                var simple = new ChatEntry { Speaker = "Assistant", Text = null }; var container = new TranscriptItem { DataContext = simple }; container.SetValue(TranscriptItem.RenderProperty, new Action<TranscriptItem>(i => Call(window, "RealizeEntry", i))); container.SetValue(TranscriptItem.ReleaseProperty, new Action<TranscriptItem>(i => Call(window, "ReleaseEntry", i))); container.RaiseEvent(new RoutedEventArgs(FrameworkElement.LoadedEvent)); Assert.IsNotNull(container.Content); container.RaiseEvent(new RoutedEventArgs(FrameworkElement.UnloadedEvent)); Assert.AreEqual(0, Get<Dictionary<ChatEntry, FrameworkElement>>(window, "entryViews").Count);
                Call(window, "ReleaseEntry", new TranscriptItem()); Call(window, "RealizeEntry", new TranscriptItem { DataContext = new TextBlock { Text = "welcome" } }); Call(window, "RealizeEntry", new TranscriptItem { DataContext = new object() });
                var card = (FrameworkElement)Call(window, "RenderEntry", new ChatEntry { Change = change }); var restore = Get<Dictionary<CodeChange, System.Windows.Forms.Button>>(window, "rollbackButtons")[change]; Set(window, "busy", true); WpfClick(restore); Set(window, "busy", false); WpfClick(restore); Assert.IsTrue(change.Restored);
                foreach (var state in new[] { 0, 1, 2 }) { change.Restored = state == 2; change.RestoredHunks.Clear(); if (state == 1) change.RestoredHunks.Add(0); Call(window, "RefreshCodeChangeCards"); Assert.AreEqual(!change.Restored, restore.Enabled); }
                foreach (var button in NativeDescendants(card).OfType<System.Windows.Forms.Button>().Where(b => b.Text == UiText.Get("Undo a block…") || b.Text == UiText.Get("Undo turn"))) WpfClick(button);
                Call(window, "RenderChange", change); Call(window, "AddCodeChangeCard", change); var already = Get<Dictionary<CodeChange, System.Windows.Forms.Button>>(window, "rollbackButtons")[change]; WpfClick(already); var host = runtime.Host; runtime.Host = r => Response.Failure("scope unavailable"); WpfClick(already); StringAssert.Contains(Get<System.Windows.Forms.Label>(window, "status").Text, "scope unavailable"); runtime.Host = host;
                var emptyTurn = (FrameworkElement)Call(window, "RenderEntry", new ChatEntry { Speaker = "Intervention", TurnId = "absent", Text = "empty" }); Assert.IsNotNull(emptyTurn);
                Get<HashSet<string>>(window, "completedStreams").Add("done"); Assert.IsNotNull(Call(window, "RenderEntry", new ChatEntry { Speaker = "Assistant", Text = "complete", StreamId = "done", References = new VbeChatReference[0], Attachments = new ChatAttachment[0] }));
                LocalizationScope.Set("ar-SA"); var code = (ChatTextContentView)Call(window, "SelectableText", null, true); Assert.AreEqual(System.Windows.Forms.RightToLeft.No, code.content.RightToLeft); var plain = (ChatTextContentView)Call(window, "SelectableText", "rtl", false); Assert.AreEqual(System.Windows.Forms.RightToLeft.Yes, plain.content.RightToLeft);
            }
        }
        /// <summary>Vérifie la pagination, les mises à jour en flux et la conservation de l’historique lors des changements de thème.</summary>
        [STATestMethod, TestCategory("Unit")]
        public void TranscriptPagingStreamingWelcomeAndThemeChangesPreserveVisibleHistory()
        {
            using (var runtime = new RuntimeScope())
            using (var window = LoadedWindow(runtime.Session))
            {
                Call(window, "ClearTranscript"); Call(window, "ShowWelcome"); var visible = Get<System.Collections.ObjectModel.ObservableCollection<object>>(window, "visibleEntries"); var welcome = (ChatDesignerHost)visible.Last(); foreach (var b in NativeDescendants(welcome).OfType<System.Windows.Forms.Button>()) { WpfClick(b); Assert.IsTrue(Get<TextBox>(window, "prompt").Text.EndsWith(" ")); }
                for (int i = 0; i < 170; i++) Call(window, "AddEntry", new ChatEntry { Speaker = "Assistant", Text = "line " + i }); Call(window, "RefreshTranscriptWindow", 90); Assert.AreEqual(81, visible.Count);
                var marker = new TranscriptItem { DataContext = Get<object>(window, "earlierEntries") }; Call(window, "RealizeEntry", marker); WpfClick(((ChatLinkView)((ChatDesignerHost)marker.Content).View).link); Assert.AreEqual(10, Get<int>(window, "firstLoadedEntry"));
                Set(window, "followConversation", false); Call(window, "FollowLatest"); Set(window, "followConversation", true); Call(window, "FollowLatest"); window.Show(); System.Windows.Forms.Application.DoEvents();
                UiInvoke.Call(typeof(UiTheme), "PreferencesChanged", null, null, new Microsoft.Win32.UserPreferenceChangedEventArgs(Microsoft.Win32.UserPreferenceCategory.Color)); System.Windows.Forms.Application.DoEvents();
                foreach (var kind in new[] { "summary", "tool", "final" }) { Call(window, "ReceiveChatUpdate", kind, kind, "delta", false); var entry = Get<Dictionary<string, ChatEntry>>(window, "liveEntries")[kind]; Call(window, "RenderEntry", entry); Call(window, "ReceiveChatUpdate", kind, kind, null, false); Call(window, "ReceiveChatUpdate", kind, kind, "", true); Assert.AreEqual("delta", entry.Text); }
                Call(window, "RefreshTranscriptWindow", Get<List<ChatEntry>>(window, "transcriptEntries").Count); Call(window, "ReceiveChatUpdate", "final", "final", "replacement", true); Assert.AreEqual("replacement", Get<Dictionary<string, ChatEntry>>(window, "liveEntries")["final"].Text);
                Call(window, "ClearTranscript"); Call(window, "FollowLatest"); Call(window, "ShowWelcome"); window.Close(); Call(window, "ReceiveChatUpdate", "final", "disposed", "ignored", false); Call(window, "Append", "Assistant", "ignored"); Call(window, "SetStatus", "ignored");
            }
        }
    }
}
namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Windows;
    using System.Windows.Controls;
    using VBAi;
    using VBAi.Tests.Infrastructure;
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
                using (var window = LoadedWindow(runtime.Session))
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
                    var marker = new TranscriptItem { DataContext = Get<object>(window, "earlierEntries") }; Call(window, "RealizeEntry", marker); WpfClick(((ChatLinkView)((ChatDesignerHost)marker.Content).View).link);
                    Call(window, "AddEntry", new ChatEntry { Speaker = "Assistant", Text = "nonempty" }); Call(window, "ShowWelcome");
                }
                using (var design = new ChatWindow())
                {
                    Call(design, "AddEntry", new ChatEntry { Speaker = "Assistant", Text = "ignored" }); Call(design, "ReceiveChatUpdate", "final", "ignored", "text", false); Call(design, "ShowWelcome"); Call(design, "FollowLatest");
                    Call(design, "InitializeComposer", runtime.Session); Call(design, "InitializeTranscript"); var change = new CodeChange { Project = "P", Module = "M", Before = "old", After = "new" }; var card = (FrameworkElement)Call(design, "RenderChange", change); WpfClick(Get<Dictionary<CodeChange, System.Windows.Forms.Button>>(design, "rollbackButtons")[change]);
                    UiInvoke.Call(typeof(UiTheme), "PreferencesChanged", null, null, new Microsoft.Win32.UserPreferenceChangedEventArgs(Microsoft.Win32.UserPreferenceCategory.Color));
                    var queuedTheme = (Action)typeof(UiTheme).GetField("Changed", System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic).GetValue(null);
                    var blocks = NativeDescendants(card).OfType<System.Windows.Forms.Button>().Single(b => b.Text == UiText.Get("Undo a block…")); WpfClick(blocks); var menu = ((ChatChangeCardView)((ChatDesignerHost)card).View).blockMenu; foreach (var item in menu.Items.OfType<System.Windows.Forms.ToolStripMenuItem>()) item.PerformClick(); menu.Close(); design.Dispose(); queuedTheme();
                }
            }
        }
    }
}

namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System.Collections.Generic;
    using System.Reflection;
    using System.Windows;
    using System.Windows.Controls;
    using VBAi;

    [TestClass, TestCategory("Unit")]
    public sealed class ChatTranscriptNativeViewsTests
    {
        [STATestMethod]
        public void ReplacingTranscriptWindowDisposesItsHostedDesignerDiff()
        {
            using (var window = new ChatWindow())
            {
                var diff = new ChatDiffView("old", "new");
                var child = diff.Child;
                var body = new StackPanel(); body.Children.Add(diff);
                var views = (Dictionary<ChatEntry, FrameworkElement>)typeof(ChatWindow).GetField("entryViews", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(window);
                views.Add(new ChatEntry { Speaker = "Agent" }, body);
                typeof(ChatWindow).GetMethod("RefreshTranscriptWindow", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(window, new object[] { 0 });
                Assert.IsTrue(child.IsDisposed);
                Assert.AreEqual(0, views.Count);
            }
        }
    }
}

namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System.Collections.Generic;
    using System.Windows.Controls;
    using VBAi;
    public sealed partial class ChatWindowStateTests
    {
        [STATestMethod, TestCategory("Unit")]
        public void RecursiveViewDisposalReleasesBothNativeHostsAndIgnoresOrdinaryLeaves()
        {
            using (var window = Surfaces())
            {
                var card = new ChatMessageView(); var native = new ChatDesignerHost(card);
                var diff = new ChatDiffView("before", "after"); var diffChild = diff.Child;
                var nested = new StackPanel(); nested.Children.Add(native); nested.Children.Add(diff); nested.Children.Add(new TextBlock { Text = "ordinary" });
                var body = new StackPanel(); body.Children.Add(nested);
                Call(window, "DisposeEntryView", body);
                Assert.IsTrue(card.IsDisposed); Assert.IsTrue(card.message.content.IsDisposed); Assert.IsTrue(card.message.MarkdownView.IsDisposed); Assert.IsTrue(diffChild.IsDisposed);
                Assert.AreEqual(1, body.Children.Count); Assert.AreEqual(3, nested.Children.Count);
            }
        }
        [STATestMethod, TestCategory("Unit")]
        public void InterventionUndoCallbackHandlesEmptyTargetsAndRetainsUnrestoredTargetsWhileBusy()
        {
            using (var window = Surfaces())
            {
                var changes = Get<List<CodeChange>>(window, "codeChanges");
                var change = new CodeChange("P", "M", "old", "oldsha", "new", "newsha", 1) { TurnId = "assigned" }; changes.Add(change);
                Set(window, "busy", true);
                foreach (var turn in new[] { "unrelated", "assigned" })
                    using (var host = (ChatDesignerHost)Call(window, "RenderEntry", new ChatEntry { Speaker = "Intervention", TurnId = turn, Text = "result" }))
                    {
                        var card = (ChatMessageView)host.View; Assert.AreEqual(turn == "assigned" ? 1 : 0, card.targets.Controls.Count);
                        Click(card.undoTurn); Assert.IsFalse(change.Restored); Assert.AreEqual(1, changes.Count);
                    }
            }
        }
    }
}

namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Windows;
    using System.Windows.Controls;
    using VBAi.Tests.Infrastructure;

    public sealed partial class ChatWindowStateTests
    {
        /// <summary>The shared fixed shell preserves VBE slots, menu actions and the discrete docking backend.</summary>
        [STATestMethod]
        public void SharedWorkspaceKeepsVbeActionsAndActualNamedPin()
        {
            using (var host = new HistoryTestHost())
            using (var window = Surfaces())
            {
                window.TopLevel = false; window.Dock = System.Windows.Forms.DockStyle.Fill;
                host.Controls.Add(window); window.Show(); host.Show(); System.Windows.Forms.Application.DoEvents();
                var workspace = Get<Blackcancer.ChatUI.Views.ChatWorkspaceView>(window, "workspace");
                Assert.AreSame(workspace.sendButton, Get<System.Windows.Forms.Button>(window, "send"));
                Assert.AreSame(workspace.inputTextBox, Get<Blackcancer.ChatUI.Views.ChatInputView>(window, "promptHost"));
                Assert.IsFalse(workspace.inputTextBox.UseNativeEditor, "VBAi keeps the existing WPF composer and its reference bindings.");
                Assert.AreSame(workspace.TranscriptHost, Get<System.Windows.Forms.Panel>(window, "transcriptPanel"));
                Assert.AreSame(workspace.PendingMessagesHost, Get<System.Windows.Forms.FlowLayoutPanel>(window, "pendingMessagesPanel"));
                Assert.AreSame(workspace.AdditionalContextHost, Get<System.Windows.Forms.Panel>(window, "contextPanel"));
                foreach (string field in new[] { "methods", "compile", "changes" })
                    Assert.AreSame(workspace.HeaderActionsHost, Get<System.Windows.Forms.Button>(window, field).Parent, field);
                foreach (string field in new[] { "configure", "projectAccess", "resumeTurn", "refreshModels", "verifyChanges", "docking", "github", "about" })
                    Assert.IsTrue(workspace.optionsMenu.Items.Contains(Get<System.Windows.Forms.ToolStripMenuItem>(window, field)), field);
                Assert.AreEqual("pinButton", workspace.windowPinButton.Name);
                Assert.AreEqual(new System.Drawing.Size(32, 30), workspace.windowPinButton.Size);
                Assert.IsTrue(workspace.windowPinButton.IconOnly);
                Assert.IsTrue(workspace.windowPinButton.Visible);
                Assert.AreSame(workspace.windowPinButton, workspace.Controls.Find("pinButton", true).Single());
                int requests = 0; window.DockRequested += () => requests++;
                workspace.windowPinButton.PerformClick(); Assert.AreEqual(1, requests);
                Set(window, "busy", true);
                workspace.windowPinButton.PerformClick(); Assert.AreEqual(1, requests, "The existing busy docking guard must be preserved.");
                Set(window, "busy", false);
            }
        }

        /// <summary>Assistant streaming uses the visible renderer, preserves user code expansion and releases virtualized views.</summary>
        [STATestMethod]
        public void AssistantStreamingKeepsSharedMarkdownBlocksAndReleasesTheirViews()
        {
            using (var window = Surfaces())
            {
                const string streamId = "shared-markdown-stream";
                string code = string.Join("\n", Enumerable.Range(0, 13).Select(i => "' line " + i));
                string source = "**Answer**\n\n```vba\n" + code + "\n```\n\n| Item | Value |\n| --- | --- |\n| A | 1 |";
                Call(window, "ReceiveChatUpdate", "final", streamId, source, false);
                var entry = Get<List<ChatEntry>>(window, "transcriptEntries").Single();
                var item = new TranscriptItem { DataContext = entry };
                item.SetValue(TranscriptItem.RenderProperty, new Action<TranscriptItem>(i => Call(window, "RealizeEntry", i)));
                item.SetValue(TranscriptItem.ReleaseProperty, new Action<TranscriptItem>(i => Call(window, "ReleaseEntry", i)));
                item.RaiseEvent(new RoutedEventArgs(FrameworkElement.LoadedEvent));
                using (var host = (ChatDesignerHost)item.Content)
                {
                    var card = (ChatMessageView)host.View;
                    var markdown = card.message.MarkdownView;
                    Assert.AreSame(markdown, Get<Dictionary<string, Blackcancer.ChatUI.Views.ChatMarkdownView>>(window, "liveMarkdownViews")[streamId]);
                    Assert.IsFalse(Get<Dictionary<string, System.Windows.Forms.RichTextBox>>(window, "liveTexts").ContainsKey(streamId));
                    Call(window, "FlushStreamText");
                    Assert.AreEqual(source, markdown.SourceText);
                    var codeView = markdown.ContentPanel.Controls.OfType<Blackcancer.ChatUI.Views.ChatCodeBlockView>().Single();
                    var table = markdown.ContentPanel.Controls.OfType<Blackcancer.ChatUI.Views.ChatTableView>().Single();
                    StringAssert.Contains(codeView.SourceText, "' line 12");
                    Assert.IsFalse(codeView.Expanded, "Long code defaults to collapsed before a user chooses.");
                    Assert.AreEqual(1, table.Grid.Rows.Count);
                    codeView.Expanded = true;
                    const string delta = "\n\nAdditional detail";
                    Call(window, "ReceiveChatUpdate", "final", streamId, delta, false);
                    Assert.AreEqual(source, markdown.SourceText, "Unflushed deltas must remain batched.");
                    Call(window, "FlushStreamText");
                    Assert.AreEqual(source + delta, markdown.SourceText);
                    Assert.AreSame(codeView, markdown.ContentPanel.Controls.OfType<Blackcancer.ChatUI.Views.ChatCodeBlockView>().Single());
                    Assert.AreSame(table, markdown.ContentPanel.Controls.OfType<Blackcancer.ChatUI.Views.ChatTableView>().Single());
                    Assert.IsTrue(codeView.Expanded);
                    Call(window, "ReceiveChatUpdate", "final", streamId, source + delta, true);
                    Assert.AreSame(codeView, markdown.ContentPanel.Controls.OfType<Blackcancer.ChatUI.Views.ChatCodeBlockView>().Single());
                    Assert.IsTrue(codeView.Expanded);
                    item.RaiseEvent(new RoutedEventArgs(FrameworkElement.UnloadedEvent));
                    Assert.IsTrue(markdown.IsDisposed);
                    Assert.IsFalse(Get<Dictionary<string, Blackcancer.ChatUI.Views.ChatMarkdownView>>(window, "liveMarkdownViews").ContainsKey(streamId));
                    Assert.AreEqual(0, Get<Dictionary<ChatEntry, FrameworkElement>>(window, "entryViews").Count);
                }
            }
        }
    }
}
