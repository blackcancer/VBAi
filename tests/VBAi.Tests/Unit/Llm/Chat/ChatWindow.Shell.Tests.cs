namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Windows.Forms;
    using VBAi;

    /// <summary>Vérifie les actions d’édition et le changement de mode dans la fenêtre de discussion.</summary>
    public sealed partial class ChatWindowStateTests
    {
        /// <summary>Checks full-width history, exclusive transcript visibility, resizing and draft preservation without activating a desktop window.</summary>
        [STATestMethod, TestCategory("Unit")]
        public void HistoryReplacesTranscriptAndRestoresTheSameDraft()
        {
            using (var host = new HistoryTestHost())
            using (var chat = Surfaces())
            {
                chat.TopLevel = false; chat.Dock = DockStyle.Fill;
                host.Controls.Add(chat); chat.Show(); host.Show(); Application.DoEvents();
                var prompt = Get<System.Windows.Controls.TextBox>(chat, "prompt");
                prompt.Text = "Keep this unsent draft";
                Call(chat, "AddTranscriptMessage", "Vous", "Existing message");
                var history = Get<Panel>(chat, "historyPanel");
                var transcript = Get<Panel>(chat, "transcriptPanel");
                var surface = Get<Panel>(chat, "conversationPanel");
                Assert.IsFalse(history.Visible); Assert.IsTrue(transcript.Visible);
                foreach (int width in new[] { 450, 840 })
                {
                    host.ClientSize = new System.Drawing.Size(width, 900);
                    Call(chat, "History_Click", null, EventArgs.Empty); host.PerformLayout();
                    Assert.IsTrue(history.Visible); Assert.IsFalse(transcript.Visible);
                    Assert.AreEqual(surface.ClientRectangle, history.Bounds);
                    Assert.AreEqual(UiSymbol.Previous, Get<ChatActionButton>(chat, "history").Symbol);
                    Assert.IsTrue(Get<ListBox>(chat, "sessionList").Width > width - 100);
                    Call(chat, "History_Click", null, EventArgs.Empty);
                    Assert.IsFalse(history.Visible); Assert.IsTrue(transcript.Visible);
                    Assert.AreEqual(UiSymbol.History, Get<ChatActionButton>(chat, "history").Symbol);
                    Assert.AreEqual("Keep this unsent draft", prompt.Text);
                    Assert.AreEqual(1, Get<List<ChatEntry>>(chat, "transcriptEntries").Count);
                }
            }
        }

        /// <summary>Ensures every chat option has real artwork and palette refresh keeps the commands intact.</summary>
        [STATestMethod, TestCategory("Unit")]
        public void EveryChatMenuCommandHasAnOwnedImage()
        {
            using (var chat = new ChatWindow())
            {
                Call(chat, "RefreshOptionsIcons");
                Call(chat, "RefreshOptionsIcons");
                foreach (ToolStripItem item in Get<ContextMenuStrip>(chat, "optionsMenu").Items)
                    if (!(item is ToolStripSeparator)) Assert.IsNotNull(item.Image, item.Name);
            }
        }

        /// <summary>Hosts rendering assertions outside the desktop without taking keyboard focus.</summary>
        private sealed class HistoryTestHost : Form
        {
            /// <summary>Creates an offscreen test host.</summary>
            internal HistoryTestHost() { StartPosition = FormStartPosition.Manual; Location = new System.Drawing.Point(-30000, -30000); ShowInTaskbar = false; ClientSize = new System.Drawing.Size(600, 900); }
            /// <inheritdoc />
            protected override bool ShowWithoutActivation => true;
        }

        [STATestMethod, TestCategory("Unit")]
        public void GitModalOwnerPreservesStandaloneAndUncreatedChatIdentity()
        {
            using (var chat = new ChatWindow())
            {
                Assert.IsFalse(chat.IsHandleCreated);
                Assert.AreSame(chat, chat.GitModalOwner());
                Assert.IsFalse(chat.IsHandleCreated, "Resolving an uncreated owner must not create a window.");
                IntPtr handle = chat.Handle;
                Assert.AreSame(chat, chat.GitModalOwner());
                Assert.AreEqual(handle, chat.GitModalOwner().Handle);
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void GitModalOwnerUsesHostRootInsteadOfDockedChatOrFocusedComboBox()
        {
            using (var host = new Form())
            using (var chat = new ChatWindow { TopLevel = false })
            {
                host.Controls.Add(chat);
                IntPtr hostHandle = host.Handle, chatHandle = chat.Handle;
                var picker = Get<System.Windows.Forms.ComboBox>(chat, "scopePicker");
                IntPtr pickerHandle = picker.Handle;
                var owner = chat.GitModalOwner();
                Assert.AreEqual(hostHandle, owner.Handle);
                Assert.AreNotEqual(chatHandle, owner.Handle);
                Assert.AreNotEqual(pickerHandle, owner.Handle);
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void GitModalOwnerRejectsDisposedChatWithoutRecreatingItsHandle()
        {
            var chat = new ChatWindow();
            chat.Dispose();
            Assert.ThrowsException<ObjectDisposedException>(() => chat.GitModalOwner());
            Assert.IsFalse(chat.IsHandleCreated);
        }

        [STATestMethod, TestCategory("Unit")]
        public void GitShellPassesTheDockingHostRootToTheActualDialogBoundary()
        {
            using (var runtime = new RuntimeScope())
            using (var host = new Form())
            {
                string document = Path.Combine(runtime.Root, "Docked.docm");
                runtime.Vbe.VBProjects[0].FileName = document;
                runtime.Host = request => Response.Success(request.Command == "list_projects"
                    ? (object)new[] { new { Name = "P", FileName = document, Path = runtime.Root } }
                    : new { SelectedProject = "P", SelectedProjectPath = document });
                var priorCache = GitWindow.CacheDirectory;
                try
                {
                    GitWindow.CacheDirectory = key => Path.Combine(runtime.Root, "docked-git-cache");
                    int dialogs = 0;
                    using (var chat = LoadedWindow(runtime.Session))
                    {
                        chat.TopLevel = false;
                        host.Controls.Add(chat);
                        IntPtr hostHandle = host.Handle, chatHandle = chat.Handle;
                        ChatWindow.ShowModal = (dialog, owner) =>
                        {
                            Assert.IsInstanceOfType<GitWindow>(dialog);
                            Assert.AreEqual(hostHandle, owner.Handle);
                            Assert.AreNotEqual(chatHandle, owner.Handle);
                            dialogs++;
                            return DialogResult.Cancel;
                        };
                        Call(chat, "GitHub_Click", null, EventArgs.Empty);
                        Assert.AreEqual(1, dialogs, Get<System.Windows.Forms.Label>(chat, "status").Text);
                    }
                }
                finally { GitWindow.CacheDirectory = priorCache; }
            }
        }

        /// <summary>Vérifie que l’action d’éditeur prépare une commande uniquement lorsque la fenêtre est inactive.</summary>
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

        [STATestMethod]
        public void MonacoActionsAttachDraftAndChooseDiscussionOrAgentMode()
        {
            using (var runtime = new RuntimeScope())
            using (var fixture = new VBAi.Tests.Infrastructure.EditorFixture())
            using (var editor = new ModernEditorWindow())
            {
                var document = editor.OpenModule(fixture).GetAwaiter().GetResult();
                runtime.Session.ModernEditor = show => editor;
                using (var window = LoadedWindow(runtime.Session))
                {
                    foreach (string action in new[] { "/expliquer", "/corriger", "/refactoriser" })
                    {
                        window.PrepareMonacoAction(action, new ChatAttachment
                        {
                            Project = @"C:\Temp\P.xlsm",
                            Module = "M",
                            Label = "Monaco selection",
                            Text = "Debug.Print 1",
                            EditorDocumentId = document.Id,
                            Sha256 = EditorDocument.Hash(document.Text),
                            StartLine = 3
                        });
                        Assert.AreEqual(action + " ", Get<System.Windows.Controls.TextBox>(window, "prompt").Text);
                        Assert.AreEqual(action == "/expliquer" ? ChatMode.Discussion : ChatMode.Agent, Get<ComboBox>(window, "modePicker").SelectedItem);
                        var attachments = (ChatAttachment[])Call(window, "PrepareAttachments", action);
                        Assert.AreEqual(1, attachments.Length);
                        Assert.AreEqual(document.Id, attachments[0].EditorDocumentId);
                        Assert.AreEqual("Debug.Print 1", attachments[0].Text);
                    }
                    window.PrepareMonacoAction("/corriger", new ChatAttachment { Project = "Other", Text = "wrong", Label = "wrong" });
                    Assert.AreEqual("/refactoriser ", Get<System.Windows.Controls.TextBox>(window, "prompt").Text);
                }
            }
        }

        /// <summary>Vérifie que le mode de session suit la sélection uniquement lorsque la fenêtre est inactive.</summary>
        [TestMethod]
        [STATestMethod]
        public void ModeSelectionUpdatesSessionOnlyWhenIdle()
        {
            using (var window = Surfaces())
            {
                var session = new ChatSessionState
                {
                    Mode = ChatMode.Agent
                };
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
    }
}
namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Web.Script.Serialization;
    using System.Windows.Controls;
    using VBAi;
    /// <summary>Vérifie les commandes de shell, les garde-fous du concepteur et l’ouverture Git locale.</summary>
    public sealed partial class ChatWindowStateTests
    {
        /// <summary>Exerce les panneaux, modes, sessions et actions de shell ainsi que leur blocage pendant une opération.</summary>
        [STATestMethod, TestCategory("Unit")]
        public void ShellActionsTogglePanelsModeArchivePinMemoryAndDockOnlyWhenIdle()
        {
            using (var runtime = new RuntimeScope())
            using (var window = LoadedWindow(runtime.Session))
            {
                window.Show(); System.Windows.Forms.Application.DoEvents();
                Call(window, "Docking_Click", null, EventArgs.Empty); int docks = 0; window.DockRequested += () => docks++; Call(window, "Docking_Click", null, EventArgs.Empty); Set(window, "busy", true); Call(window, "Docking_Click", null, EventArgs.Empty); Assert.AreEqual(1, docks); Set(window, "busy", false); window.ReportDockFailure("test"); StringAssert.Contains(Get<System.Windows.Forms.Label>(window, "status").Text, "test");
                foreach (var method in new[] { "History_Click", "MemoryToggle_Click", "ContextToggle_Click" }) { Call(window, method, null, EventArgs.Empty); Call(window, method, null, EventArgs.Empty); }
                Get<System.Windows.Forms.Panel>(window, "historyPanel").Visible = true; Get<System.Windows.Forms.GroupBox>(window, "memoryPanel").Visible = false; Call(window, "MemoryToggle_Click", null, EventArgs.Empty); Assert.AreEqual(226F, Get<System.Windows.Forms.TableLayoutPanel>(window, "historyLayout").RowStyles[7].Height); Call(window, "MemoryToggle_Click", null, EventArgs.Empty);
                Call(window, "Options_Click", null, EventArgs.Empty); Get<System.Windows.Forms.ContextMenuStrip>(window, "optionsMenu").Close(); Call(window, "JumpToLatest_Click", null, EventArgs.Empty); Call(window, "ShowArchived_CheckedChanged", null, EventArgs.Empty);
                foreach (var mode in new[] { ChatMode.Agent, ChatMode.Discussion, ChatMode.Plan }) { Get<System.Windows.Forms.ComboBox>(window, "modePicker").SelectedItem = mode; Call(window, "ModePicker_SelectedIndexChanged", null, EventArgs.Empty); Assert.AreEqual(mode, Get<ChatSessionState>(window, "currentSession").Mode); }
                Set(window, "loadingSession", true); Call(window, "ModePicker_SelectedIndexChanged", null, EventArgs.Empty); Set(window, "loadingSession", false); Set(window, "busy", true); Call(window, "Compile_Click", null, EventArgs.Empty); Call(window, "Pin_Click", null, EventArgs.Empty); Call(window, "GitHub_Click", null, EventArgs.Empty); Set(window, "busy", false);
                Call(window, "Pin_Click", null, EventArgs.Empty); Assert.IsTrue(Get<ChatSessionState>(window, "currentSession").Pinned); Get<System.Windows.Forms.TextBox>(window, "chatTitleEditor").Text = "Renamed"; Call(window, "Rename_Click", null, EventArgs.Empty); Assert.AreEqual("Renamed", Get<ChatSessionState>(window, "currentSession").Title);
                Call(window, "SaveMemory_Click", null, EventArgs.Empty); Call(window, "AttachMemory_CheckedChanged", null, EventArgs.Empty); Call(window, "Selection_Click", null, EventArgs.Empty); Call(window, "Modules_Click", null, EventArgs.Empty); Call(window, "Methods_Click", null, EventArgs.Empty); Call(window, "Export_Click", null, EventArgs.Empty);
                var prompt = Get<TextBox>(window, "prompt"); foreach (var text in new[] { "", "word", "word " }) { prompt.Text = text; prompt.CaretIndex = text.Length; Call(window, "InsertReferencePrefix", '#'); Assert.IsTrue(prompt.Text.EndsWith("#")); }
                Call(window, "GitHub_Click", null, EventArgs.Empty); Call(window, "Compile_Click", null, EventArgs.Empty); Call(window, "NewChat_Click", null, EventArgs.Empty); Call(window, "Archive_Click", null, EventArgs.Empty);
                var args = new object[] { System.Windows.Forms.Message.Create(window.Handle, 0, IntPtr.Zero, IntPtr.Zero), System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.N }; Assert.IsTrue((bool)Call(window, "ProcessCmdKey", args)); args[1] = System.Windows.Forms.Keys.Escape; Call(window, "ProcessCmdKey", args);
            }
        }
        /// <summary>Vérifie les garde-fous sans hôte et l’ouverture Git après sauvegarde du document.</summary>
        [STATestMethod, TestCategory("Unit")]
        public void ShellDesignerGuardsAndSavedDocumentDialogUseOnlyLocalUi()
        {
            using (var runtime = new RuntimeScope())
            {
                string document = System.IO.Path.Combine(runtime.Root, "unique.xlsm"); runtime.Vbe.VBProjects[0].FileName = document; runtime.Host = r => Response.Success(r.Command == "list_projects" ? (object)new[] { new { Name = "P", FileName = document } } : new { SelectedProject = "P", SelectedProjectPath = document });
                string cache = Path.Combine(runtime.Root, "git-shell-cache"); var priorCache = GitWindow.CacheDirectory;
                GitWindow.CacheDirectory = key => { Assert.AreEqual(document, key); return cache; };
                int dialogs = 0; ChatWindow.ShowModal = (d, o) => { Assert.IsInstanceOfType<GitWindow>(d); dialogs++; return System.Windows.Forms.DialogResult.Cancel; };
                try { using (var window = LoadedWindow(runtime.Session)) { Call(window, "GitHub_Click", null, EventArgs.Empty); Assert.AreEqual(1, dialogs); Get<System.Windows.Forms.ComboBox>(window, "scopePicker").SelectedIndex = -1; Call(window, "GitHub_Click", null, EventArgs.Empty); StringAssert.Contains(Get<System.Windows.Forms.Label>(window, "status").Text, UiText.Get("The project for this conversation is closed or ambiguous.")); Assert.AreEqual(1, dialogs); } }
                finally { GitWindow.CacheDirectory = priorCache; if (System.IO.Directory.Exists(cache)) System.IO.Directory.Delete(cache, true); }
                runtime.Host = r => Response.Success(r.Command == "list_projects" ? (object)new[] { new { Name = "P", FileName = "" } } : new { SelectedProject = "P" }); using (var window = LoadedWindow(runtime.Session)) { Call(window, "GitHub_Click", null, EventArgs.Empty); StringAssert.Contains(Get<System.Windows.Forms.Label>(window, "status").Text, UiText.Get("Save the document")); }
                using (var design = new ChatWindow()) { Call(design, "InsertReferencePrefix", '#'); Call(design, "StartNewChat"); Call(design, "ContextToggle_Click", null, EventArgs.Empty); Call(design, "ContextToggle_Click", null, EventArgs.Empty); }
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void GitShellUsesNativeDocumentScopeAndPreservesLegacyBindingAndConversationGuards()
        {
            foreach (string state in new[] { "none", "exact", "legacy", "both" })
                using (var runtime = new RuntimeScope())
                {
                    var priorCache = GitWindow.CacheDirectory;
                    try
                    {
                        string document = Path.GetFullPath(Path.Combine(runtime.Root, "MixedCase", "ClasseurÉté.xlsm"));
                        runtime.Vbe.VBProjects[0].FileName = document;
                        runtime.Host = r => Response.Success(r.Command == "list_projects" ? (object)new[] { new { Name = "P", FileName = document } } :
                            new { SelectedProject = "P", SelectedProjectPath = document });
                        Func<string, string> rawCache = key => Path.Combine(runtime.Root, "git-scope-cache", Path.GetFileName(MacroGitRepository.ScopeDirectory(key)));
                        string exact = rawCache(document), legacy = rawCache(document.ToUpperInvariant());
                        byte[] binding = System.Text.Encoding.UTF8.GetBytes(new JavaScriptSerializer().Serialize(new { Remote = "https://github.com/fixture/repository.git", Branch = "main" }));
                        foreach (string cache in new[] { exact, legacy })
                            if (state == "both" || state == (cache == exact ? "exact" : "legacy"))
                            { Directory.CreateDirectory(cache); File.WriteAllBytes(Path.Combine(cache, "binding.json"), binding); }
                        var scopes = new List<string>(); int dialogs = 0;
                        GitWindow.CacheDirectory = key => { scopes.Add(key); return MacroGitRepository.ResolveScopeDirectory(key, rawCache, File.GetAttributes); };
                        ChatWindow.ShowModal = (dialog, owner) => { Assert.IsInstanceOfType<GitWindow>(dialog); dialogs++; return System.Windows.Forms.DialogResult.Cancel; };
                        using (var window = LoadedWindow(runtime.Session))
                        {
                            Call(window, "GitHub_Click", null, EventArgs.Empty);
                            CollectionAssert.AreEqual(new[] { document }, scopes.ToArray());
                            Assert.AreEqual(state == "both" ? 0 : 1, dialogs);
                            if (state == "both") StringAssert.Contains(Get<System.Windows.Forms.Label>(window, "status").Text, "both");
                            else Assert.IsTrue(Directory.Exists(state == "legacy" ? legacy : exact));
                            if (state == "none" || state == "exact") Assert.IsFalse(Directory.Exists(legacy));
                            if (state == "legacy") Assert.IsFalse(Directory.Exists(exact));
                            foreach (string cache in new[] { exact, legacy })
                                if (File.Exists(Path.Combine(cache, "binding.json"))) CollectionAssert.AreEqual(binding, File.ReadAllBytes(Path.Combine(cache, "binding.json")));
                            Set(window, "busy", true); Call(window, "GitHub_Click", null, EventArgs.Empty); Set(window, "busy", false);
                            Assert.AreEqual(1, scopes.Count, "A busy chat must not look up another cache.");
                            runtime.Host = r => Response.Success(new[] { new { Name = "P", FileName = document }, new { Name = "Other", FileName = document } });
                            Call(window, "GitHub_Click", null, EventArgs.Empty);
                            StringAssert.Contains(Get<System.Windows.Forms.Label>(window, "status").Text, UiText.Get("The project for this conversation is closed or ambiguous."));
                            Assert.AreEqual(1, scopes.Count, "An ambiguous conversation must refuse before cache lookup.");
                        }
                    }
                    finally { GitWindow.CacheDirectory = priorCache; }
                }
        }
    }
}
