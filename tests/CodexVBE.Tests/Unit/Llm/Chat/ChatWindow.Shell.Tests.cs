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

    /// <summary>Vérifie les actions d’édition et le changement de mode dans la fenêtre de discussion.</summary>
    public sealed partial class ChatWindowStateTests
    {
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
namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Linq;
    using System.Threading.Tasks;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Input;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    /// <summary>Vérifie les commandes de shell, les garde-fous du concepteur et l’ouverture Git locale.</summary>
    public sealed partial class ChatWindowStateTests
    {
        /// <summary>Exerce les panneaux, modes, sessions et actions de shell ainsi que leur blocage pendant une opération.</summary>
        [STATestMethod, TestCategory("Unit")]
        public void ShellActionsTogglePanelsModeArchivePinMemoryAndDockOnlyWhenIdle()
        {
            using (var runtime = new RuntimeScope())
            using (var window = new ChatWindow(runtime.Session))
            {
                window.Show(); System.Windows.Forms.Application.DoEvents();
                var button = (Button)Call(window, "ChatButton", "Primary", true); Assert.AreEqual("Primary", button.Content);
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
                string cache = MacroGitRepository.ScopeDirectory(document.ToUpperInvariant()); int dialogs = 0; ChatWindow.ShowModal = (d, o) => { Assert.IsInstanceOfType<GitWindow>(d); dialogs++; return System.Windows.Forms.DialogResult.Cancel; };
                try { using (var window = new ChatWindow(runtime.Session)) { Call(window, "GitHub_Click", null, EventArgs.Empty); Assert.AreEqual(1, dialogs); Get<System.Windows.Forms.ComboBox>(window, "scopePicker").SelectedIndex = -1; Call(window, "GitHub_Click", null, EventArgs.Empty); StringAssert.Contains(Get<System.Windows.Forms.Label>(window, "status").Text, UiText.Get("Save the document")); } }
                finally { if (System.IO.Directory.Exists(cache)) System.IO.Directory.Delete(cache, true); }
                runtime.Host = r => Response.Success(r.Command == "list_projects" ? (object)new[] { new { Name = "P", FileName = "" } } : new { SelectedProject = "P" }); using (var window = new ChatWindow(runtime.Session)) { Call(window, "GitHub_Click", null, EventArgs.Empty); StringAssert.Contains(Get<System.Windows.Forms.Label>(window, "status").Text, UiText.Get("Save the document")); }
                using (var design = new ChatWindow()) { Call(design, "InsertReferencePrefix", '#'); Call(design, "StartNewChat"); Call(design, "ContextToggle_Click", null, EventArgs.Empty); Call(design, "ContextToggle_Click", null, EventArgs.Empty); }
            }
        }
    }
}
