namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Linq;
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

    /// <summary>Vérifie le contexte sélectionné et les jetons de référence du compositeur.</summary>
    public sealed partial class ChatWindowStateTests
    {
        /// <summary>Ne conserve que la référence sélectionnée la plus récente pour un jeton exact.</summary>
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

        /// <summary>Affiche les pièces jointes et la mémoire active sous forme de puces supprimables.</summary>
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
                typeof(Control).GetMethod("OnClick", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(chips.Controls[1].Controls.Find("open", true).Single(), new object[] { EventArgs.Empty });
                Assert.AreEqual(0, attachments.Count);
                Assert.AreEqual(1, chips.Controls.Count);
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
    using System.Windows.Controls.Primitives;
    using System.Windows.Input;
    using CodexVBE;
    using CodexVBE.Tests.Infrastructure;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    /// <summary>Vérifie les interactions clavier et la résolution des références dans le compositeur.</summary>
    public sealed partial class ChatWindowStateTests
    {
        /// <summary>Parcourt les commandes et références du catalogue avec la fenêtre contextuelle réelle.</summary>
        [STATestMethod, TestCategory("Unit")]
        public void ComposerCommandReferencePopupAndKeyboardNavigationUseRealReferenceCatalogue()
        {
            using (var runtime = new RuntimeScope())
            using (var window = new ChatWindow(runtime.Session))
            {
                var prompt = Get<TextBox>(window, "prompt"); var popup = Get<Popup>(window, "referencePopup"); var list = Get<System.Windows.Forms.ListBox>(window, "referenceList"); window.Show(); System.Windows.Forms.Application.DoEvents();
                foreach (var text in new[] { "", "ordinary", "prefix#P", "x@P", "#P", "@P", "#missing", "#P.M", "#P_M", "#P:M", "/", "/unknown", "/plan", "/plan instructions" }) { prompt.Text = text; prompt.CaretIndex = text.Length; Call(window, "UpdateReferences"); }
                prompt.Text = "#P"; prompt.CaretIndex = 2; Call(window, "UpdateReferences"); var timer = Get<System.Windows.Forms.Timer>(window, "referenceTimer"); typeof(System.Windows.Forms.Timer).GetMethod("OnTick", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).Invoke(timer, new object[] { EventArgs.Empty }); Call(window, "UpdateReferences"); Assert.IsTrue(list.Items.Count > 0);
                foreach (var key in new[] { Key.Down, Key.Up, Key.Escape }) { popup.IsOpen = true; Assert.IsTrue(RunKey(window, key).Handled); }
                popup.IsOpen = true; Assert.IsFalse(RunKey(window, Key.F1).Handled);
                prompt.Text = "/plan"; prompt.CaretIndex = prompt.Text.Length; Call(window, "UpdateReferences"); Assert.IsTrue(popup.IsOpen); Assert.IsTrue(RunKey(window, Key.Tab).Handled); Assert.AreEqual(ChatMode.Plan, Get<System.Windows.Forms.ComboBox>(window, "modePicker").SelectedItem);
                Set(window, "busy", true); prompt.Text = "/fix"; prompt.CaretIndex = prompt.Text.Length; Call(window, "UpdateReferences"); Call(window, "AcceptReference"); Assert.AreEqual(ChatMode.Plan, Get<System.Windows.Forms.ComboBox>(window, "modePicker").SelectedItem); Set(window, "busy", false);
                prompt.Text = "#P"; prompt.CaretIndex = 2; Call(window, "UpdateReferences"); Call(window, "AcceptReference"); Assert.IsTrue(Get<List<VbeChatReference>>(window, "selectedReferences").Count > 0); Assert.IsFalse(popup.IsOpen); Call(window, "UpdateReferences"); Assert.IsFalse(popup.IsOpen);
                ChatWindow.ReadModifiers = () => ModifierKeys.Shift; popup.IsOpen = true; Assert.IsFalse(RunKey(window, Key.Enter).Handled); Assert.IsFalse(popup.IsOpen); ChatWindow.ReadModifiers = () => ModifierKeys.None;
                prompt.Text = ""; Assert.IsTrue(RunKey(window, Key.Enter).Handled); Assert.IsFalse(RunKey(window, Key.F1).Handled);
                list.DataSource = new object[0]; popup.IsOpen = true; Assert.IsTrue(RunKey(window, Key.Down).Handled); Assert.IsFalse(RunKey(window, Key.Tab).Handled); Assert.IsTrue(RunKey(window, Key.Enter).Handled); Call(window, "AcceptReference");
                list.DataSource = new[] { new VbeChatReference { Project = "P" } }; list.SelectedIndex = 0; Set(window, "referenceStart", -1); Call(window, "AcceptReference");
                var disconnected = Get<VbeChatReferences>(window, "referenceIndex"); disconnected.Entries.Clear(); prompt.Text = "@none"; prompt.CaretIndex = 5; Set(window, "referenceIndexReady", true); Call(window, "UpdateReferences"); Assert.IsFalse(list.Visible);
            }
        }
        /// <summary>Vérifie les limites de résolution, la suppression des puces et la navigation vers une référence.</summary>
        [STATestMethod, TestCategory("Unit")]
        public void ComposerReferenceResolutionBoundariesChipsAndNavigationKeepLatestAttachments()
        {
            using (var runtime = new RuntimeScope())
            using (var window = new ChatWindow(runtime.Session))
            {
                var refs = Get<List<VbeChatReference>>(window, "selectedReferences"); var module = new VbeChatReference { Project = "P", Module = "M" }; refs.Add(module); refs.Add(new VbeChatReference { Project = "P", Module = "M" });
                foreach (var text in new[] { "#P.M", " #P.M!", "#P.M suffix", "prefix#P.M", "#P.Msuffix", "prefix#P.M #P.M", "missing" }) { var found = (VbeChatReference[])Call(window, "CurrentReferences", text); Assert.AreEqual(text == "missing" || text == "prefix#P.M" || text == "#P.Msuffix" ? 0 : 1, found.Length); }
                Assert.AreEqual("unchanged", Call(window, "ResolveReferences", "unchanged")); StringAssert.Contains((string)Call(window, "ResolveReferences", "#P.M"), "<references-vbe>");
                Assert.ThrowsException<System.Reflection.TargetInvocationException>(() => Call(window, "ResolveReferences", new string('x', 48001) + " #P.M")); runtime.Module.DeleteLines(1, runtime.Module.CountOfLines); runtime.Module.InsertLines(1, new string('x', 48001)); Assert.ThrowsException<System.Reflection.TargetInvocationException>(() => Call(window, "ResolveReferences", "#P.M"));
                Get<TextBox>(window, "prompt").Text = "#P.M"; Set(window, "projectMemory", "note"); Get<System.Windows.Forms.CheckBox>(window, "attachMemory").Checked = true; Call(window, "RefreshContextChips"); var chips = Get<System.Windows.Forms.FlowLayoutPanel>(window, "contextChips"); Click(chips.Controls[0].Controls.Find("open", true).Single()); Assert.IsFalse(Get<System.Windows.Forms.CheckBox>(window, "attachMemory").Checked);
                var row = chips.Controls.OfType<ChatContextChipView>().Single(); Click(row.Controls.Find("open", true).Single()); Click(row.Controls.Find("remove", true).Single()); Assert.AreEqual(0, refs.Count);
                Call(window, "NavigateReference", new VbeChatReference { Project = "P" }); Assert.AreEqual("#P.", Get<TextBox>(window, "prompt").Text);
                Call(window, "NavigateReference", new VbeChatReference { Project = "P", Module = "missing" }); StringAssert.Contains(Get<System.Windows.Forms.Label>(window, "status").Text, "missing");
            }
        }
    }
}
namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Linq;
    using System.Windows.Controls;
    using System.Windows.Controls.Primitives;
    using System.Windows.Input;
    using CodexVBE;
    using CodexVBE.Tests.Infrastructure;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    /// <summary>Vérifie les états du catalogue de références et les erreurs de navigation.</summary>
    public sealed partial class ChatWindowStateTests
    {
        /// <summary>Exerce les états chargement, résultat vide, erreur hôte et navigation du catalogue.</summary>
        [STATestMethod, TestCategory("Unit")]
        public void ComposerCatalogueStatusHandlesLoadingSuccessEmptyAndHostFailure()
        {
            using (var runtime = new RuntimeScope())
            {
                LocalizationScope.Set("ar-SA");
                using (var window = new ChatWindow(runtime.Session))
                {
                    var prompt = Get<TextBox>(window, "prompt"); var index = Get<VbeChatReferences>(window, "referenceIndex"); var list = Get<System.Windows.Forms.ListBox>(window, "referenceList"); var popup = Get<Popup>(window, "referencePopup");
                    prompt.Text = "#P"; prompt.CaretIndex = 2; Call(window, "UpdateReferences"); while (index.IsLoading) index.Step(); Set(window, "referenceIndexReady", true); Call(window, "UpdateReferences"); Assert.IsFalse(string.IsNullOrWhiteSpace(Get<System.Windows.Forms.Label>(window, "referenceStatus").Text));
                    runtime.Vbe.VBProjects[0].VBComponents.Items.Clear(); index.Refresh(); while (index.IsLoading) index.Step(); index.Refresh(); typeof(System.Windows.Forms.Timer).GetMethod("OnTick", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).Invoke(Get<System.Windows.Forms.Timer>(window, "referenceTimer"), new object[] { EventArgs.Empty }); Assert.IsFalse(index.IsLoading);
                    foreach (var text in new[] { "#P", "#missing" }) { prompt.Text = text; prompt.CaretIndex = text.Length; while (index.IsLoading) index.Step(); Set(window, "referenceIndexReady", true); Call(window, "UpdateReferences"); Assert.IsFalse(index.IsLoading); Assert.IsTrue(string.IsNullOrEmpty(index.Error)); }
                    window.Show(); Call(window, "NavigateReference", new VbeChatReference { Project = "P", Module = "missing" });
                    prompt.Text = "#P"; prompt.CaretIndex = 2; Set(window, "referenceStart", 0); list.DataSource = new[] { new VbeChatReference { Project = "P" } }; list.SelectedIndex = 0; popup.IsOpen = true; Assert.IsTrue(RunKey(window, Key.Enter).Handled);
                }
                using (var disconnected = new ChatWindow(new VbeSession(new UnavailableReferenceHost()))) { var text = Get<TextBox>(disconnected, "prompt"); text.Text = "#P"; text.CaretIndex = 2; Call(disconnected, "UpdateReferences"); Assert.IsFalse(string.IsNullOrWhiteSpace(Get<VbeChatReferences>(disconnected, "referenceIndex").Error)); }
                var host = new VbeDebugTests.FakeVbe(); var project = new VbeDebugTests.FakeProject { Name = "P", FileName = @"C:\Temp\P.xlsm", Mode = 2 }; var component = new VbeDebugTests.FakeComponent { Name = "M", Type = 1 }; component.CodeModule = new VbeDebugTests.FakeModule(component, "Sub A()\r\nEnd Sub"); project.VBComponents.Add(component); host.VBProjects.Add(project); host.ActiveVBProject = project; host.ActiveCodePane = component.CodeModule.CodePane;
                using (var navigation = new ChatWindow(new VbeSession(host))) { Call(navigation, "NavigateReference", new VbeChatReference { Project = "P", Module = "M" }); Assert.AreEqual(1, component.CodeModule.CodePane.ShowCount); }
                using (var design = new ChatWindow()) { Call(design, "UpdateReferences"); Call(design, "RefreshContextChips"); Call(design, "HideReferences"); Call(design, "DisposeComposer"); Call(design, "InitializeComposer", new object[] { null }); Call(design, "NavigateReference", new VbeChatReference { Project = "P", Module = "M" }); StringAssert.Contains(Get<System.Windows.Forms.Label>(design, "status").Text, UiText.Get("Unable to navigate: ")); }
            }
        }
    }
}
