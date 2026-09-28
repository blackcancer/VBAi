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

    public sealed partial class ChatWindowStateTests
    {
        [TestMethod]
        [STATestMethod]
        public void ForkKeepsOnlySelectedConversationPrefixAndBuildsResumeContext()
        {
            using (var window = ReadyCodexWindow(new ChatSessionState { Scope = "temporary:test", Title = "Original" }))
            {
                window.ModelCatalogueOverride = provider => Task.FromResult(new LlmModelOption[0]);
                var original = Get<ChatSessionState>(window, "currentSession");
                original.ReadProjectGrants = new[] { "C:\\Authorized.xlsm" };
                original.SharedContextReadAllowed = true;
                Get<List<ChatSessionState>>(window, "scopeSessions").Add(original);
                var first = new ChatEntry
                {
                    Speaker = "Vous",
                    Text = "Question initiale"
                };
                var reply = new ChatEntry
                {
                    Speaker = "Assistant",
                    Text = "Réponse initiale"
                };
                Call(window, "AddEntry", first);
                Call(window, "AddEntry", reply);
                Call(window, "AddEntry", new ChatEntry { Speaker = "Vous", Text = "Suite exclue" });
                Call(window, "ForkChat", reply);
                var fork = Get<ChatSessionState>(window, "currentSession");
                Assert.AreNotSame(original, fork);
                CollectionAssert.AreEqual(original.ReadProjectGrants, fork.ReadProjectGrants);
                Assert.AreNotSame(original.ReadProjectGrants, fork.ReadProjectGrants);
                Assert.IsTrue(fork.SharedContextReadAllowed);
                Assert.AreEqual(2, fork.Entries.Count);
                StringAssert.Contains(fork.MessagesJson, "Question initiale");
                Assert.IsFalse(fork.MessagesJson.Contains("Suite exclue"));
                StringAssert.Contains(fork.ResumeContext, "Réponse initiale");
                Assert.AreEqual(3, original.Entries.Count);
                Set(window, "currentSession", null);
            }
        }

        [TestMethod]
        [STATestMethod]
        public void WorkflowRejectsStaleModuleAttachmentBeforeProviderCall()
        {
            var host = new VbeSessionTests.FakeVbe();
            var project = new VbeSessionTests.FakeProject
            {
                Name = "P",
                FileName = @"C:\Temp\P.xlsm",
                Mode = 2
            };
            project.VBComponents.Items.Add(new VbeSessionTests.FakeComponent { Name = "Module1", Type = 1, CodeModule = new VbeSessionTests.FakeModule("Sub Test()\r\nEnd Sub") });
            host.VBProjects.Add(project);
            using (var window = Surfaces())
            {
                Set(window, "scopeSession", new VbeSession(host));
                Get<List<ChatAttachment>>(window, "draftAttachments").Add(new ChatAttachment { Label = "Sélection", Text = "Sub Test()", Project = "P", Module = "Module1", Sha256 = "stale" });
                var error = Assert.ThrowsException<TargetInvocationException>(() => Call(window, "PrepareAttachments", "Question"));
                StringAssert.Contains(error.InnerException.Message, "Sélection");
            }
        }

        [TestMethod]
        [STATestMethod]
        public void VerificationReportsUnverifiedWhenNoProjectIsConnected()
        {
            using (var window = Surfaces())
            {
                CompleteOnSta((Task)Call(window, "VerifyProjectAsync"));
                var entries = Get<List<ChatEntry>>(window, "transcriptEntries");
                Assert.AreEqual(1, entries.Count);
                Assert.AreEqual("Vérification", entries[0].Speaker);
                Assert.IsFalse(string.IsNullOrWhiteSpace(entries[0].Text));
            }
        }
    }
}
namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Threading.Tasks;
    using System.Web.Script.Serialization;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    public sealed partial class ChatWindowStateTests
    {
        [STATestMethod, TestCategory("Unit")]
        public void WorkflowSelectionValidatesLocationScopeRangesAndStaleAttachments()
        {
            using (var runtime = new RuntimeScope())
            using (var window = new ChatWindow(runtime.Session))
            {
                Call(window, "CaptureSelection"); var attachments = Get<List<ChatAttachment>>(window, "draftAttachments"); Assert.AreEqual(1, attachments.Count); Assert.AreEqual("Sub A()\nEnd Sub", attachments[0].Text); Call(window, "CaptureSelection"); Assert.AreEqual(1, attachments.Count);
                foreach (var range in new[] { new[] { 1, 1, 2, 5 }, new[] { 1, 1, 2, 2 }, new[] { 0, 1, 1, 1 }, new[] { 2, 1, 1, 1 }, new[] { 1, 8, 1, 1 } })
                {
                    var prior = runtime.Host; runtime.Host = r => r.Command == "code_panes" ? Response.Success(new { ActiveCodePane = new { Properties = new { Project = "P", ProjectPath = @"C:\Temp\P.xlsm", Module = "M", Selection = new { StartLine = range[0], EndLine = range[1], StartColumn = range[2], EndColumn = range[3] } } } }) : prior(r); Call(window, "CaptureSelection"); runtime.Host = prior;
                }
                var normal = runtime.Host;
                foreach (var pane in new object[] { null, new { }, new { Properties = new { } }, new { Properties = new { Selection = new { }, Project = "Other", ProjectPath = @"C:\Temp\Other.xlsm", Module = "M" } } }) { runtime.Host = r => r.Command == "code_panes" ? Response.Success(new { ActiveCodePane = pane }) : normal(r); Call(window, "CaptureSelection"); StringAssert.Contains(Get<System.Windows.Forms.Label>(window, "status").Text, UiText.Get("Selection: ")); }
                runtime.Host = normal; Call(window, "RefreshContextPreview"); Set(window, "projectMemory", "memory"); Get<System.Windows.Forms.CheckBox>(window, "attachMemory").Checked = true; Call(window, "RefreshContextPreview"); Assert.IsTrue(Get<System.Windows.Forms.FlowLayoutPanel>(window, "contextPreview").Controls.Count >= 2);
                attachments.Clear(); attachments.Add(new ChatAttachment { Label = "plain", Text = "plain" }); Assert.AreEqual(1, ((ChatAttachment[])Call(window, "PrepareAttachments", "Question")).Length); attachments[0].Module = "M"; attachments[0].Project = "P"; attachments[0].Sha256 = "SHA"; Assert.AreEqual(1, ((ChatAttachment[])Call(window, "PrepareAttachments", "Question")).Length); attachments[0].Sha256 = "stale"; Assert.ThrowsException<System.Reflection.TargetInvocationException>(() => Call(window, "PrepareAttachments", "Question")); Call(window, "RefreshContextPreview");
                runtime.Host = r => Response.Failure("host failure"); Assert.ThrowsException<System.Reflection.TargetInvocationException>(() => Call(window, "ReadWorkflow", "read_module", "P", "M")); runtime.Host = r => Response.Success("unexpected"); Assert.ThrowsException<System.Reflection.TargetInvocationException>(() => Call(window, "ReadWorkflow", "read_module", "P", "M"));
                runtime.Host = normal; Call(window, "NavigateAttachment", new ChatAttachment { Project = "P", Module = "M", StartLine = 0, Sha256 = "sha" }); runtime.Host = r => r.Command == "select_code" ? Response.Failure("navigation failure") : normal(r); Call(window, "NavigateAttachment", new ChatAttachment { Project = "P", Module = "M", StartLine = 1 }); Assert.AreEqual("navigation failure", Get<System.Windows.Forms.Label>(window, "status").Text);
            }
        }
        [STATestMethod, TestCategory("Unit")]
        public void WorkflowVerificationExportsForksAndRollbackHandleSuccessAndFailure()
        {
            using (var runtime = new RuntimeScope())
            using (var window = new ChatWindow(runtime.Session))
            {
                var json = new JavaScriptSerializer();
                foreach (var payload in new[] { "null", json.Serialize(Response.Failure("compile failure")), json.Serialize(Response.Success((object)null)), json.Serialize(Response.Success(new { Compiled = false, Diagnostic = "diagnostic" })), json.Serialize(Response.Success(new { Compiled = true })) })
                { ChatWindow.InvokeTool = (t, n, a) => Task.FromResult(payload); CompleteOnSta((Task)Call(window, "VerifyProjectAsync")); Assert.AreEqual("Vérification", Get<List<ChatEntry>>(window, "transcriptEntries").Last().Speaker); }
                ChatWindow.InvokeTool = (t, n, a) => Task.FromResult(json.Serialize(Response.Success(new { Compiled = false, Diagnostic = "diagnostic" })));
                var host = runtime.Host; foreach (var state in new object[] { new { SelectedProjectPath = @"C:\Temp\Other.xlsm", ActiveModule = "M", Selection = new { StartLine = 1 } }, new { SelectedProjectPath = @"C:\Temp\P.xlsm", ActiveModule = "", Selection = new { StartLine = 1 } }, new { SelectedProjectPath = @"C:\Temp\P.xlsm", ActiveModule = "M", Selection = (object)null } }) { runtime.Host = r => r.Command == "debug_state" ? Response.Success(state) : host(r); CompleteOnSta((Task)Call(window, "VerifyProjectAsync")); }
                runtime.Host = r => r.Command == "debug_state" ? Response.Failure("location unavailable") : host(r); CompleteOnSta((Task)Call(window, "VerifyProjectAsync")); runtime.Host = host;
                string export = Path.Combine(runtime.Root, "export.md"); ChatWindow.ShowSaveDialog = (d, o) => { ((System.Windows.Forms.SaveFileDialog)d).FileName = export; return System.Windows.Forms.DialogResult.OK; }; Call(window, "ExportCurrentChat"); Assert.IsTrue(File.ReadAllText(export).Contains("Document:")); ChatWindow.ShowSaveDialog = (d, o) => { ((System.Windows.Forms.SaveFileDialog)d).FileName = runtime.Root; return System.Windows.Forms.DialogResult.OK; }; Call(window, "ExportCurrentChat"); StringAssert.Contains(Get<System.Windows.Forms.Label>(window, "status").Text, UiText.Get("Unable to export: "));
                var change = new CodeChange(@"C:\Temp\P.xlsm", "M", "old", "before", "new", "after", 1) { TurnId = "turn" }; var entry = new ChatEntry { Speaker = "Code", Change = change }; Call(window, "AddEntry", entry); Call(window, "ForkChat", new ChatEntry()); Call(window, "ForkChat", entry); Assert.IsNull(Get<ChatSessionState>(window, "currentSession").Entries.Last().Change);
                Get<List<CodeChange>>(window, "codeChanges").Add(change); Call(window, "RollbackIntervention", change, null, true); Assert.IsTrue(change.Restored); Call(window, "RollbackIntervention", change, 0, false);
                Set(window, "busy", true); Call(window, "ExportCurrentChat"); Call(window, "ForkChat", entry); Call(window, "RollbackIntervention", change, null, false); Set(window, "busy", false);
            }
        }
        [STATestMethod, TestCategory("Unit")]
        public void WorkflowReferencesTemporarySelectionForkHistoryAndNavigationFailureAreValidated()
        {
            using (var runtime = new RuntimeScope())
            {
                runtime.Host = r => Response.Success(r.Command == "list_projects" ? (object)new[] { new { Name = "P", FileName = "" } } : r.Command == "debug_state" ? new { SelectedProject = "P", ActiveModule = "M", Selection = new { StartLine = 1 } } : r.Command == "code_panes" ? (object)new { ActiveCodePane = new { Properties = new { Project = "P", Module = "M", Selection = new { StartLine = 1, EndLine = 1, StartColumn = 1, EndColumn = 1 } } } } : new { Code = "new", Sha256 = "sha" });
                using (var window = new ChatWindow(runtime.Session))
                {
                    var references = Get<List<VbeChatReference>>(window, "selectedReferences"); references.Add(new VbeChatReference { Project = "P", Module = "M" }); var attachments = (ChatAttachment[])Call(window, "PrepareAttachments", "#P.M"); Assert.AreEqual(1, attachments.Length); StringAssert.Contains(attachments[0].Text, "new"); Call(window, "CaptureSelection"); Assert.AreEqual(1, Get<List<ChatAttachment>>(window, "draftAttachments").Count);
                    ChatWindow.InvokeTool = (t, n, a) => Task.FromResult(new JavaScriptSerializer().Serialize(Response.Success(new { Diagnostic = "no compiled field" }))); CompleteOnSta((Task)Call(window, "VerifyProjectAsync")); Assert.IsNotNull(Get<List<ChatEntry>>(window, "transcriptEntries").Last().Attachments); var existingHost = runtime.Host; runtime.Host = r => r.Command == "debug_state" ? Response.Success(new { SelectedProject = "P", ActiveModule = "M", Selection = new { StartLine = (object)null } }) : existingHost(r); CompleteOnSta((Task)Call(window, "VerifyProjectAsync")); Assert.AreEqual(0, Get<List<ChatEntry>>(window, "transcriptEntries").Last().Attachments[0].StartLine); runtime.Host = existingHost;
                    Call(window, "AddEntry", new ChatEntry { Speaker = "Vous", Text = "request" }); var final = new ChatEntry { Speaker = "Assistant", Text = "answer" }; Call(window, "AddEntry", final); Call(window, "ForkChat", final); StringAssert.Contains(Get<ChatSessionState>(window, "currentSession").MessagesJson, "request");
                    var host = runtime.Host; runtime.Host = r => { throw new InvalidOperationException("host disconnected"); }; Call(window, "NavigateAttachment", new ChatAttachment { Project = "P", Module = "M" }); StringAssert.Contains(Get<System.Windows.Forms.Label>(window, "status").Text, "host disconnected"); Call(window, "RollbackIntervention", new CodeChange { Project = "P", Module = "M" }, null, false); StringAssert.Contains(Get<System.Windows.Forms.Label>(window, "status").Text, "host disconnected");
                    runtime.Host = host; Get<System.Windows.Forms.ComboBox>(window, "scopePicker").SelectedIndex = -1; Call(window, "CaptureSelection"); StringAssert.Contains(Get<System.Windows.Forms.Label>(window, "status").Text, UiText.Get("Selection: "));
                }
                runtime.Host = r => Response.Success(r.Command == "list_projects" ? (object)new[] { new { Name = "P", FileName = @"C:\Temp\P.xlsm" } } : r.Command == "code_panes" ? (object)new { ActiveCodePane = new { Properties = new { Project = "P", Module = "M", Selection = new { StartLine = 1, EndLine = 1, StartColumn = 1, EndColumn = 1 } } } } : new { SelectedProject = "P", SelectedProjectPath = @"C:\Temp\P.xlsm" }); using (var window = new ChatWindow(runtime.Session)) { Call(window, "CaptureSelection"); StringAssert.Contains(Get<System.Windows.Forms.Label>(window, "status").Text, UiText.Get("another document")); }
                using (var design = new ChatWindow()) { Call(design, "ExportCurrentChat"); Call(design, "ForkChat", new ChatEntry()); Call(design, "RefreshContextPreview"); }
            }
        }
    }
}
