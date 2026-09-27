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
