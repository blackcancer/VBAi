using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Diagnostics;
using System.IO;

namespace VBAi.Tests.Unit
{
    /// <summary>Vérifie l’isolation des environnements enfants et la reprise de l’historique local.</summary>
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class ProviderSessionStorageTests
    {
        /// <summary>Les chemins privés remplacent les variables héritées sans modifier le parent ou copier ses fichiers.</summary>
        [TestMethod]
        public void ChildHomesOverrideInheritedCliHomesWithoutChangingParentEnvironment()
        {
            string codex = Environment.GetEnvironmentVariable("CODEX_HOME"), copilot = Environment.GetEnvironmentVariable("COPILOT_HOME");
            try
            {
                Environment.SetEnvironmentVariable("CODEX_HOME", "external-codex");
                Environment.SetEnvironmentVariable("COPILOT_HOME", "external-copilot");
                var codexInfo = new ProcessStartInfo("codex.exe") { UseShellExecute = false };
                var copilotInfo = new ProcessStartInfo("copilot.exe") { UseShellExecute = false };
                ProviderSessionStorage.ConfigureCodex(codexInfo); ProviderSessionStorage.ConfigureCopilot(copilotInfo);
                string root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "VBAi", "Providers");
                Assert.AreEqual(Path.Combine(root, "Codex"), codexInfo.EnvironmentVariables["CODEX_HOME"]);
                Assert.AreEqual(Path.Combine(root, "Copilot"), copilotInfo.EnvironmentVariables["COPILOT_HOME"]);
                Assert.AreEqual("external-codex", Environment.GetEnvironmentVariable("CODEX_HOME"));
                Assert.AreEqual("external-copilot", Environment.GetEnvironmentVariable("COPILOT_HOME"));
                Assert.AreEqual("external-copilot", codexInfo.EnvironmentVariables["COPILOT_HOME"]);
                Assert.AreEqual("external-codex", copilotInfo.EnvironmentVariables["CODEX_HOME"]);
            }
            finally { Environment.SetEnvironmentVariable("CODEX_HOME", codex); Environment.SetEnvironmentVariable("COPILOT_HOME", copilot); }
        }

        /// <summary>Un fil extérieur devient un contexte local une seule fois ; les fils privés continuent normalement.</summary>
        [TestMethod]
        public void ExternalThreadsBecomeLocalContextWithoutLosingConversationOrExistingResumeContext()
        {
            ProviderSessionStorage.PrepareCodexSession(null);
            foreach (string id in new[] { null, "" })
            {
                var session = new ChatSessionState { CodexThreadId = id, ResumeContext = "untouched" };
                ProviderSessionStorage.PrepareCodexSession(session); Assert.AreEqual("untouched", session.ResumeContext);
            }
            var isolated = new ChatSessionState { CodexThreadId = "private", CodexThreadHome = ProviderSessionStorage.CodexHome.ToUpperInvariant() };
            ProviderSessionStorage.PrepareCodexSession(isolated); Assert.AreEqual("private", isolated.CodexThreadId); Assert.IsNull(isolated.ResumeContext);
            foreach (string home in new[] { null, "external-store" })
            {
                var legacy = new ChatSessionState
                {
                    CodexThreadId = "old",
                    CodexThreadHome = home,
                    CodexDeveloperInstructionsHash = "old-hash",
                    ResumeContext = "earlier branch"
                };
                legacy.Entries.Add(new ChatEntry { Speaker = "Vous", Text = "question été" });
                legacy.Entries.Add(new ChatEntry { Speaker = "Assistant", Text = "answer" });
                ProviderSessionStorage.PrepareCodexSession(legacy);
                Assert.IsNull(legacy.CodexThreadId); Assert.IsNull(legacy.CodexThreadHome);
                Assert.IsNull(legacy.CodexDeveloperInstructionsHash);
                StringAssert.Contains(legacy.ResumeContext, "earlier branch"); StringAssert.Contains(legacy.ResumeContext, "question été"); StringAssert.Contains(legacy.ResumeContext, "answer");
                string context = legacy.ResumeContext; ProviderSessionStorage.PrepareCodexSession(legacy); Assert.AreEqual(context, legacy.ResumeContext);
                Assert.AreEqual(2, legacy.Entries.Count);
            }
        }
    }
}
