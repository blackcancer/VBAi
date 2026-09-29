using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using VBAi;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    /// <summary>Opt-in qualification of existing accounts and explicitly authorized disposable source.</summary>
    [TestClass, TestCategory("AuthenticatedIntegration")]
    public sealed class ConnectedProviderTests
    {
        public TestContext TestContext { get; set; }

        /// <summary>Uses the production GCM and GitHub API to read the authenticated identity only.</summary>
        [TestMethod]
        public async Task ExistingGitHubAccountReadsIdentityAndAuthorizedPrivateRepository()
        {
            RequireOptIn();
            RequireSourceOptIn();
            var settings = LlmSettings.Load();
            using (var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(60)))
            {
                var accounts = await new GitHubAccountService().ListAsync(deadline.Token);
                Assert.IsTrue(accounts.Length > 0, "No existing GCM account.");
                string account = settings.GitHubAccount;
                if (string.IsNullOrWhiteSpace(account)) account = accounts.Single();
                Assert.IsTrue(accounts.Contains(account, StringComparer.OrdinalIgnoreCase));
                using (var api = new GitHubApi(account))
                {
                    var identity = await api.Request<Dictionary<string, object>>(HttpMethod.Get, "/user", null, deadline.Token);
                    Assert.AreEqual(account, Convert.ToString(identity["login"]), true);
                    var repository = await api.Request<Dictionary<string, object>>(HttpMethod.Get, "/repos/blackcancer/VBAi", null, deadline.Token);
                    Assert.AreEqual("blackcancer/VBAi", repository["full_name"]); Assert.AreEqual(true, repository["private"]);
                    var pull = await api.Request<GitHubPull>(HttpMethod.Get, "/repos/blackcancer/VBAi/pulls/10", null, deadline.Token);
                    Assert.IsTrue(pull.merged); Assert.AreEqual("closed", pull.state);
                    TestContext.WriteLine("PASS: production GCM account, authenticated identity, private repository and merged PR10; GET only.");
                }
            }
        }

        /// <summary>Reads the actual catalogue and the explicitly authorized test/test1 module using a lightweight model.</summary>
        [TestMethod]
        public async Task ExistingCodexAccountListsModelsAndReadsDisposableSolidWorksModule()
        {
            RequireOptIn();
            RequireSourceOptIn();
            int pid;
            Assert.IsTrue(int.TryParse(Environment.GetEnvironmentVariable("VBAi_SOLIDWORKS_PID"), out pid) && pid > 0);
            var account = await CodexAccount.ReadStatusAsync();
            Assert.IsTrue(account.ChatGptConnected, "The add-in private Codex account is not connected through ChatGPT.");
            var settings = LlmSettings.Load();
            var before = VbeBridgeClient.Read(pid, new { Command = "read_module", Project = "test", Module = "test1" });
            Assert.AreEqual(true, before["Ok"]);
            var tools = new LlmVbeTools(null, null, settings) { BoundProject = "test", CurrentProviderName = "Codex", Mode = ChatMode.Discussion };
            int moduleReads = 0;
            tools.Execute = request =>
            {
                var allowed = new[] { "status", "vbe_environment", "list_projects", "list_modules", "read_module" };
                if (!allowed.Contains(request.Command) ||
                    (!string.IsNullOrWhiteSpace(request.Project) && request.Project != "test") ||
                    (request.Command == "read_module" && request.Module != "test1"))
                    return Response.Failure("Only read-only metadata and the authorized disposable test/test1 module are permitted.");
                var response = VbeBridgeClient.Read(pid, request);
                if (response == null) return Response.Failure("The preloaded VBE bridge is unavailable.");
                if (request.Command == "read_module" && Convert.ToBoolean(response["Ok"])) Interlocked.Increment(ref moduleReads);
                return new Response { Ok = Convert.ToBoolean(response["Ok"]), Error = Convert.ToString(response["Error"]), Data = response["Data"] };
            };
            using (var client = new CodexAppServerClient(new SynchronizationContext(), tools, _ => { }, settings))
            {
                var models = await Bounded(client.ListModelsAsync(), 60);
                Assert.IsTrue(models.Length > 0);
                var model = models.FirstOrDefault(m => m.Id.IndexOf("mini", StringComparison.OrdinalIgnoreCase) >= 0 || m.Id.IndexOf("luna", StringComparison.OrdinalIgnoreCase) >= 0 || m.Id.IndexOf("nano", StringComparison.OrdinalIgnoreCase) >= 0);
                Assert.IsNotNull(model, "A lightweight model is required for this qualification.");
                string effort = model.Efforts.FirstOrDefault(e => e.Id == "low")?.Id ?? model.DefaultEffort;
                int activities = 0; client.ActivityUpdate += _ => Interlocked.Increment(ref activities);
                var reply = await Bounded(client.TurnAsync("Test de connexion : utilise read_module pour lire uniquement test/test1 dans le VBE (discover_tools/invoke_tool si nécessaire). Aucun autre projet, fichier, commande système, édition ou exécution. Réponds VBAi_TEST_OK puis recopie la ligne contenant Application.SldWorks.", model.Id, effort), 90);
                Assert.IsTrue(moduleReads > 0, "No successful native module read was observed.");
                StringAssert.Contains(reply, "VBAi_TEST_OK"); StringAssert.Contains(reply, "Application.SldWorks");
                Assert.IsFalse(string.IsNullOrWhiteSpace(client.ThreadId));
                var after = VbeBridgeClient.Read(pid, new { Command = "read_module", Project = "test", Module = "test1" });
                Assert.AreEqual(VbeBridgeClient.Object(before["Data"])["Sha256"], VbeBridgeClient.Object(after["Data"])["Sha256"]);
                TestContext.WriteLine("PASS: private Codex home, ChatGPT authentication, {0} models, model={1}, effort={2}, native reads={3}, activity events={4}; module unchanged.", models.Length, model.Id, effort, moduleReads, activities);
            }
        }

        private static void RequireOptIn()
        {
            if (Environment.GetEnvironmentVariable("VBAi_CONNECTED_PROVIDER_TESTS") != "1")
                Assert.Inconclusive("Explicit existing-account qualification opt-in is required.");
        }
        private static void RequireSourceOptIn()
        {
            if (Environment.GetEnvironmentVariable("VBAi_CONNECTED_SOURCE_TESTS") != "1")
                Assert.Inconclusive("Explicit permission for the private repository and disposable module payload is required.");
        }
        private static async Task<T> Bounded<T>(Task<T> operation, int seconds)
        {
            if (await Task.WhenAny(operation, Task.Delay(TimeSpan.FromSeconds(seconds))) != operation)
                throw new TimeoutException("The authenticated provider qualification exceeded its deadline.");
            return await operation;
        }
    }
}
