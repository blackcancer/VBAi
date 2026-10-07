using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

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
            var fixture = ReadManifest("VBAi_TEST_GITHUB_MANIFEST");
            string fullName = SyntheticRepository(fixture);
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
                    Assert.AreEqual(account, fullName.Split('/')[0], true, "Synthetic repository must belong to the connected account.");
                    string route = "/repos/" + fullName;
                    var repository = await api.Request<Dictionary<string, object>>(HttpMethod.Get, route, null, deadline.Token);
                    Assert.AreEqual(fullName, repository["full_name"]);
                    Assert.AreEqual(true, repository["private"]);
                    Assert.AreEqual(Convert.ToString(fixture["repositoryId"]), Convert.ToString(repository["id"]));
                    var main = await api.Request<Dictionary<string, object>>(HttpMethod.Get, route + "/git/ref/heads/main", null, deadline.Token);
                    Assert.AreEqual(Convert.ToString(fixture["mainCommit"]), Convert.ToString(VbeBridgeClient.Object(main["object"])["sha"]));
                    TestContext.WriteLine("PASS: existing production GCM account, identity, owned synthetic private repository ID and main commit; GET only.");
                }
            }
        }

        /// <summary>Reads the actual catalogue and an explicitly manifested disposable module using a lightweight model.</summary>
        [TestMethod]
        public async Task ExistingCodexAccountListsModelsAndReadsDisposableSolidWorksModule()
        {
            RequireOptIn();
            RequireSourceOptIn();
            var fixture = ReadManifest("VBAi_TEST_SOLIDWORKS_MANIFEST");
            string project = DisposableProject(fixture);
            string module = Convert.ToString(fixture["Module"]);
            int pid = Convert.ToInt32(fixture["Pid"]);
            var account = await CodexAccount.ReadStatusAsync();
            if (!account.ChatGptConnected)
                Assert.Inconclusive("BLOCKED: the existing VBAi private Codex account is not connected; no sign-in or authentication-state migration attempted.");
            var settings = LlmSettings.Load();
            Action verify = () => VerifyDisposableHost(fixture, project, pid);
            verify();
            var before = ReadModule(pid, project, module);
            Assert.AreEqual(Convert.ToString(fixture["ModuleSha256"]), Convert.ToString(before["Sha256"]), true);
            var tools = new LlmVbeTools(null, null, settings) { BoundProject = project, CurrentProviderName = "Codex", Mode = ChatMode.Discussion };
            int moduleReads = 0;
            tools.Execute = request =>
            {
                // Never send global project lists, names, status or host paths to the provider.
                if (!AllowsDisposableRead(request.Command, request.Project, request.Module, project, module))
                    return Response.Failure("Only read_module for the explicitly authorized disposable fixture is permitted.");
                verify();
                var data = ReadModule(pid, project, module);
                Assert.AreEqual(Convert.ToString(fixture["ModuleSha256"]), Convert.ToString(data["Sha256"]), true);
                Interlocked.Increment(ref moduleReads);
                return new Response { Ok = true, Data = new { Code = data["Code"], Sha256 = data["Sha256"] } };
            };
            using (var client = new CodexAppServerClient(new SynchronizationContext(), tools, _ => { }, settings))
            {
                client.InvokeTool = (name, arguments) => AllowsProviderTool(name, arguments)
                    ? tools.InvokeAsync(name, arguments)
                    : Task.FromResult(new JavaScriptSerializer().Serialize(Response.Failure("This qualification allows only the disposable module read.")));
                var models = await Bounded(client.ListModelsAsync(), 60);
                Assert.IsTrue(models.Length > 0);
                var model = models.FirstOrDefault(m => m.Id.IndexOf("mini", StringComparison.OrdinalIgnoreCase) >= 0 || m.Id.IndexOf("luna", StringComparison.OrdinalIgnoreCase) >= 0 || m.Id.IndexOf("nano", StringComparison.OrdinalIgnoreCase) >= 0);
                Assert.IsNotNull(model, "A lightweight model is required for this qualification.");
                string effort = model.Efforts.FirstOrDefault(e => e.Id == "low")?.Id ?? model.DefaultEffort;
                int activities = 0; client.ActivityUpdate += _ => Interlocked.Increment(ref activities);
                var reply = await Bounded(client.TurnAsync("Read only the authorized disposable Project=" + project + ", Module=" + module + " through read_module (discover_tools/invoke_tool if needed). Do not list projects, inspect other files, execute code or edit anything. Reply with the exact ObservedMarker string from this module; it is not provided in this prompt.", model.Id, effort), 90);
                Assert.IsTrue(moduleReads > 0, "No successful native module read was observed.");
                StringAssert.Contains(reply, Convert.ToString(fixture["Marker"]));
                Assert.IsFalse(string.IsNullOrWhiteSpace(client.ThreadId));
                verify();
                var after = ReadModule(pid, project, module);
                Assert.AreEqual(before["Sha256"], after["Sha256"]);
                TestContext.WriteLine("PASS: private Codex home, ChatGPT authentication, {0} models, model={1}, effort={2}, native reads={3}, activity events={4}; module unchanged.", models.Length, model.Id, effort, moduleReads, activities);
            }
        }

        [TestMethod]
        public void ConnectedFixtureRejectsUnownedResourcesAndGlobalMetadata()
        {
            Assert.IsFalse(AllowsProviderTool("read_user_file", "{}"));
            Assert.IsFalse(AllowsProviderTool("invoke_tool", "{\"ToolName\":\"read_user_file\",\"ArgumentsJson\":\"{}\"}"));
            Assert.IsFalse(AllowsProviderTool("invoke_tool", "{\"ToolName\":\"invoke_tool\",\"ArgumentsJson\":\"{}\"}"));
            Assert.IsTrue(AllowsProviderTool("read_module", "{}"));
            Assert.IsFalse(AllowsDisposableRead("list_projects", "owned", "Module1", "owned", "Module1"));
            Assert.IsFalse(AllowsDisposableRead("read_module", "other", "Module1", "owned", "Module1"));
            Assert.IsFalse(AllowsDisposableRead("read_module", "owned", "OtherModule", "owned", "Module1"));
            Assert.IsFalse(AllowsDisposableRead("execute_immediate", "owned", "Module1", "owned", "Module1"));
            Assert.IsTrue(AllowsDisposableRead("read_module", "owned", "Module1", "owned", "Module1"));
            Assert.ThrowsException<InvalidOperationException>(() => SyntheticRepository(new Dictionary<string, object>
            {
                ["repositoryUrl"] = "https://github.com/owner/production",
                ["privateVerified"] = true,
                ["remoteCreated"] = true,
                ["repositoryId"] = "1",
                ["mainCommit"] = new string('a', 40)
            }));
            Assert.ThrowsException<InvalidOperationException>(() => DisposableProject(new Dictionary<string, object> { ["OwnedDisposable"] = false }));
        }

        private static bool AllowsProviderTool(string name, string arguments)
        {
            if (name == "read_module" || name == "discover_tools") return true;
            if (name != "invoke_tool") return false;
            try
            {
                var values = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(arguments);
                return values != null && values.ContainsKey("ToolName") && Convert.ToString(values["ToolName"]) == "read_module";
            }
            catch (ArgumentException) { return false; }
        }

        private static bool AllowsDisposableRead(string command, string project, string module, string ownedProject, string ownedModule)
        {
            return command == "read_module" && string.Equals(project, ownedProject, StringComparison.OrdinalIgnoreCase) && module == ownedModule;
        }

        private static Dictionary<string, object> ReadManifest(string variable)
        {
            string path = Environment.GetEnvironmentVariable(variable);
            if (string.IsNullOrWhiteSpace(path)) Assert.Inconclusive("BLOCKED: supply the explicitly owned fixture manifest through " + variable + ".");
            return new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(File.ReadAllText(path));
        }

        private static string SyntheticRepository(IDictionary<string, object> fixture)
        {
            Uri uri;
            string url = Convert.ToString(fixture["repositoryUrl"]);
            if (!Uri.TryCreate(url, UriKind.Absolute, out uri) || uri.Scheme != "https" || uri.Host != "github.com" ||
                uri.Port != 443 || uri.UserInfo != "" || uri.Query != "" || uri.Fragment != "" ||
                !Regex.IsMatch(uri.AbsolutePath, @"^/[A-Za-z0-9-]+/vbai-qualification-[A-Za-z0-9-]+$") ||
                !Convert.ToBoolean(fixture["privateVerified"]) || !Convert.ToBoolean(fixture["remoteCreated"]) ||
                !Regex.IsMatch(Convert.ToString(fixture["repositoryId"]), @"^[1-9][0-9]*$") ||
                !Regex.IsMatch(Convert.ToString(fixture["mainCommit"]), @"^[a-fA-F0-9]{40}$"))
                throw new InvalidOperationException("Only the verified owned synthetic private repository manifest is accepted.");
            return uri.AbsolutePath.TrimStart('/');
        }

        private static string DisposableProject(IDictionary<string, object> fixture)
        {
            if (!Convert.ToBoolean(fixture["OwnedDisposable"])) throw new InvalidOperationException("An owned disposable fixture is required.");
            string path = Convert.ToString(fixture["Path"]);
            if (Path.IsPathRooted(path)) path = Path.GetFullPath(path);
            if (!Path.IsPathRooted(path) || !string.Equals(Path.GetExtension(path), ".swp", StringComparison.OrdinalIgnoreCase) ||
                path.Replace('\\', '/').IndexOf("/artifacts/qualification-v1/", StringComparison.OrdinalIgnoreCase) < 0 ||
                string.IsNullOrWhiteSpace(Convert.ToString(fixture["Module"])) || string.IsNullOrWhiteSpace(Convert.ToString(fixture["Marker"])) ||
                !Regex.IsMatch(Convert.ToString(fixture["FileSha256"]), @"^[a-fA-F0-9]{64}$") ||
                !Regex.IsMatch(Convert.ToString(fixture["ModuleSha256"]), @"^[a-fA-F0-9]{64}$"))
                throw new InvalidOperationException("The disposable SWP path and its file/module hashes must be explicit.");
            return Path.GetFullPath(path);
        }

        private static void VerifyDisposableHost(IDictionary<string, object> fixture, string project, int pid)
        {
            using (var process = Process.GetProcessById(pid)) Assert.AreEqual("SLDWORKS", process.ProcessName, true);
            var file = new FileInfo(project);
            long length = file.Length;
            DateTime written = file.LastWriteTimeUtc;
            for (int read = 0; read < 2; read++)
            {
                using (var sha = SHA256.Create())
                using (var stream = new FileStream(project, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    Assert.AreEqual(Convert.ToString(fixture["FileSha256"]), BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", ""), true);
                file.Refresh();
                Assert.AreEqual(length, file.Length, "Owned SWP size changed during verification.");
                Assert.AreEqual(written, file.LastWriteTimeUtc, "Owned SWP write time changed during verification.");
            }
            var status = VbeBridgeClient.Read(pid, "status");
            Assert.IsNotNull(status, "The explicitly selected host bridge is unavailable.");
            Assert.AreEqual(true, status["Ok"]);
            var data = VbeBridgeClient.Object(status["Data"]);
            Assert.AreEqual(pid, Convert.ToInt32(data["HostProcessId"]));
            Assert.AreEqual(Convert.ToString(fixture["Mvid"]), Convert.ToString(data["AssemblyModuleVersionId"]), true);
            Assert.AreEqual(typeof(VbeSession).Module.ModuleVersionId.ToString("D"), Convert.ToString(data["AssemblyModuleVersionId"]), true);
        }

        private static IDictionary<string, object> ReadModule(int pid, string project, string module)
        {
            var response = VbeBridgeClient.Read(pid, new { Command = "read_module", Project = project, Module = module });
            Assert.IsNotNull(response);
            Assert.AreEqual(true, response["Ok"], "The exact disposable module could not be read.");
            return VbeBridgeClient.Object(response["Data"]);
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
