namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net;
    using System.Net.Http;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Web.Script.Serialization;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using VBAi.Tests.Infrastructure;
    using System.IO;

    /// <summary>Vérifie les appels GitHub avec handlers mémoire et fixtures de processus sans accès réseau réel.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed partial class GitReviewTests
    {
        /// <summary>Teste la validation des chemins, statuts HTTP, entrées, annulation et corps de création.</summary>
        /// <returns>Tâche asynchrone du scénario.</returns>
        [TestMethod]
        public async Task EveryGitHubPathAndHttpErrorGuardRunsOnlyAgainstMemoryHandlers()
        {
            using (var scope = new LlmBoundaryScope())
            {
                using (var native = new GitHubApi(null)) { }
                Assert.AreEqual("/repos/owner/repo", GitHubApi.RepositoryPath("https://github.com/owner/repo")); Assert.AreEqual("/repos/owner/repo", GitHubApi.RepositoryPath("https://github.com/owner/repo.GIT"));
                foreach (var path in new[] { "relative", "//other.invalid", "/back\\slash", "/a/../b" }) using (var api = new GitHubApi(null, new LlmHttpFixture(), ct => Task.FromResult("fixture"))) await Assert.ThrowsExceptionAsync<ArgumentException>(() => api.Request<object>(HttpMethod.Get, path, null, CancellationToken.None));
                foreach (var status in new[] { 401, 403, 429, 404, 422, 500 }) { var handler = new LlmHttpFixture(); handler.Replies.Enqueue(new LlmHttpFixture.Reply("secret-body") { Status = (HttpStatusCode)status }); using (var api = new GitHubApi(null, handler, ct => Task.FromResult("secret-token"))) { var error = await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => api.Request<object>(HttpMethod.Get, "/fixture", null, CancellationToken.None)); StringAssert.Contains(error.Message, status.ToString()); Assert.IsFalse(error.Message.Contains("secret")); } }
                foreach (var name in new[] { null, "", "bad space", new string('x', 101) }) using (var api = new GitHubApi(null, new LlmHttpFixture(), ct => Task.FromResult("fixture"))) await Assert.ThrowsExceptionAsync<ArgumentException>(() => api.CreateRepository(name, null, false, CancellationToken.None));
                foreach (var organization in new[] { "", "owner", null }) { var handler = new LlmHttpFixture("{\"full_name\":\"owner/repo\",\"clone_url\":\"https://github.com/owner/repo.git\",\"default_branch\":\"main\"}"); using (var api = new GitHubApi(null, handler, ct => Task.FromResult("fixture"))) { var repo = await api.CreateRepository("repo", organization, true, CancellationToken.None); Assert.AreEqual("owner/repo", repo.ToString()); Assert.AreEqual("https://github.com/owner/repo.git", repo.clone_url); Assert.AreEqual("main", repo.default_branch); Assert.AreEqual(string.IsNullOrEmpty(organization) ? "/user/repos" : "/orgs/owner/repos", handler.Uris.Single().AbsolutePath); Assert.AreEqual(true, LlmBoundaryScope.Object(new JavaScriptSerializer().DeserializeObject(handler.Bodies.Single()))["private"]); Assert.AreEqual("2026-03-10", handler.Headers.Single()["X-GitHub-Api-Version"]); } }
                using (var api = new GitHubApi(null, new LlmHttpFixture(), ct => Task.FromResult("fixture"))) { await Assert.ThrowsExceptionAsync<ArgumentException>(() => api.CreateRepository("repo", "bad name", false, CancellationToken.None)); await Assert.ThrowsExceptionAsync<ArgumentException>(() => api.CreatePull("https://github.com/owner/repo.git", "feature", "main", " ", "", false, CancellationToken.None)); foreach (var sha in new[] { null, "BAD", new string('a', 39) }) await Assert.ThrowsExceptionAsync<ArgumentException>(() => api.Checks("https://github.com/owner/repo.git", sha, CancellationToken.None)); }
                using (var cancellation = new CancellationTokenSource()) { cancellation.Cancel(); using (var api = new GitHubApi(null, new LlmHttpFixture(), ct => Task.FromResult("fixture"))) await Assert.ThrowsExceptionAsync<OperationCanceledException>(() => api.Organizations(cancellation.Token)); }
                using (var cancellation = new CancellationTokenSource()) { var handler = new LlmHttpFixture("{}") { BeforeResponse = () => cancellation.Cancel() }; using (var api = new GitHubApi(null, handler, ct => Task.FromResult("fixture"))) { Exception error = null; try { await api.Request<object>(HttpMethod.Get, "/fixture", null, cancellation.Token); } catch (OperationCanceledException ex) { error = ex; } Assert.IsNotNull(error); } }
            }
        }

        /// <summary>Vérifie que les modèles de dépôt, pull request, vérifications et commentaires conservent leurs champs.</summary>
        /// <returns>Tâche asynchrone du scénario.</returns>
        [TestMethod]
        public async Task GitHubListsChecksAndDisplayModelsPreserveAllNativeShapes()
        {
            using (var scope = new LlmBoundaryScope())
            {
                var handler = new LlmHttpFixture("[{\"login\":\"organization\"}]", "[{\"name\":\"main\"}]", "[{\"number\":12,\"title\":\"title\",\"body\":\"body\",\"state\":\"open\",\"draft\":true,\"merged\":false,\"html_url\":\"https://github.com/owner/repo/pull/12\",\"head\":{\"sha\":\"abc\"}}]"); using (var api = new GitHubApi(null, handler, ct => Task.FromResult("fixture"))) { Assert.AreEqual("organization", (await api.Organizations(CancellationToken.None)).Single().login); Assert.AreEqual("main", (await api.Branches("https://github.com/owner/repo.git", CancellationToken.None)).Single().name); var pull = (await api.Pulls("https://github.com/owner/repo.git", CancellationToken.None)).Single(); StringAssert.Contains(pull.ToString(), "title"); Assert.AreEqual("body", pull.body); Assert.IsFalse(pull.merged); Assert.AreEqual("abc", pull.head.sha); StringAssert.Contains(pull.html_url, "/12"); pull.draft = false; Assert.AreEqual("#12 · open · title", pull.ToString()); }
                var first = new JavaScriptSerializer().Serialize(new { check_runs = Enumerable.Range(0, 100).Select(i => new { name = "c" + i, status = "completed", conclusion = "success" }).ToArray() }); handler = new LlmHttpFixture(first, "{\"check_runs\":[{\"name\":\"pending\",\"status\":\"queued\",\"conclusion\":null}]}", "{\"state\":\"success\"}"); using (var api = new GitHubApi(null, handler, ct => Task.FromResult("fixture"))) { var result = await api.Checks("https://github.com/owner/repo.git", new string('a', 40), CancellationToken.None); StringAssert.Contains(result, "pending: queued"); StringAssert.Contains(result, "c99: success"); Assert.AreEqual(3, handler.Uris.Count); StringAssert.Contains(handler.Uris[1].Query, "page=2"); }
                var file = new GitHubFile { filename = "module.bas", status = "modified" }; Assert.AreEqual("modified · module.bas", file.ToString()); Assert.AreEqual("comment", new GitHubComment { body = "comment" }.ToString()); Assert.AreEqual("module.bas:3 · comment", new GitHubComment { body = "comment", path = "module.bas", line = 3 }.ToString());
            }
        }

        /// <summary>Vérifie l’entrée GCM non interactive du processus natif, ses erreurs et son annulation.</summary>
        /// <returns>Tâche asynchrone du scénario.</returns>
        [TestMethod]
        public async Task NativeCredentialChildReceivesNoninteractiveInputAndNeverLaunchesGcm()
        {
            using (var scope = new LlmBoundaryScope()) using (var fixture = new NativeProtocolFixture())
            {
                var original = GitHubApi.StartCredentialProcess;
                try
                {
                    foreach (var account in new[] { null, "", "fixture-account" })
                    {
                        GitHubApi.StartCredentialProcess = p => { Assert.AreEqual("git.exe", p.StartInfo.FileName); Assert.AreEqual("credential-manager get", p.StartInfo.Arguments); Assert.AreEqual("never", p.StartInfo.EnvironmentVariables["GCM_INTERACTIVE"]); Assert.AreEqual("0", p.StartInfo.EnvironmentVariables["GIT_TERMINAL_PROMPT"]); return fixture.Start(p, "token"); }; var handler = new LlmHttpFixture("{}"); using (var api = new GitHubApi(account, handler)) { await api.Request<object>(HttpMethod.Get, "/fixture", null, CancellationToken.None); Assert.AreEqual("Bearer fixture-token", handler.Headers.Single()["Authorization"]); Assert.AreEqual("protocol=https\nhost=github.com\n" + (string.IsNullOrEmpty(account) ? "" : "username=fixture-account\n") + "\n", File.ReadAllText(Path.Combine(fixture.Root, "input.txt"))); }
                    }
                    using (var api = new GitHubApi("bad account", new LlmHttpFixture())) await Assert.ThrowsExceptionAsync<ArgumentException>(() => api.Request<object>(HttpMethod.Get, "/fixture", null, CancellationToken.None));
                    foreach (var mode in new[] { "missing", "error" }) { GitHubApi.StartCredentialProcess = p => fixture.Start(p, mode); using (var api = new GitHubApi(null, new LlmHttpFixture())) await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => api.Request<object>(HttpMethod.Get, "/fixture", null, CancellationToken.None)); }
                    using (var cancellation = new CancellationTokenSource()) { GitHubApi.StartCredentialProcess = p => { var result = fixture.Start(p, "wait"); cancellation.CancelAfter(100); return result; }; using (var api = new GitHubApi(null, new LlmHttpFixture())) await Assert.ThrowsExceptionAsync<OperationCanceledException>(() => api.Request<object>(HttpMethod.Get, "/fixture", null, cancellation.Token)); }
                    using (var cancellation = new CancellationTokenSource()) { GitHubApi.StartCredentialProcess = p => { cancellation.Cancel(); return false; }; using (var api = new GitHubApi(null, new LlmHttpFixture())) await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => api.Request<object>(HttpMethod.Get, "/fixture", null, cancellation.Token)); }
                    using (var cancellation = new CancellationTokenSource()) { GitHubApi.StartCredentialProcess = p => { var result = fixture.Start(p, "immediate"); Assert.IsTrue(p.WaitForExit(5000)); cancellation.Cancel(); return result; }; using (var api = new GitHubApi(null, new LlmHttpFixture())) await Assert.ThrowsExceptionAsync<OperationCanceledException>(() => api.Request<object>(HttpMethod.Get, "/fixture", null, cancellation.Token)); }
                }
                finally { GitHubApi.StartCredentialProcess = original; }
            }
        }
        /// <summary>Checks exact cancellation and IO outcomes independently of native pipe buffering.</summary>
        /// <returns>The asynchronous fixture scenario.</returns>
        [TestMethod]
        public async Task CredentialInputPipeFailureIsCancellationOnlyWhenItsTokenWasCancelled()
        {
            using (var scope = new LlmBoundaryScope())
            using (var fixture = new NativeProtocolFixture())
            {
                var original = GitHubApi.StartCredentialProcess;
                var originalInput = GitHubApi.WriteCredentialInput;
                try
                {
                    foreach (var cancel in new[] { false, true })
                    using (var cancellation = new CancellationTokenSource())
                    {
                        int inputCalls = 0;
                        var inputFailure = new IOException("Fixture input pipe failure.");
                        GitHubApi.WriteCredentialInput = (process, input) =>
                        {
                            inputCalls++;
                            Assert.IsTrue(process.HasExited);
                            Assert.AreEqual("protocol=https\nhost=github.com\n\n", input);
                            return Task.FromException(inputFailure);
                        };
                        GitHubApi.StartCredentialProcess = process =>
                        {
                            var started = fixture.Start(process, "wait");
                            Assert.IsTrue(started);
                            // Terminate the real fixture child; inject the input failure independently of pipe buffering.
                            process.Kill();
                            Assert.IsTrue(process.WaitForExit(5000), "The fixture child must exit before the input pipe is used.");
                            if (cancel) cancellation.Cancel();
                            return started;
                        };
                        if (cancel)
                        {
                            var error = await Assert.ThrowsExceptionAsync<OperationCanceledException>(() => GitHubApi.ReadCredential(null, cancellation.Token));
                            Assert.IsTrue(error.CancellationToken.IsCancellationRequested);
                        }
                        else
                        {
                            var error = await Assert.ThrowsExceptionAsync<IOException>(() => GitHubApi.ReadCredential(null, cancellation.Token));
                            Assert.AreSame(inputFailure, error);
                            Assert.IsFalse(cancellation.IsCancellationRequested);
                        }
                        Assert.AreEqual(1, inputCalls);
                    }
                }
                finally
                {
                    GitHubApi.WriteCredentialInput = originalInput;
                    GitHubApi.StartCredentialProcess = original;
                }
            }
        }

        /// <summary>Vérifie la pagination des dépôts et l’usage exclusif de l’hôte API GitHub.</summary>
        /// <returns>Tâche asynchrone du scénario.</returns>
        [TestMethod]
        public async Task GitHubPaginatesRepositoriesAndUsesOnlyTheApiHost()
        {
            int calls = 0;
            var handler = new Handler
            {
                Send = request =>
                {
                    Assert.AreEqual("api.github.com", request.RequestUri.Host);
                    Assert.AreEqual("Bearer", request.Headers.Authorization.Scheme);
                    calls++;
                    var count = calls == 1 ? 100 : 1;
                    string body = new JavaScriptSerializer().Serialize(Enumerable.Range(0, count).Select(i => new { full_name = "owner/repo" + calls + i, clone_url = "https://github.com/owner/repo.git", default_branch = "main" }).ToArray());
                    return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(body) });
                }
            };
            using (var api = new GitHubApi(null, handler, ct => Task.FromResult("test-token")))
            {
                Assert.AreEqual(101, (await api.Repositories(CancellationToken.None)).Length);
                Assert.AreEqual(2, calls);
            }
        }

        /// <summary>Vérifie le corps structuré d’une pull request et l’absence de tentative répétée.</summary>
        /// <returns>Tâche asynchrone du scénario.</returns>
        [TestMethod]
        public async Task PullCreationUsesStructuredBodyAndDoesNotRetry()
        {
            int calls = 0;
            var handler = new Handler
            {
                Send = async request =>
                {
                    calls++;
                    Assert.AreEqual(HttpMethod.Post, request.Method);
                    Assert.AreEqual("/repos/owner/repo/pulls", request.RequestUri.AbsolutePath);
                    var body = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(await request.Content.ReadAsStringAsync());
                    Assert.AreEqual("feature/change", body["head"]);
                    Assert.AreEqual("main", body["base"]);
                    Assert.AreEqual("A\nB", body["body"]);
                    return new HttpResponseMessage(HttpStatusCode.Created)
                    {
                        Content = new StringContent("{\"number\":12,\"title\":\"Example\"}")
                    };
                }
            };
            using (var api = new GitHubApi(null, handler, ct => Task.FromResult("test-token")))
            {
                Assert.AreEqual(12, (await api.CreatePull("https://github.com/owner/repo.git", "feature/change", "main", "Example", "A\nB", true, CancellationToken.None)).number);
                await Assert.ThrowsExceptionAsync<ArgumentException>(() => api.CreatePull("https://github.com/owner/repo.git", "main", "main", "Example", "", false, CancellationToken.None));
            }

            Assert.AreEqual(1, calls);
        }

        /// <summary>Vérifie que les erreurs ne révèlent ni corps sensibles ni identifiants et filtre les liens dangereux.</summary>
        /// <returns>Tâche asynchrone du scénario.</returns>
        [TestMethod]
        public async Task GitHubErrorsDoNotExposeResponseBodiesOrCredentials()
        {
            var handler = new Handler
            {
                Send = request => Task.FromResult(new HttpResponseMessage(HttpStatusCode.Forbidden) { Content = new StringContent("secret response") })
            };
            using (var api = new GitHubApi(null, handler, ct => Task.FromResult("test-token")))
            {
                var error = await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => api.Repositories(CancellationToken.None));
                Assert.IsFalse(error.Message.Contains("secret response"));
                Assert.IsFalse(error.Message.Contains("test-token"));
                await Assert.ThrowsExceptionAsync<ArgumentException>(() => api.Request<object>(HttpMethod.Get, "//other.example/", null, CancellationToken.None));
            }

            Assert.IsFalse(SafeLinks.Allowed("file:///C:/test.vbs"));
            Assert.IsFalse(SafeLinks.Allowed("https://token@github.com/owner/repo"));
            Assert.IsTrue(SafeLinks.Allowed("https://github.com/owner/repo"));
        }
    }
}

namespace VBAi.Tests.Unit
{
    public sealed partial class GitReviewTests
    {
        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod, Microsoft.VisualStudio.TestTools.UnitTesting.TestCategory("Unit")]
        public async System.Threading.Tasks.Task IssueTitleOverTheLimitIsRejectedBeforeCredentialsOrTransport()
        {
            int credentials = 0;
            var handler = new VBAi.Tests.Infrastructure.LlmHttpFixture();
            using (var api = new GitHubApi(null, handler, ct => { credentials++; return System.Threading.Tasks.Task.FromResult("unused"); }))
            {
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsException<System.ArgumentException>(() => api.CreateIssue("https://github.com/owner/repo", new string('t', 181), "report", System.Threading.CancellationToken.None));
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(0, credentials);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(0, handler.Uris.Count);
                await System.Threading.Tasks.Task.CompletedTask;
            }
        }
    }
}
namespace VBAi.Tests.Unit
{
    public sealed partial class GitReviewTests
    {
        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod, Microsoft.VisualStudio.TestTools.UnitTesting.TestCategory("Unit")]
        public async System.Threading.Tasks.Task IssueAcceptsOptionalBodyAndExactLimitsWhileRejectingOversizedReports()
        {
            var handler = new VBAi.Tests.Infrastructure.LlmHttpFixture("{\"number\":7,\"html_url\":\"https://github.com/owner/repo/issues/7\"}", "{\"number\":8,\"html_url\":\"https://github.com/owner/repo/issues/8\"}");
            int credentials = 0;
            using (var api = new GitHubApi(null, handler, ct => { credentials++; return System.Threading.Tasks.Task.FromResult("owned-token"); }))
            {
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsException<System.ArgumentException>(() => api.CreateIssue("https://github.com/owner/repo", "Title", new string('b', 60001), System.Threading.CancellationToken.None));
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(0, credentials);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(0, handler.Uris.Count);
                var issue = await api.CreateIssue("https://github.com/owner/repo", "Title", null, System.Threading.CancellationToken.None);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(7, issue.number);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("https://github.com/owner/repo/issues/7", issue.html_url);
                var json = new System.Web.Script.Serialization.JavaScriptSerializer();
                var payload = (System.Collections.Generic.IDictionary<string, object>)json.DeserializeObject(handler.Bodies[0]);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsNull(payload["body"]);
                issue = await api.CreateIssue("https://github.com/owner/repo", new string('t', 180), new string('b', 60000), System.Threading.CancellationToken.None);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(8, issue.number);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(2, credentials);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(2, handler.Uris.Count);
                payload = (System.Collections.Generic.IDictionary<string, object>)json.DeserializeObject(handler.Bodies[1]);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(180, ((string)payload["title"]).Length);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(60000, ((string)payload["body"]).Length);
            }
        }
    }
}