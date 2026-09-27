using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class GitReviewTests
    {
        [TestMethod]
        public void PartialSnapshotsKeepFormResourcesAndReferencesConsistent()
        {
            VbaGitSnapshot make(byte resource, string reference) { return new VbaGitSnapshot(new VbaGitManifest { References = reference, Components = new[] {
                new VbaGitComponent { Name = "Form1", Type = 3, HasResources = true }, new VbaGitComponent { Name = "Module1", Type = 1 } } },
                new Dictionary<string, byte[]> { { "Form1.frm", VbaGitSnapshot.Utf8.GetBytes("Attribute VB_Name = \"Form1\"\n") }, { "Form1.frx", new[] { resource } }, { "Module1.bas", VbaGitSnapshot.Utf8.GetBytes("Attribute VB_Name = \"Module1\"\n") } }); }
            var old = make(1, "a"); var current = make(2, "b");
            var selected = VbaGitSnapshot.Select(old, current, new[] { "Form1" });
            Assert.AreEqual((byte)2, selected.Files["Form1.frx"][0]); Assert.AreEqual("a", selected.Manifest.References);
            Assert.AreEqual("b", VbaGitSnapshot.Select(old, current, new string[0], true).Manifest.References);
            Assert.ThrowsException<ArgumentException>(() => VbaGitSnapshot.Select(old, current, new[] { "Unknown" }));
        }
        [TestMethod]
        public void DiffIncludesLargeModulesAndFoldsOnlyUnchangedContext()
        {
            string before = string.Join("\n", Enumerable.Range(0, 5000).Select(x => "line " + x));
            string after = before.Replace("line 4200\n", "changed\n");
            var rows = DiffModel.Build(before, after, false, false);
            Assert.AreEqual(5000, rows.Count); Assert.AreEqual("changed", rows[4200].Right);
            var folded = DiffModel.Build(before, after, true, true);
            Assert.IsTrue(folded.Count < 20); Assert.IsTrue(folded.Any(x => x.Left == "line 4200")); Assert.IsTrue(folded.Any(x => x.Right == "changed"));
            Assert.AreEqual(1, folded.Count(x => x.Right == "changed"));
        }
        private sealed class Handler : HttpMessageHandler
        {
            internal Func<HttpRequestMessage, Task<HttpResponseMessage>> Send;
            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken token) { token.ThrowIfCancellationRequested(); return Send(request); }
        }
        [TestMethod]
        public async Task GitHubPaginatesRepositoriesAndUsesOnlyTheApiHost()
        {
            int calls = 0;
            var handler = new Handler { Send = request => {
                Assert.AreEqual("api.github.com", request.RequestUri.Host); Assert.AreEqual("Bearer", request.Headers.Authorization.Scheme);
                calls++; var count = calls == 1 ? 100 : 1;
                string body = new JavaScriptSerializer().Serialize(Enumerable.Range(0, count).Select(i => new { full_name = "owner/repo" + calls + i, clone_url = "https://github.com/owner/repo.git", default_branch = "main" }).ToArray());
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(body) });
            } };
            using (var api = new GitHubApi(null, handler, ct => Task.FromResult("test-token")))
            { Assert.AreEqual(101, (await api.Repositories(CancellationToken.None)).Length); Assert.AreEqual(2, calls); }
        }
        [TestMethod]
        public async Task PullCreationUsesStructuredBodyAndDoesNotRetry()
        {
            int calls = 0;
            var handler = new Handler { Send = async request => {
                calls++; Assert.AreEqual(HttpMethod.Post, request.Method); Assert.AreEqual("/repos/owner/repo/pulls", request.RequestUri.AbsolutePath);
                var body = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(await request.Content.ReadAsStringAsync());
                Assert.AreEqual("feature/change", body["head"]); Assert.AreEqual("main", body["base"]); Assert.AreEqual("A\nB", body["body"]);
                return new HttpResponseMessage(HttpStatusCode.Created) { Content = new StringContent("{\"number\":12,\"title\":\"Example\"}") };
            } };
            using (var api = new GitHubApi(null, handler, ct => Task.FromResult("test-token")))
            {
                Assert.AreEqual(12, (await api.CreatePull("https://github.com/owner/repo.git", "feature/change", "main", "Example", "A\nB", true, CancellationToken.None)).number);
                await Assert.ThrowsExceptionAsync<ArgumentException>(() => api.CreatePull("https://github.com/owner/repo.git", "main", "main", "Example", "", false, CancellationToken.None));
            }
            Assert.AreEqual(1, calls);
        }
        [TestMethod]
        public async Task GitHubErrorsDoNotExposeResponseBodiesOrCredentials()
        {
            var handler = new Handler { Send = request => Task.FromResult(new HttpResponseMessage(HttpStatusCode.Forbidden) { Content = new StringContent("secret response") }) };
            using (var api = new GitHubApi(null, handler, ct => Task.FromResult("test-token")))
            {
                var error = await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => api.Repositories(CancellationToken.None));
                Assert.IsFalse(error.Message.Contains("secret response")); Assert.IsFalse(error.Message.Contains("test-token"));
                await Assert.ThrowsExceptionAsync<ArgumentException>(() => api.Request<object>(HttpMethod.Get, "//other.example/", null, CancellationToken.None));
            }
            Assert.IsFalse(SafeLinks.Allowed("file:///C:/test.vbs")); Assert.IsFalse(SafeLinks.Allowed("https://token@github.com/owner/repo")); Assert.IsTrue(SafeLinks.Allowed("https://github.com/owner/repo"));
        }
    }
}
