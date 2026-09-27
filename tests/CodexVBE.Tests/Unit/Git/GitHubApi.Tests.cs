namespace CodexVBE.Tests.Unit
{
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

    [TestClass, TestCategory("Unit")]
    public sealed partial class GitReviewTests
    {
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
