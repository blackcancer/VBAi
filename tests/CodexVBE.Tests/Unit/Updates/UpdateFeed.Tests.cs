using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class UpdateFeedTests
    {
        internal sealed class Handler : HttpMessageHandler
        {
            internal readonly Queue<HttpResponseMessage> Replies = new Queue<HttpResponseMessage>();
            internal readonly List<Uri> Uris = new List<Uri>();
            internal readonly List<string> Tokens = new List<string>();
            internal Handler(params string[] replies) { foreach (string body in replies) Add(body); }
            internal void Add(string body, HttpStatusCode status = HttpStatusCode.OK) { Replies.Enqueue(new HttpResponseMessage(status) { Content = new StringContent(body, Encoding.UTF8) }); }
            internal void Redirect(string url) { var reply = new HttpResponseMessage(HttpStatusCode.Found); reply.Headers.Location = new Uri(url); Replies.Enqueue(reply); }
            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken ct)
            { ct.ThrowIfCancellationRequested(); Uris.Add(request.RequestUri); Tokens.Add(request.Headers.Authorization?.ToString()); return Task.FromResult(Replies.Dequeue()); }
        }
        internal static UpdateAsset Asset(string body) { using (var scope = new UpdateScope()) { string path = Path.Combine(scope.Root, "bytes"); File.WriteAllBytes(path, Encoding.UTF8.GetBytes(body)); return new UpdateAsset { id = 7, name = "VBAi-Setup-win-x64.exe", size = Encoding.UTF8.GetByteCount(body), digest = "sha256:" + UpdatePaths.Hash(path) }; } }
        [TestMethod]
        public async Task PublicReleaseDiscoveryNeverRequestsCredentialsAndFiltersDraftsPreviewsAndSkippedVersions()
        {
            string body = "[{\"tag_name\":\"v1.10.0\"},{\"tag_name\":\"v1.2.0\"},{\"tag_name\":\"v9.0.0\",\"draft\":true},{\"tag_name\":\"v2.0.0-beta.2\",\"prerelease\":true},{\"tag_name\":\"bad\"}]";
            var handler = new Handler(body, body, body);
            using (var feed = new UpdateFeed(handler, ct => throw new AssertFailedException("Public repository requested credentials")))
            {
                Assert.AreEqual("1.10.0", (await feed.Check(UpdateVersion.Parse("1.0.0"), false, null, CancellationToken.None)).Version.Text);
                Assert.AreEqual("2.0.0-beta.2", (await feed.Check(UpdateVersion.Parse("1.0.0"), true, null, CancellationToken.None)).Version.Text);
                Assert.AreEqual("1.2.0", (await feed.Check(UpdateVersion.Parse("1.0.0"), false, "v1.10.0", CancellationToken.None)).Version.Text);
                Assert.IsTrue(handler.Tokens.All(x => x == null));
            }
            var pages = new Handler(new JavaScriptSerializer().Serialize(Enumerable.Repeat(new { tag_name = "0.0.1" }, 100).ToArray()), "[{\"tag_name\":\"1.1.0\"}]");
            using (var feed = new UpdateFeed(pages)) Assert.AreEqual("1.1.0", (await feed.Check(UpdateVersion.Parse("1.0.0"), false, null, CancellationToken.None)).Version.Text);
            StringAssert.Contains(pages.Uris[1].Query, "page=2");
        }
        [TestMethod]
        public async Task PrivateAssetsAuthenticateOnlyToGitHubAndVerifyCachedBytes()
        {
            const string bytes = "fixture installer"; var asset = Asset(bytes);
            var handler = new Handler(); handler.Add("Denied", HttpStatusCode.NotFound); handler.Redirect("https://release-assets.githubusercontent.com/fixture?signature=test"); handler.Add(bytes);
            int credentials = 0;
            using (var scope = new UpdateScope())
            using (var feed = new UpdateFeed(handler, ct => { credentials++; return Task.FromResult("fixture-token"); }))
            {
                string path = await feed.Download(asset, scope.Root, null, CancellationToken.None);
                Assert.AreEqual(asset.Hash, UpdatePaths.Hash(path)); Assert.AreEqual(1, credentials);
                Assert.IsNull(handler.Tokens[0]); Assert.AreEqual("Bearer fixture-token", handler.Tokens[1]); Assert.IsNull(handler.Tokens[2]);
                Assert.AreEqual(path, await feed.Download(asset, scope.Root, null, CancellationToken.None)); Assert.AreEqual(3, handler.Uris.Count);
                File.WriteAllText(path, "tampered"); handler.Add(bytes);
                Assert.AreEqual(path, await feed.Download(asset, scope.Root, null, CancellationToken.None)); Assert.AreEqual(asset.Hash, UpdatePaths.Hash(path));
            }
        }
        [TestMethod]
        public async Task InvalidRedirectsChecksumsSizesAndCancellationNeverProduceAnInstaller()
        {
            foreach (string url in new[] { "https://evil.example/asset", "http://release-assets.githubusercontent.com/asset", "https://user@release-assets.githubusercontent.com/asset" })
            using (var scope = new UpdateScope())
            {
                var handler = new Handler(); handler.Redirect(url);
                using (var feed = new UpdateFeed(handler)) await Assert.ThrowsExceptionAsync<InvalidDataException>(() => feed.Download(Asset("fixture"), scope.Root, null, CancellationToken.None));
                Assert.AreEqual(1, handler.Uris.Count); Assert.AreEqual(0, Directory.GetFiles(scope.Root, "*.part", SearchOption.AllDirectories).Length);
            }
            using (var scope = new UpdateScope())
            using (var feed = new UpdateFeed(new Handler("changed", "short")))
            {
                await Assert.ThrowsExceptionAsync<InvalidDataException>(() => feed.Download(Asset("fixture"), scope.Root, null, CancellationToken.None));
                await Assert.ThrowsExceptionAsync<InvalidDataException>(() => feed.Download(Asset("fixture"), scope.Root, null, CancellationToken.None));
                var invalid = Asset("fixture"); invalid.digest = null;
                await Assert.ThrowsExceptionAsync<InvalidDataException>(() => feed.Download(invalid, scope.Root, null, CancellationToken.None));
                Assert.AreEqual(0, Directory.GetFiles(scope.Root, "VBAi-Setup*", SearchOption.AllDirectories).Length);
                using (var cancellation = new CancellationTokenSource()) { cancellation.Cancel(); await Assert.ThrowsExceptionAsync<OperationCanceledException>(() => feed.Check(UpdateVersion.Parse("1.0.0"), false, null, cancellation.Token)); }
            }
        }
    }
}
