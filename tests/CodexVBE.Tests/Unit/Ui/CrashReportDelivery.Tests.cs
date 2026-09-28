using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class CrashReportDeliveryTests
    {
        [TestMethod]
        public async Task GitHubSuccessUsesNoEmailAndRejectsUnexpectedIssueUrls()
        {
            int emails = 0;
            var delivery = new CrashReportDelivery {
                Publish = (title, body, ct) => Task.FromResult(CrashReport.Repository + "/issues/7"),
                SendOutlook = (title, body) => { emails++; return true; },
                OpenDraft = url => Assert.Fail("Unexpected draft")
            };
            Assert.AreEqual(CrashDeliveryResult.GitHub, await delivery.Send("Title", "Full body", "local.md", CancellationToken.None));
            Assert.AreEqual(0, emails);
            delivery.Publish = (title, body, ct) => Task.FromResult("https://example.com/issue");
            Assert.AreEqual(CrashDeliveryResult.Uncertain, await delivery.Send("Title", "Body", "local.md", CancellationToken.None));
            Assert.AreEqual(0, emails);
        }
        [TestMethod]
        public async Task RefusalsUseOutlookThenDraftButAmbiguousFailuresNeverSendAgain()
        {
            int emails = 0, drafts = 0; string opened = null, sentBody = null;
            var delivery = new CrashReportDelivery {
                Publish = (title, body, ct) => Task.FromException<string>(new GitHubApiFailure(403, "Denied")),
                SendOutlook = (title, body) => { emails++; sentBody = body; return true; },
                OpenDraft = url => { drafts++; opened = url; }
            };
            foreach (int code in new[] { 401, 403, 404, 422, 429 })
            {
                delivery.Publish = (title, body, ct) => Task.FromException<string>(new GitHubApiFailure(code, "Denied"));
                Assert.AreEqual(CrashDeliveryResult.Outlook, await delivery.Send("Title", "Full body", "report.md", CancellationToken.None));
            }
            Assert.AreEqual("Full body", sentBody); Assert.AreEqual(5, emails); Assert.AreEqual(0, drafts);
            delivery.Publish = (title, body, ct) => Task.FromException<string>(new CrashCredentialUnavailable());
            delivery.SendOutlook = (title, body) => { emails++; return false; };
            Assert.AreEqual(CrashDeliveryResult.Draft, await delivery.Send("Title & test", "Full body", "report.md", CancellationToken.None));
            StringAssert.StartsWith(opened, "mailto:" + CrashReport.Recipient); StringAssert.Contains(opened, "Title%20%26%20test"); StringAssert.Contains(Uri.UnescapeDataString(opened), "report.md");
            Assert.AreEqual(1, drafts); int before = emails;
            foreach (Exception error in new Exception[] { new HttpRequestException("Timeout"), new OperationCanceledException(), new GitHubApiFailure(500, "Unavailable") })
            {
                delivery.Publish = (title, body, ct) => Task.FromException<string>(error);
                Assert.AreEqual(CrashDeliveryResult.Uncertain, await delivery.Send("Title", "Body", "report.md", CancellationToken.None));
            }
            Assert.AreEqual(before, emails); Assert.AreEqual(1, drafts);
            delivery.Publish = (title, body, ct) => Task.FromException<string>(new CrashCredentialUnavailable());
            delivery.SendOutlook = (title, body) => throw new CrashMailUncertain();
            Assert.AreEqual(CrashDeliveryResult.Uncertain, await delivery.Send("Title", "Body", "report.md", CancellationToken.None));
            Assert.AreEqual(1, drafts);
        }
        [TestMethod]
        public async Task IssueApiSerializesExactlyTheReportToTheProductRepository()
        {
            var fixture = new LlmHttpFixture("{\"number\":7,\"html_url\":\"https://github.com/blackcancer/CodexVBE/issues/7\"}");
            using (var api = new GitHubApi(null, fixture, ct => Task.FromResult("fixture-secret")))
            {
                var issue = await api.CreateIssue(CrashReport.Repository, "Report title", "Report body", CancellationToken.None);
                Assert.AreEqual(7, issue.number);
                Assert.AreEqual("/repos/blackcancer/CodexVBE/issues", fixture.Uris[0].AbsolutePath);
                StringAssert.Contains(fixture.Bodies[0], "Report title"); StringAssert.Contains(fixture.Bodies[0], "Report body");
                Assert.IsFalse(fixture.Bodies[0].Contains("fixture-secret"));
                await Assert.ThrowsExceptionAsync<ArgumentException>(() => api.CreateIssue(CrashReport.Repository, "", "", CancellationToken.None));
            }
            var denied = new LlmHttpFixture(); denied.Replies.Enqueue(new LlmHttpFixture.Reply("secret response") { Status = HttpStatusCode.Forbidden });
            using (var api = new GitHubApi(null, denied, ct => Task.FromResult("secret-token")))
            {
                var error = await Assert.ThrowsExceptionAsync<GitHubApiFailure>(() => api.CreateIssue(CrashReport.Repository, "Title", "Body", CancellationToken.None));
                Assert.AreEqual(403, error.Status); Assert.IsFalse(error.Message.Contains("secret"));
            }
        }
    }
}
