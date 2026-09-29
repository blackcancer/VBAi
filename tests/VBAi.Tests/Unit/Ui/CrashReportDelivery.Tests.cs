using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using VBAi;
using VBAi.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed partial class CrashReportDeliveryTests
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
            var fixture = new LlmHttpFixture("{\"number\":7,\"html_url\":\"https://github.com/blackcancer/VBAi/issues/7\"}");
            using (var api = new GitHubApi(null, fixture, ct => Task.FromResult("fixture-secret")))
            {
                var issue = await api.CreateIssue(CrashReport.Repository, "Report title", "Report body", CancellationToken.None);
                Assert.AreEqual(7, issue.number);
                Assert.AreEqual("/repos/blackcancer/VBAi/issues", fixture.Uris[0].AbsolutePath);
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
        [TestMethod]
        public async Task PublicationRejectsEveryUnexpectedUrlAndCancellationStopsEmailFallback()
        {
            var delivery = new CrashReportDelivery { SendOutlook = (title, body) => throw new AssertFailedException("Unexpected mail") };
            foreach (string url in new[] { null, "invalid", "http://github.com/blackcancer/VBAi/issues/1", "https://example.com/blackcancer/VBAi/issues/1", "https://github.com/other/repository/issues/1" })
            {
                delivery.Publish = (title, body, token) => Task.FromResult(url);
                Assert.AreEqual(CrashDeliveryResult.Uncertain, await delivery.Send("title", "body", "owned.md", CancellationToken.None));
            }
            delivery.Publish = (title, body, token) => Task.FromException<string>(new GitHubApiFailure(399, "unexpected"));
            Assert.AreEqual(CrashDeliveryResult.Uncertain, await delivery.Send("title", "body", "owned.md", CancellationToken.None));
            delivery.Publish = (title, body, token) => Task.FromException<string>(new CrashCredentialUnavailable());
            using (var cancel = new CancellationTokenSource())
            {
                cancel.Cancel();
                await Assert.ThrowsExceptionAsync<OperationCanceledException>(() => delivery.Send("title", "body", "owned.md", cancel.Token));
            }
        }

        [TestMethod]
        public async Task NativePublicationUsesIsolatedSettingsCredentialAndHttpAndHandlesNullIssue()
        {
            using (var fixture = new CrashReportDeliveryFixture())
            {
                var delivery = new CrashReportDelivery();
                CrashReportDelivery.LoadSettings = () => throw new IOException("owned settings failure");
                await Assert.ThrowsExceptionAsync<CrashCredentialUnavailable>(() => delivery.Publish("title", "body", CancellationToken.None));
                CrashReportDelivery.LoadSettings = () => new LlmSettings { GitHubAccount = "owned-account" };
                var http = fixture.ConfigureHttp("{\"number\":3,\"html_url\":\"https://github.com/blackcancer/VBAi/issues/3\"}");
                Assert.AreEqual(CrashReport.Repository + "/issues/3", await delivery.Publish("owned title", "owned body", CancellationToken.None));
                Assert.AreEqual("/repos/blackcancer/VBAi/issues", http.Uris[0].AbsolutePath);
                StringAssert.Contains(http.Bodies[0], "owned title"); StringAssert.Contains(http.Bodies[0], "owned body");
                fixture.ConfigureHttp("null"); Assert.IsNull(await delivery.Publish("title", "body", CancellationToken.None));
                foreach (Exception failure in new Exception[] { new OperationCanceledException("credential timeout"), new IOException("owned credential failure") })
                {
                    fixture.ConfigureHttp("{}"); CrashReportDelivery.ReadCredential = (account, token) => Task.FromException<string>(failure);
                    await Assert.ThrowsExceptionAsync<CrashCredentialUnavailable>(() => delivery.Publish("title", "body", CancellationToken.None));
                }
                using (var cancel = new CancellationTokenSource())
                {
                    cancel.Cancel(); CrashReportDelivery.ReadCredential = (account, token) => Task.FromCanceled<string>(token);
                    await Assert.ThrowsExceptionAsync<TaskCanceledException>(() => delivery.Publish("title", "body", cancel.Token));
                }
            }
        }

        [TestMethod]
        public void NativeOutlookSuccessSetsExactReportAndReleasesReferencesInReverseOwnershipOrder()
        {
            using (var fixture = new CrashReportDeliveryFixture())
            {
                Assert.IsTrue(fixture.Send()); Assert.AreEqual(1, fixture.Sends);
                Assert.AreEqual(CrashReport.Recipient, fixture.Application.Mail.To);
                Assert.AreEqual("owned title", fixture.Application.Mail.Subject); Assert.AreEqual("owned body", fixture.Application.Mail.Body);
                CollectionAssert.AreEqual(new object[] { fixture.Application.Mail, fixture.Application.Current.Accounts, fixture.Application.Current, fixture.Application }, fixture.Released);
                Assert.AreEqual(0, fixture.Profiles.Count);
            }
            using (var fixture = new CrashReportDeliveryFixture())
            {
                CrashReportDelivery.ActiveOutlook = name => throw new System.Runtime.InteropServices.COMException("no running application");
                CrashReportDelivery.ProfileCount = version => { fixture.Profiles.Add(version); return version == "15.0" ? 1 : (int?)null; };
                Assert.IsTrue(fixture.Send()); CollectionAssert.AreEqual(new[] { "16.0", "15.0" }, fixture.Profiles);
            }
        }

        [TestMethod]
        public void MissingProfilesTypesAccountsAndSessionFailuresNeverCreateDuplicateMail()
        {
            using (var fixture = new CrashReportDeliveryFixture())
            {
                CrashReportDelivery.ActiveOutlook = name => throw new System.Runtime.InteropServices.COMException("no running application");
                Assert.IsFalse(fixture.Send()); CollectionAssert.AreEqual(new[] { "16.0", "15.0", "14.0" }, fixture.Profiles);
                fixture.Profiles.Clear(); CrashReportDelivery.ProfileCount = version => 0;
                Assert.IsFalse(fixture.Send()); Assert.AreEqual(0, fixture.Sends);
                CrashReportDelivery.ProfileCount = version => 1; CrashReportDelivery.OutlookType = name => null;
                Assert.IsFalse(fixture.Send()); Assert.AreEqual(0, fixture.Sends);
                CrashReportDelivery.OutlookType = name => typeof(CrashOutlookApplication);
                CrashReportDelivery.CreateOutlook = type => throw new InvalidOperationException("owned activation failure");
                Assert.IsFalse(fixture.Send()); Assert.AreEqual(0, fixture.Sends);
            }
            using (var fixture = new CrashReportDeliveryFixture())
            {
                fixture.Application.Current.Accounts.Count = 0; Assert.IsFalse(fixture.Send()); Assert.AreEqual(0, fixture.Sends);
                Assert.AreEqual(3, fixture.Released.Count);
                fixture.Application.FailSession = true; fixture.Released.Clear(); Assert.IsFalse(fixture.Send()); Assert.AreEqual(1, fixture.Released.Count);
            }
        }

        [TestMethod]
        public void NativeSendFailureIsUncertainAndCleanupContinuesAfterAReleaseFailure()
        {
            using (var fixture = new CrashReportDeliveryFixture())
            {
                fixture.Application.Mail.Deliver = () => throw new IOException("owned Send refusal");
                Assert.ThrowsException<CrashMailUncertain>(() => fixture.Send()); Assert.AreEqual(4, fixture.Released.Count);
            }
            using (var fixture = new CrashReportDeliveryFixture())
            {
                CrashReportDelivery.ReleaseReference = reference =>
                {
                    fixture.Released.Add(reference);
                    if (reference == fixture.Application.Mail) throw new System.Runtime.InteropServices.COMException("owned release failure");
                    return 0;
                };
                Assert.IsTrue(fixture.Send()); Assert.AreEqual(4, fixture.Released.Count); Assert.AreEqual(1, fixture.Logs.Count);
                StringAssert.Contains(fixture.Logs[0], "Outlook report COM reference could not be released.");
                CrashReportDelivery.IsComReference = reference => false; fixture.Released.Clear();
                Assert.IsTrue(fixture.Send()); Assert.AreEqual(0, fixture.Released.Count);
                CrashReportDelivery.ActiveOutlook = name => null;
                Assert.IsFalse(fixture.Send()); Assert.AreEqual(0, fixture.Released.Count);
            }
        }
        [TestMethod]
        public void NativeProfileReaderDisposesOwnedProfileSnapshotsAndPreservesMissingVersusEmpty()
        {
            using (var fixture = new CrashReportDeliveryFixture())
            {
                fixture.UseNativeProfileReader();
                CrashReportDelivery.OpenProfiles = path => { StringAssert.EndsWith(path, "16.0\\Outlook\\Profiles"); return null; };
                Assert.IsNull(CrashReportDelivery.ProfileCount("16.0"));
                foreach (int count in new[] { 0, 1 })
                {
                    var profile = new CrashOutlookProfile { Count = count };
                    CrashReportDelivery.OpenProfiles = path => profile;
                    CrashReportDelivery.ReadProfileSubKeys = value => { Assert.AreSame(profile, value); return profile.Count; };
                    Assert.AreEqual((int?)count, CrashReportDelivery.ProfileCount("16.0")); Assert.IsTrue(profile.Disposed);
                }
            }
        }
    }
}
