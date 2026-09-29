using System;
using System.IO;
using System.Threading.Tasks;
using VBAi;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace VBAi.Tests.Unit.Updates
{
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed partial class WebViewRuntimePrerequisiteTests
    {
        [TestMethod]
        public async Task TimedOutRuntimeInstallationRetainsPayloadAndOriginalFailure()
        {
            using (var fixture = new UpdatesNativeFixture())
            {
                string payload = null; int attempts = 0;
                var failure = new TimeoutException("owned uncertain installer");
                var runtime = new WebViewRuntimePrerequisite {
                    IsInstalled = () => false,
                    Download = path => { payload = path; File.WriteAllText(path, "owned bootstrapper"); return Task.CompletedTask; },
                    Verify = path => true,
                    Install = path => { attempts++; throw failure; }
                };
                var actual = await Assert.ThrowsExceptionAsync<TimeoutException>(() => runtime.Ensure(fixture.Scope.Root));
                Assert.AreSame(failure, actual);
                Assert.AreEqual(1, attempts);
                Assert.IsTrue(File.Exists(payload));
                Assert.AreEqual("owned bootstrapper", File.ReadAllText(payload));
            }
        }

        [TestMethod]
        public async Task CleanupFailureDoesNotHideTheOriginalRuntimeFailure()
        {
            using (var fixture = new UpdatesNativeFixture())
            {
                FileStream locked = null;
                var failure = new InvalidDataException("owned verification refusal");
                var runtime = new WebViewRuntimePrerequisite {
                    IsInstalled = () => false,
                    Download = path => { locked = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None); return Task.CompletedTask; },
                    Verify = path => throw failure,
                    Install = path => throw new AssertFailedException("No unverified installation")
                };
                try
                {
                    var actual = await Assert.ThrowsExceptionAsync<InvalidDataException>(() => runtime.Ensure(fixture.Scope.Root));
                    Assert.AreSame(failure, actual);
                }
                finally { locked?.Dispose(); }
            }
        }

        [TestMethod]
        public async Task InstalledRuntimeSkipsDownloadAndInstallation()
        {
            var runtime = new WebViewRuntimePrerequisite { IsInstalled = () => true, Download = p => throw new AssertFailedException("Unexpected download") };
            await runtime.Ensure(Path.GetTempPath());
            Assert.IsTrue(WebViewRuntimePrerequisite.ValidVersion("120.0.1.2"));
            Assert.IsFalse(WebViewRuntimePrerequisite.ValidVersion("0.0.0.0"));
            Assert.IsFalse(WebViewRuntimePrerequisite.ValidVersion(null));
        }
        [TestMethod]
        public async Task UnsignedInstallerIsNeverExecuted()
        {
            string downloaded = null;
            var runtime = new WebViewRuntimePrerequisite { IsInstalled = () => false, Download = p => { downloaded = p; File.WriteAllText(p, "fixture"); return Task.CompletedTask; }, Verify = p => false, Install = p => throw new AssertFailedException("Untrusted execution") };
            await Assert.ThrowsExceptionAsync<InvalidDataException>(() => runtime.Ensure(Path.GetTempPath()));
            Assert.IsFalse(File.Exists(downloaded));
        }
        [TestMethod]
        public async Task ExitSuccessRequiresInstalledRuntimeReadback()
        {
            var runtime = new WebViewRuntimePrerequisite { IsInstalled = () => false, Download = p => { File.WriteAllText(p, "fixture"); return Task.CompletedTask; }, Verify = p => true, Install = p => 0 };
            await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => runtime.Ensure(Path.GetTempPath()));
        }
        [TestMethod]
        public void RuntimeRegistryProbeChecksBothHivesAndViewsAndDisposesEachOwnedKey()
        {
            using (var fixture = new UpdatesNativeFixture())
            {
                var roots = new System.Collections.Generic.List<UpdateOwnedRegistry>();
                var keys = new System.Collections.Generic.List<UpdateOwnedRegistry>();
                var probes = new System.Collections.Generic.List<string>();
                object[] values = { null, 7, "0.0.0.0", "120.0.1.2" }; int index = 0;
                WebViewRuntimePrerequisite.OpenRegistryRoot = (hive, view) => { probes.Add(hive + ":" + view); var root = new UpdateOwnedRegistry(); roots.Add(root); return root; };
                WebViewRuntimePrerequisite.OpenRuntimeKey = (root, path) =>
                {
                    StringAssert.Contains(path, "EdgeUpdate\\Clients\\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}");
                    object value = values[index++]; if (value == null) return null;
                    var key = new UpdateOwnedRegistry { Value = value }; keys.Add(key); return key;
                };
                WebViewRuntimePrerequisite.ReadRuntimeVersion = key => ((UpdateOwnedRegistry)key).Value;
                Assert.IsTrue(WebViewRuntimePrerequisite.Installed());
                CollectionAssert.AreEqual(new[] { "CurrentUser:Registry32", "CurrentUser:Registry64", "LocalMachine:Registry32", "LocalMachine:Registry64" }, probes);
                Assert.IsTrue(roots.TrueForAll(root => root.Disposed)); Assert.IsTrue(keys.TrueForAll(key => key.Disposed));
                WebViewRuntimePrerequisite.OpenRuntimeKey = (root, path) => null;
                Assert.IsFalse(WebViewRuntimePrerequisite.Installed()); Assert.AreEqual(8, roots.Count);
                Assert.IsTrue(roots.TrueForAll(root => root.Disposed));
            }
        }

        [TestMethod]
        public async Task RuntimeEnsureCleansMissingDownloadsAndChecksExitAndInstalledReadback()
        {
            using (var fixture = new UpdatesNativeFixture())
            {
                var runtime = new WebViewRuntimePrerequisite { IsInstalled = () => false, Download = path => Task.FromException(new IOException("owned download failure")) };
                await Assert.ThrowsExceptionAsync<IOException>(() => runtime.Ensure(fixture.Scope.Root));
                runtime.Download = path => { File.WriteAllText(path, "owned bootstrapper"); return Task.CompletedTask; };
                runtime.Verify = path => true; runtime.Install = path => 1;
                await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => runtime.Ensure(fixture.Scope.Root));
                int checks = 0, installed = 0; runtime.IsInstalled = () => checks++ != 0;
                runtime.Install = path => { installed++; return 0; };
                await runtime.Ensure(fixture.Scope.Root); Assert.AreEqual(1, installed); Assert.AreEqual(2, checks);
                Assert.AreEqual(0, Directory.GetDirectories(fixture.Scope.Root, "WebView2-*").Length);
            }
        }

        [TestMethod]
        public async Task NativeRuntimeDownloadEnforcesHttpsHttpStatusAndBootstrapperSizeWithLocalHttp()
        {
            using (var fixture = new UpdatesNativeFixture())
            {
                using (var client = WebViewRuntimePrerequisite.CreateClient()) Assert.AreEqual(TimeSpan.FromMinutes(2), client.Timeout);
                var runtime = new WebViewRuntimePrerequisite();
                foreach (int outcome in new[] { 0, 1, 2, 3, 4 })
                {
                    var handler = new UpdateRuntimeHttp();
                    if (outcome == 1) handler.FinalUrl = "http://example.invalid/runtime";
                    if (outcome == 2) handler.Status = System.Net.HttpStatusCode.Forbidden;
                    if (outcome == 3) handler.Bytes = new byte[20 * 1024 * 1024 + 1];
                    if (outcome == 4) handler.Bytes = new byte[0];
                    WebViewRuntimePrerequisite.CreateClient = () => new System.Net.Http.HttpClient(handler);
                    string path = Path.Combine(fixture.Scope.Root, "bootstrapper-" + outcome + ".exe");
                    if (outcome == 1 || outcome == 3) await Assert.ThrowsExceptionAsync<InvalidDataException>(() => runtime.Download(path));
                    else if (outcome == 2) await Assert.ThrowsExceptionAsync<System.Net.Http.HttpRequestException>(() => runtime.Download(path));
                    else { await runtime.Download(path); CollectionAssert.AreEqual(handler.Bytes, File.ReadAllBytes(path)); }
                    Assert.AreEqual(1, handler.Calls);
                }
            }
        }

        [TestMethod]
        public void NativeCertificateVerificationChecksTrustAndExactMicrosoftSubjectWithoutCertificateStoreWrites()
        {
            using (var fixture = new UpdatesNativeFixture())
            {
                string signed = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "dotnet", "dotnet.exe");
                using (var actual = WebViewRuntimePrerequisite.LoadCertificate(signed)) Assert.IsFalse(string.IsNullOrEmpty(actual.Subject));
                var runtime = new WebViewRuntimePrerequisite();
                WebViewRuntimePrerequisite.VerifyTrust = path => false;
                Assert.IsFalse(runtime.Verify(signed));
                WebViewRuntimePrerequisite.VerifyTrust = path => true;
                foreach (string subject in new[] { "Microsoft Corporation", "Owned Fixture Vendor" })
                using (var rsa = System.Security.Cryptography.RSA.Create())
                {
                    rsa.KeySize = 2048;
                    var request = new System.Security.Cryptography.X509Certificates.CertificateRequest("CN=" + subject, rsa, System.Security.Cryptography.HashAlgorithmName.SHA256, System.Security.Cryptography.RSASignaturePadding.Pkcs1);
                    var certificate = request.CreateSelfSigned(DateTimeOffset.UtcNow.AddDays(-1), DateTimeOffset.UtcNow.AddDays(1));
                    WebViewRuntimePrerequisite.LoadCertificate = path => certificate;
                    Assert.AreEqual(subject == "Microsoft Corporation", runtime.Verify("owned-certificate.exe"));
                    Assert.AreEqual(IntPtr.Zero, certificate.Handle);
                }
            }
        }

        [TestMethod]
        public void NativeRuntimeInstallerBuildsSilentArgumentsButOnlyRunsAnOwnedExitHelper()
        {
            using (var fixture = new UpdatesNativeFixture())
            {
                var runtime = new WebViewRuntimePrerequisite();
                WebViewRuntimePrerequisite.StartProcess = start => null;
                Assert.ThrowsException<InvalidOperationException>(() => runtime.Install("owned-bootstrapper.exe"));
                WebViewRuntimePrerequisite.StartProcess = start =>
                {
                    Assert.AreEqual("owned-bootstrapper.exe", start.FileName); Assert.AreEqual("/silent /install", start.Arguments);
                    Assert.IsFalse(start.UseShellExecute); Assert.IsTrue(start.CreateNoWindow); return UpdatesNativeFixture.ExitHelper(7);
                };
                Assert.AreEqual(7, runtime.Install("owned-bootstrapper.exe"));
            }
        }
    }
}
