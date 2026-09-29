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
using VBAi;
using VBAi.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class UpdateStateTests
    {
        [TestMethod]
        public void PreferencesThrottleDailyChecksAndDeploymentMarkersDoNotEnableDeveloperInstalls()
        {
            using (var scope = new UpdateScope())
            {
                Assert.IsFalse(UpdateInstallation.IsManaged(UpdateState.InstallationDirectory)); scope.Managed(); Assert.IsTrue(UpdateInstallation.IsManaged(UpdateState.InstallationDirectory));
                var p = UpdateState.Load(); Assert.IsTrue(p.CheckAutomatically); Assert.IsTrue(p.DownloadAutomatically); Assert.IsFalse(p.InstallAutomatically);
                DateTime now = DateTime.UtcNow; Assert.IsTrue(UpdateState.Due(p, now)); p.LastCheckUtc = now; Assert.IsFalse(UpdateState.Due(p, now));
                Assert.IsTrue(UpdateState.Due(p, now.AddHours(24))); Assert.IsTrue(UpdateState.Due(p, now.AddHours(-1)));
                p.LastCheckUtc = DateTime.MinValue; p.LastAttemptUtc = now; Assert.IsFalse(UpdateState.Due(p, now.AddMinutes(59))); Assert.IsTrue(UpdateState.Due(p, now.AddHours(1)));
                p.CheckAutomatically = false; p.IncludePrereleases = true; UpdateState.Save(p); Assert.IsTrue(UpdateState.Load().IncludePrereleases); Assert.IsFalse(UpdateState.Due(p, now.AddDays(2)));
                UpdateState.RecordHost(); Assert.AreEqual(1, Directory.GetFiles(Path.Combine(scope.Root, "hosts")).Length);
                File.WriteAllText(Path.Combine(UpdateState.InstallationDirectory, "vbai-installation.json"), "{\"Product\":\"Other\"}"); Assert.IsFalse(UpdateInstallation.IsManaged(UpdateState.InstallationDirectory));
            }
        }
        [TestMethod]
        public void CachedReleaseRejectsEveryUntrustedOrIneligibleOwnedPayloadAndBoundsOnlyItsCopy()
        {
            using (var scope = new UpdateScope())
            {
                string path = Path.Combine(scope.Root, "latest-release.json");
                Assert.IsNull(UpdateState.CachedRelease());
                foreach (string payload in new[] { "null", "broken", new string('x', 262145) })
                { File.WriteAllText(path, payload); Assert.IsNull(UpdateState.CachedRelease()); }
                var preferences = new UpdatePreferences(); UpdateState.Save(preferences);
                foreach (string kind in new[] { "invalid", "draft", "older", "prerelease", "preview", "stable" })
                {
                    var release = new UpdateRelease { tag_name = kind == "invalid" ? "invalid" : kind == "older" ? "0.0.0" : kind == "preview" ? "999.0.0-alpha" : "999.0.0", draft = kind == "draft", prerelease = kind == "prerelease", body = null };
                    UpdateState.CacheRelease(release);
                    Assert.AreEqual(kind == "stable", UpdateState.CachedRelease() != null, kind);
                    if (kind == "prerelease" || kind == "preview")
                    { preferences.IncludePrereleases = true; UpdateState.Save(preferences); Assert.IsNotNull(UpdateState.CachedRelease()); preferences.IncludePrereleases = false; UpdateState.Save(preferences); }
                }
                UpdateState.CacheRelease(null); Assert.AreEqual("null", File.ReadAllText(path));
                var original = new UpdateRelease { tag_name = "999.0.0", body = new string('n', 64001), assets = new[] { UpdateFeedTests.Asset("owned") } };
                UpdateState.CacheRelease(original); Assert.AreEqual(64001, original.body.Length);
                var cached = UpdateState.CachedRelease(); Assert.AreEqual(64000, cached.body.Length); Assert.AreEqual(original.assets[0].digest, cached.assets[0].digest);
            }
        }
        [TestMethod]
        public void OwnedPreferencesMarkersAndHostLeasesCoverNullPayloadsAndExplicitRoots()
        {
            using (var scope = new UpdateScope())
            {
                string path = Path.Combine(scope.Root, "preferences.json");
                foreach (string payload in new[] { "null", "malformed" })
                { File.WriteAllText(path, payload); Assert.IsTrue(UpdateState.Load(scope.Root).DownloadAutomatically); }
                UpdateState.Save(new UpdatePreferences { CheckAutomatically = false }, scope.Root); Assert.IsFalse(UpdateState.Load(scope.Root).CheckAutomatically);
                var now = DateTime.UtcNow; Assert.IsTrue(UpdateState.Due(new UpdatePreferences { LastAttemptUtc = now.AddHours(1) }, now));
                UpdateState.RecordHost(scope.Root);
                var lease = new JavaScriptSerializer().Deserialize<UpdateHostLease>(File.ReadAllText(Directory.GetFiles(Path.Combine(scope.Root, "hosts")).Single()));
                Assert.AreEqual(Process.GetCurrentProcess().Id, lease.Pid); Assert.AreEqual(UpdateState.InstallationDirectory, lease.InstallationDirectory); Assert.IsTrue(lease.StartTimeUtcTicks > 0);
                string markerPath = Path.Combine(UpdateState.InstallationDirectory, "vbai-installation.json");
                foreach (string kind in new[] { "null", "product", "architecture", "protocol", "id" })
                {
                    var marker = new UpdateInstallation { Product = kind == "product" ? "Other" : "VBAi", Architecture = kind == "architecture" ? "win-x86" : "win-x64", UpdateProtocol = kind == "protocol" ? 2 : 1, InstallationId = kind == "id" ? "invalid" : Guid.NewGuid().ToString() };
                    File.WriteAllText(markerPath, new JavaScriptSerializer().Serialize(kind == "null" ? null : marker)); Assert.IsFalse(UpdateInstallation.IsManaged(UpdateState.InstallationDirectory));
                }
                UpdateState.InstallationDirectoryOverride = null;
                Assert.AreEqual(Path.GetDirectoryName(typeof(UpdateState).Assembly.Location), UpdateState.InstallationDirectory);
            }
        }
        [TestMethod]
        public void ProductAssemblyMetadataFallsBackToOwnedDynamicAssemblyVersion()
        {
            var read = UpdateState.ReadProductAssembly;
            try
            {
                foreach (string information in new[] { null, "1.2.3+owned", "null-attribute" })
                {
                    var assembly = AppDomain.CurrentDomain.DefineDynamicAssembly(new System.Reflection.AssemblyName("OwnedUpdateVersion" + Guid.NewGuid().ToString("N")) { Version = new Version(8, 7, 6, 5) }, System.Reflection.Emit.AssemblyBuilderAccess.Run);
                    if (information != null)
                        assembly.SetCustomAttribute(new System.Reflection.Emit.CustomAttributeBuilder(typeof(System.Reflection.AssemblyInformationalVersionAttribute).GetConstructor(new[] { typeof(string) }), new object[] { information == "null-attribute" ? null : information }));
                    UpdateState.ReadProductAssembly = () => assembly;
                    Assert.AreEqual(information == "1.2.3+owned" ? information : "8.7.6.5", UpdateState.ProductVersion);
                }
            }
            finally { UpdateState.ReadProductAssembly = read; }
        }
    }
}
