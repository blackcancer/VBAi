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
    }
}
