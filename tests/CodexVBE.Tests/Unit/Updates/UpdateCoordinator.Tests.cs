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
    public sealed class UpdateCoordinatorTests
    {
        [TestMethod]
        public async Task AutomaticChecksRespectPreferencesAndScheduleOnlyManagedDeployments()
        {
            using (var scope = new UpdateScope())
            {
                scope.Managed(); var asset = UpdateFeedTests.Asset("fixture installer");
                string release = new JavaScriptSerializer().Serialize(new[] { new { tag_name = "v1.2.3", assets = new[] { asset } } });
                var handler = new UpdateFeedTests.Handler(release, "fixture installer"); int launched = 0;
                UpdateCoordinator.CreateFeed = () => new UpdateFeed(handler, ct => throw new AssertFailedException());
                UpdateCoordinator.LaunchWorker = background => { Assert.IsTrue(background); launched++; };
                var p = UpdateState.Load(); p.InstallAutomatically = true; UpdateState.Save(p);
                var result = await UpdateCoordinator.Check(true, null, CancellationToken.None);
                Assert.AreEqual("1.2.3", result.Version.Text); Assert.AreEqual(1, launched);
                Assert.IsNotNull(UpdateInstallJob.Load(scope.Root)); Assert.IsNull(await UpdateCoordinator.Check(true, null, CancellationToken.None)); Assert.AreEqual(1, launched);
                p = UpdateState.Load(); p.LastCheckUtc = DateTime.MinValue; p.CheckAutomatically = false; UpdateState.Save(p);
                Assert.IsNull(await UpdateCoordinator.Check(true, null, CancellationToken.None));
            }
        }
        [TestMethod]
        public void SchedulingRejectsDeveloperCheckoutsAndDuplicatePendingInstalls()
        {
            using (var scope = new UpdateScope())
            {
                var asset = UpdateFeedTests.Asset("fixture installer"); var release = new UpdateRelease { tag_name = "1.2.3", assets = new[] { asset } };
                string path = UpdatePaths.AssetPath(scope.Root, asset.Hash, asset.name);
                Assert.ThrowsException<InvalidOperationException>(() => UpdateCoordinator.Schedule(release, path, false));
                scope.Managed(); Directory.CreateDirectory(Path.GetDirectoryName(path)); File.WriteAllText(path, "fixture installer", new UTF8Encoding(false)); UpdateCoordinator.LaunchWorker = background => { };
                UpdateCoordinator.Schedule(release, path, false);
                Assert.ThrowsException<InvalidOperationException>(() => UpdateCoordinator.Schedule(release, path, false));
            }
        }
    }
}
