using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class UpdateReleaseTests
    {
        [TestMethod]
        public void VersionsCompareSemanticallyAndNeverDowngradeOrConfusePreviewNumbers()
        {
            var versions = new[] { "1.0.0-alpha", "1.0.0-alpha.1", "1.0.0-alpha.beta", "1.0.0-beta", "1.0.0-beta.2", "1.0.0-beta.11", "1.0.0-rc.1", "1.0.0", "1.2.0", "1.10.0", "2.0.0" };
            for (int i = 1; i < versions.Length; i++) Assert.IsTrue(UpdateVersion.Parse(versions[i]).CompareTo(UpdateVersion.Parse(versions[i - 1])) > 0);
            Assert.AreEqual(0, UpdateVersion.Parse("v1.2.3+build").CompareTo(UpdateVersion.Parse("1.2.3.0")));
            Assert.IsTrue(UpdateVersion.Parse("1.2.3-beta.99999999999999999999").CompareTo(UpdateVersion.Parse("1.2.3-beta.100")) > 0);
            foreach (string invalid in new[] { "", null, "main", "v1.2", "01.2.3", "1.0.0-beta.01", "1.0.0/unsafe", "999999999999.1.0" }) Assert.IsNull(UpdateVersion.Parse(invalid), invalid);
            var release = new UpdateRelease { assets = new[] { new UpdateAsset { name = "source.zip" }, new UpdateAsset { name = "VBAi-Setup-win-x64.exe" }, new UpdateAsset { name = "VBAi-Setup-win-x64.msi" } } };
            Assert.AreEqual("VBAi-Setup-win-x64.msi", release.Installer.name);
        }
        [TestMethod]
        public void VersionAndAssetContractsOrderAllPreviewKindsAndRejectInvalidDigests()
        {
            var ordered = new[] { "1.0.0-0", "1.0.0-0.0", "1.0.0-1", "1.0.0-2", "1.0.0-10", "1.0.0-0a", "1.0.0-a", "1.0.0-a.0", "1.0.0-a.a", "1.0.0-b", "1.0.0" };
            for (int i = 0; i < ordered.Length; i++)
            {
                var left = UpdateVersion.Parse(ordered[i]); Assert.IsNotNull(left); Assert.AreEqual(0, left.CompareTo(UpdateVersion.Parse(ordered[i]))); Assert.AreEqual(1, left.CompareTo(null));
                for (int j = 0; j < ordered.Length; j++) Assert.AreEqual(Math.Sign(i.CompareTo(j)), Math.Sign(left.CompareTo(UpdateVersion.Parse(ordered[j]))), ordered[i] + " vs " + ordered[j]);
            }
            Assert.IsNull(UpdateVersion.Parse("1.0.0+" + new string('x', 123)));
            Assert.IsNull(new UpdateRelease().Installer); Assert.IsNull(new UpdateRelease { assets = new UpdateAsset[0] }.Installer);
            var executable = new UpdateAsset { name = "VBAi-Setup-win-x64.exe" }; Assert.AreSame(executable, new UpdateRelease { assets = new[] { executable } }.Installer);
            foreach (string digest in new[] { null, "", "sha256:bad", "sha512:" + new string('a', 64) }) Assert.IsNull(new UpdateAsset { digest = digest }.Hash);
            Assert.AreEqual(new string('a', 64), new UpdateAsset { digest = "sha256:" + new string('A', 64) }.Hash);
        }
    }
}
