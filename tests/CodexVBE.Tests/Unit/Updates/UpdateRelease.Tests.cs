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
    }
}
