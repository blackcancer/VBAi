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

    public sealed partial class GitReviewTests
    {
        [TestMethod]
        public void PartialSnapshotsKeepFormResourcesAndReferencesConsistent()
        {
            VbaGitSnapshot make(byte resource, string reference)
            {
                return new VbaGitSnapshot(new VbaGitManifest { References = reference, Components = new[] { new VbaGitComponent { Name = "Form1", Type = 3, HasResources = true }, new VbaGitComponent { Name = "Module1", Type = 1 } } }, new Dictionary<string, byte[]> { { "Form1.frm", VbaGitSnapshot.Utf8.GetBytes("Attribute VB_Name = \"Form1\"\n") }, { "Form1.frx", new[] { resource } }, { "Module1.bas", VbaGitSnapshot.Utf8.GetBytes("Attribute VB_Name = \"Module1\"\n") } });
            }

            var old = make(1, "a");
            var current = make(2, "b");
            var selected = VbaGitSnapshot.Select(old, current, new[] { "Form1" });
            Assert.AreEqual((byte)2, selected.Files["Form1.frx"][0]);
            Assert.AreEqual("a", selected.Manifest.References);
            Assert.AreEqual("b", VbaGitSnapshot.Select(old, current, new string[0], true).Manifest.References);
            Assert.ThrowsException<ArgumentException>(() => VbaGitSnapshot.Select(old, current, new[] { "Unknown" }));
        }
    }
}
