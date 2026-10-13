using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks package sharing, immutable snapshots and every pre-inspection refusal without Office.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class OfficeVbeFixtureWordControlPackageTests
    {
        [TestMethod]
        public void FrozenTemplateIsInspectedWhileWriteHandleIsHeldAndNoBytesAreChanged()
        {
            WithTemplate((path, original, hash) =>
            {
                int calls = 0;
                using (var writer = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.Read))
                    OfficeVbeFixture.InspectWordControlPackage(path, hash.ToLowerInvariant(), package =>
                    {
                        calls++;
                        var part = package.GetParts().Single(p => !PackUriHelper.IsRelationshipPartUri(p.Uri));
                        Assert.AreEqual("application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml", part.ContentType);
                        Assert.AreEqual(part.Uri, package.GetRelationships().Single().TargetUri);
                    });
                Assert.AreEqual(1, calls); CollectionAssert.AreEqual(original, File.ReadAllBytes(path));
            });
        }

        [TestMethod]
        public void SnapshotRemainsReadableIfTheOwnedTestFileChangesAfterInspectionStarts()
        {
            WithTemplate((path, original, hash) =>
                OfficeVbeFixture.InspectWordControlPackage(path, hash, package =>
                {
                    File.WriteAllBytes(path, new byte[] { 1, 2 });
                    Assert.AreEqual(1, package.GetParts().Count(p => !PackUriHelper.IsRelationshipPartUri(p.Uri)));
                    File.WriteAllBytes(path, original);
                }));
        }

        [DataTestMethod, DataRow("changed"), DataRow("denied-sharing"), DataRow("empty"), DataRow("large"), DataRow("missing"), DataRow("invalid-zip")]
        public void InvalidOrChangedPackagesNeverReachTheInspector(string fault)
        {
            WithTemplate((path, original, hash) =>
            {
                if (fault == "changed") File.WriteAllBytes(path, new byte[] { 7 });
                if (fault == "empty") File.WriteAllBytes(path, new byte[0]);
                if (fault == "large") using (var file = File.OpenWrite(path)) file.SetLength(2 * 1024 * 1024 + 1);
                if (fault == "missing") File.Delete(path);
                if (fault == "invalid-zip") { File.WriteAllBytes(path, new byte[] { 1, 2, 3 }); hash = OfficeVbeFixture.ReadWordLifecycleSourceHash(path); }
                int calls = 0;
                using (var deny = fault == "denied-sharing" ? new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None) : null)
                {
                    Exception error = null;
                    try { OfficeVbeFixture.InspectWordControlPackage(path, hash, package => calls++); } catch (Exception failure) { error = failure; }
                    Assert.IsNotNull(error, fault); Assert.AreEqual(0, calls);
                    if (fault == "changed") Assert.IsInstanceOfType(error, typeof(InvalidOperationException));
                    if (fault == "denied-sharing") Assert.IsInstanceOfType(error, typeof(IOException));
                }
            });
        }

        [DataTestMethod, DataRow("null-path"), DataRow("relative-path"), DataRow("null-hash"), DataRow("short-hash"), DataRow("invalid-hash"), DataRow("null-inspector")]
        public void InvalidInputsRefuseBeforeAnyFileRead(string fault)
        {
            int calls = 0;
            Assert.ThrowsException<ArgumentException>(() => OfficeVbeFixture.InspectWordControlPackage(
                fault == "null-path" ? null : fault == "relative-path" ? "relative.dotx" : @"C:\does-not-exist\control.dotx",
                fault == "null-hash" ? null : fault == "short-hash" ? "00" : new string(fault == "invalid-hash" ? 'Q' : '0', 64),
                fault == "null-inspector" ? (Action<Package>)null : package => calls++));
            Assert.AreEqual(0, calls);
        }

        [TestMethod]
        public void InspectorFailureIsNotRetriedAndLeavesTheOwnedFileUnchanged()
        {
            WithTemplate((path, original, hash) =>
            {
                var primary = new IOException("inspector"); int calls = 0;
                Assert.AreSame(primary, Assert.ThrowsException<IOException>(() => OfficeVbeFixture.InspectWordControlPackage(path, hash,
                    package => { calls++; throw primary; })));
                Assert.AreEqual(1, calls); CollectionAssert.AreEqual(original, File.ReadAllBytes(path));
            });
        }

        private static void WithTemplate(Action<string, byte[], string> test)
        {
            string root = Path.Combine(Path.GetTempPath(), "VBAi-ControlPackage-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root); string path = Path.Combine(root, "control.dotx");
            try
            {
                OfficeVbeFixture.WriteWordLifecycleTemplateSeed(path);
                test(path, File.ReadAllBytes(path), OfficeVbeFixture.ReadWordLifecycleSourceHash(path));
            }
            finally { File.Delete(path); Directory.Delete(root); }
        }
    }
}
