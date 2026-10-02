using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class OfficeVbeFixtureEmbeddedGitTests
    {
        [TestMethod]
        public void AliasedProjectLeaseIsReleasedOnceAndNullLeasesAreIgnored()
        {
            var project = new object(); var collection = new object(); var released = new List<object>();
            OfficeVbeFixture.ReleaseWordGitMenuReferences(new[] { project, project, null, collection }, released.Add, null);
            CollectionAssert.AreEqual(new[] { project, collection }, released);
        }

        [TestMethod]
        public void CleanupFinishesAllDistinctLeasesAndPreservesPrimaryBeforeReleaseErrors()
        {
            var project = new object(); var collection = new object(); var released = new List<object>();
            var primary = new InvalidOperationException("Native menu failed");
            var cleanup = new InvalidOperationException("Project release failed");
            var observed = Assert.ThrowsException<AggregateException>(() =>
                OfficeVbeFixture.ReleaseWordGitMenuReferences(new[] { project, project, collection }, lease => {
                    released.Add(lease); if (ReferenceEquals(lease, project)) throw cleanup;
                }, primary));
            CollectionAssert.AreEqual(new[] { project, collection }, released);
            Assert.AreEqual(2, observed.InnerExceptions.Count);
            Assert.AreSame(primary, observed.InnerExceptions[0]); Assert.AreSame(cleanup, observed.InnerExceptions[1]);
        }

        [TestMethod]
        public void CleanupFailureWithoutPrimaryRemainsObservableAndSuccessDoesNotRewrapPrimary()
        {
            var error = new InvalidOperationException("Release failed");
            var observed = Assert.ThrowsException<AggregateException>(() =>
                OfficeVbeFixture.ReleaseWordGitMenuReferences(new[] { new object() }, lease => { throw error; }, null));
            Assert.AreEqual(1, observed.InnerExceptions.Count); Assert.AreSame(error, observed.InnerExceptions[0]);
            // The caller's catch retains the original exception when cleanup succeeds.
            OfficeVbeFixture.ReleaseWordGitMenuReferences(new[] { new object() }, lease => { }, error);
        }

        [TestMethod]
        public void OwnerBridgeReferenceIdentityPreservesGuidSyntaxAndSortsCompleteVersions()
        {
            var lower = Reference("{000204ef-0000-0000-c000-000000000046}", 4, 2);
            var higher = Reference("{00020905-0000-0000-C000-000000000046}", 8, 7);
            Assert.AreEqual("{000204EF-0000-0000-C000-000000000046}:4:2;{00020905-0000-0000-C000-000000000046}:8:7",
                OfficeVbeFixture.WordGitReferenceIdentity(new object[] { higher, lower }));
        }

        [TestMethod]
        public void MissingEmptyAndOversizedOwnerReferenceInventoriesAreRefused()
        {
            foreach (var rows in new[] { null, new object[0], new object[257] })
                Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.WordGitReferenceIdentity(rows));
        }

        [TestMethod]
        [DataRow("")][DataRow("invalid")][DataRow(" {000204EF-0000-0000-C000-000000000046}")]
        public void MalformedOwnerReferenceIdentityIsRefused(string guid)
        {
            Assert.ThrowsException<InvalidOperationException>(() =>
                OfficeVbeFixture.WordGitReferenceIdentity(new object[] { Reference(guid, 4, 2) }));
        }

        [TestMethod]
        [DataRow(-1, 2)][DataRow(65536, 2)][DataRow(4, -1)][DataRow(4, 65536)]
        public void OutOfRangeOwnerReferenceVersionIsRefused(int major, int minor)
        {
            Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.WordGitReferenceIdentity(new object[] {
                Reference("{000204EF-0000-0000-C000-000000000046}", major, minor) }));
        }

        [TestMethod]
        public void BrokenOwnerReferenceIsRefused()
        {
            var row = Reference("{000204EF-0000-0000-C000-000000000046}", 4, 2); row["IsBroken"] = true;
            Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.WordGitReferenceIdentity(new object[] { row }));
        }

        [TestMethod]
        public void NativeTemplateReferenceRetainsEmptyGuidAndZeroVersion()
        {
            var template = TemplateReference();
            Assert.AreEqual(":0:0;{000204EF-0000-0000-C000-000000000046}:4:2",
                OfficeVbeFixture.WordGitReferenceIdentity(new object[] {
                    Reference("{000204EF-0000-0000-C000-000000000046}", 4, 2), template }));
        }

        [TestMethod]
        [DataRow("BuiltIn", true)][DataRow("Name", " ")][DataRow("FullPath", "Normal")]
        [DataRow("Name", "OtherTemplate")]
        [DataRow("Major", 1)][DataRow("Minor", 1)][DataRow("Guid", null)][DataRow("IsBroken", true)]
        public void EmptyGuidWithoutExactTemplateMetadataIsRefused(string field, object value)
        {
            var template = TemplateReference(); template[field] = value;
            Assert.ThrowsException<InvalidOperationException>(() =>
                OfficeVbeFixture.WordGitReferenceIdentity(new object[] { template }));
        }

        private static Dictionary<string, object> TemplateReference()
        {
            var row = Reference("", 0, 0);
            row["BuiltIn"] = false; row["Name"] = "Normal"; row["FullPath"] = @"C:\Synthetic\Normal";
            return row;
        }

        private static Dictionary<string, object> Reference(string guid, int major, int minor)
            => new Dictionary<string, object> { ["Guid"] = guid, ["Major"] = major, ["Minor"] = minor, ["IsBroken"] = false };
    }
}
