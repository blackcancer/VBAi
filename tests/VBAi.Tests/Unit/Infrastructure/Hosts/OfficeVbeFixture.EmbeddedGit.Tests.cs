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
    }
}
