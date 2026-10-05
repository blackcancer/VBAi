using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Verifies independent fixture project acquisitions and balanced failure cleanup without native hosts.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class ExcelVbeFixtureGitTests
    {
        /// <summary>The caller receives only its unique wrapper after both temporary acquisitions are released once.</summary>
        [TestMethod]
        public void UniqueHandoffBalancesIdentityAndSharedAcquisitionWithoutReleasingCallerOwnership()
        {
            object shared = new object(), unique = new object();
            var events = new List<string>();
            var result = ExcelVbeFixture.AcquireIndependentGitProject(
                () => { events.Add("shared"); return shared; },
                value => { Assert.AreSame(shared, value); events.Add("identity"); return new IntPtr(27); },
                value => { Assert.AreEqual(new IntPtr(27), value); events.Add("unique"); return unique; },
                value => { Assert.AreEqual(new IntPtr(27), value); events.Add("release identity"); },
                value => { Assert.AreSame(shared, value); events.Add("release shared"); });
            Assert.AreSame(unique, result);
            CollectionAssert.AreEqual(new[] { "shared", "identity", "unique", "release identity", "release shared" }, events);
        }

        /// <summary>An uncertain acquisition is attempted once and only successfully acquired resources are released.</summary>
        [DataTestMethod, DataRow(0), DataRow(1), DataRow(2)]
        public void AcquisitionFailurePreservesOriginalExceptionAndBalancesOnlyAcquiredReferences(int boundary)
        {
            object shared = new object(); var failure = new InvalidOperationException("acquire " + boundary);
            int borrowed = 0, identified = 0, acquired = 0, identityReleases = 0, sharedReleases = 0;
            var thrown = Assert.ThrowsException<InvalidOperationException>(() => ExcelVbeFixture.AcquireIndependentGitProject(
                () => { borrowed++; if (boundary == 0) throw failure; return shared; },
                value => { identified++; if (boundary == 1) throw failure; return new IntPtr(27); },
                value => { acquired++; throw failure; },
                value => identityReleases++, value => { Assert.AreSame(shared, value); sharedReleases++; }));
            Assert.AreSame(failure, thrown); Assert.AreEqual(1, borrowed);
            Assert.AreEqual(boundary == 0 ? 0 : 1, identified); Assert.AreEqual(boundary == 2 ? 1 : 0, acquired);
            Assert.AreEqual(boundary == 2 ? 1 : 0, identityReleases); Assert.AreEqual(boundary == 0 ? 0 : 1, sharedReleases);
        }

        /// <summary>A failed handoff retains all release errors and still disposes the unreturned unique wrapper exactly once.</summary>
        [DataTestMethod]
        [DataRow(true, false, false), DataRow(false, true, false), DataRow(true, true, false)]
        [DataRow(true, false, true), DataRow(false, true, true), DataRow(true, true, true)]
        public void CleanupFailureCannotLeakTheUnreturnedUniqueWrapperOrSuppressOtherReleaseErrors(bool identityFails, bool sharedFails, bool uniqueFails)
        {
            object shared = new object(), unique = new object(); var events = new List<string>();
            InvalidOperationException identityError = new InvalidOperationException("identity"), sharedError = new InvalidOperationException("shared"),
                uniqueError = new InvalidOperationException("unique");
            var expected = new List<Exception>();
            if (identityFails) expected.Add(identityError); if (sharedFails) expected.Add(sharedError); if (uniqueFails) expected.Add(uniqueError);
            Exception thrown = null;
            try
            {
                ExcelVbeFixture.AcquireIndependentGitProject(() => shared, value => new IntPtr(27), value => unique,
                    value => { events.Add("identity"); if (identityFails) throw identityError; },
                    value => { bool isShared = ReferenceEquals(value, shared); events.Add(isShared ? "shared" : "unique");
                        if (isShared && sharedFails) throw sharedError; if (!isShared && uniqueFails) throw uniqueError; });
                Assert.Fail("A cleanup failure must prevent the unique wrapper handoff");
            }
            catch (InvalidOperationException error) { thrown = error; }
            catch (AggregateException error) { thrown = error; }
            Assert.IsNotNull(thrown);
            var actual = thrown is AggregateException aggregate ? aggregate.InnerExceptions.ToArray() : new[] { thrown };
            CollectionAssert.AreEqual(expected, actual);
            CollectionAssert.AreEqual(new[] { "identity", "shared", "unique" }, events);
        }

        /// <summary>A primary acquisition error survives simultaneous identity and shared release failures.</summary>
        [DataTestMethod, DataRow(false), DataRow(true)]
        public void FailedAcquisitionAndCleanupRetainEveryOriginalErrorWithoutReacquisition(bool identified)
        {
            object shared = new object(); int attempts = 0; var primary = new InvalidOperationException("primary");
            InvalidOperationException identityError = new InvalidOperationException("identity release"), sharedError = new InvalidOperationException("shared release");
            var thrown = Assert.ThrowsException<AggregateException>(() => ExcelVbeFixture.AcquireIndependentGitProject(() => shared,
                value => { if (!identified) throw primary; return new IntPtr(27); },
                value => { attempts++; throw primary; }, value => { throw identityError; }, value => { throw sharedError; }));
            CollectionAssert.AreEqual(identified ? new Exception[] { primary, identityError, sharedError } : new Exception[] { primary, sharedError }, thrown.InnerExceptions.ToArray());
            Assert.AreEqual(identified ? 1 : 0, attempts);
        }
    }
}