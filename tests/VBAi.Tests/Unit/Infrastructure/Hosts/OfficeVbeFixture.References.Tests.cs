using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks acquired-reference accounting and uncertain-delivery retention without launching Office.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class OfficeVbeFixtureReferenceTests
    {
        [DataTestMethod]
        [DataRow(false, false)]
        [DataRow(true, false)]
        [DataRow(false, true)]
        [DataRow(true, true)]
        public void ReleaseBalancesEachAcquisitionOrRetainsEveryAliasDuringUncertainDelivery(bool pending, bool uncertain)
        {
            var fixture = Create();
            var containment = (OfficeCommandContainment)Field(fixture, "commandContainment");
            typeof(OfficeCommandContainment).GetProperty("Pending", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(containment, pending);
            typeof(OfficeCommandContainment).GetProperty("Uncertain", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(containment, uncertain);
            var alias = new object();
            var other = new object();
            var released = new List<object>();
            fixture.ReleaseAcquiredReference = released.Add;
            Invoke(fixture, alias); Invoke(fixture, other); Invoke(fixture, alias); Invoke(fixture, null);
            var retained = (IList)Field(fixture, "retainedDiagnosticReferences");
            object[] expected = { alias, other, alias };
            if (pending || uncertain)
            {
                Assert.AreEqual(0, released.Count);
                CollectionAssert.AreEqual(expected, retained);
            }
            else
            {
                CollectionAssert.AreEqual(expected, released);
                Assert.AreEqual(0, retained.Count);
            }
        }

        [TestMethod]
        public void DefaultBoundaryDoesNotReleaseManagedObjectsAndNullDoesNotInvokeIt()
        {
            var fixture = Create();
            var borrowed = new object();
            Invoke(fixture, borrowed);
            fixture.ReleaseAcquiredReference = _ => Assert.Fail("Null is not an acquisition.");
            Invoke(fixture, null);
            Assert.AreEqual(0, ((IList)Field(fixture, "retainedDiagnosticReferences")).Count);
        }

        [TestMethod]
        public void ReleaseFailureIsNotAutomaticallyRetried()
        {
            var fixture = Create();
            var reference = new object();
            var primary = new InvalidOperationException("release failed");
            int calls = 0;
            fixture.ReleaseAcquiredReference = value => { Assert.AreSame(reference, value); calls++; throw primary; };
            var wrapper = Assert.ThrowsException<TargetInvocationException>(() => Invoke(fixture, reference));
            Assert.AreSame(primary, wrapper.InnerException);
            Assert.AreEqual(1, calls);
        }

        private static OfficeVbeFixture Create() => (OfficeVbeFixture)Activator.CreateInstance(typeof(OfficeVbeFixture), true);
        private static object Field(object target, string name) => typeof(OfficeVbeFixture).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(target);
        private static void Invoke(OfficeVbeFixture target, object reference) => typeof(OfficeVbeFixture).GetMethod("Release", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(target, new[] { reference });
    }
}
