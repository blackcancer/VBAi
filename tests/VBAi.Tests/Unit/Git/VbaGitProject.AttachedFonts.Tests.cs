using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbaGitProjectAttachedFontsTests
    {
        [TestMethod]
        public void DisabledAttachedObservationDoesNotInvokeIdentityGuard()
        {
            VbaGitProject.ObserveAttachedImport(null, null, () => Assert.Fail("Disabled hook must add no COM reads."));
        }
        [TestMethod]
        public void AttachedImportUsesOriginalBorrowedObjectAfterItsGuardExactlyOnce()
        {
            var original = new object(); var order = new List<string>(); int calls = 0;
            Action guard = () => order.Add("identity");
            VbaGitProject.ObserveAttachedImport((actual, validate) =>
            {
                calls++; Assert.AreSame(original, actual); Assert.AreSame(guard, validate); order.Add("observation");
            }, original, guard);
            Assert.AreEqual(1, calls); CollectionAssert.AreEqual(new[] { "identity", "observation" }, order.ToArray());
        }
        [TestMethod]
        public void AttachedGuardOrObservationFailureNeverReplaysBorrowedCallback()
        {
            foreach (bool guardFails in new[] { true, false })
            {
                var error = new InvalidOperationException("known failure"); int observations = 0, guards = 0;
                try
                {
                    VbaGitProject.ObserveAttachedImport((component, guard) => { observations++; throw error; }, new object(),
                        () => { guards++; if (guardFails) throw error; });
                    Assert.Fail("The original diagnostic failure must escape.");
                }
                catch (InvalidOperationException actual) { Assert.AreSame(error, actual); }
                Assert.AreEqual(1, guards); Assert.AreEqual(guardFails ? 0 : 1, observations);
            }
        }
    }
}
