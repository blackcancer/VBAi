using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;

namespace VBAi.Tests.Unit
{
    /// <summary>Verifies independent shutdown attempts, original failures and no implicit resource retry.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class AddInShutdownCleanupTests
    {
        /// <summary>A failed owned resource still releases its managed field and permits later cleanup.</summary>
        [TestMethod, DataRow(false, false), DataRow(false, true), DataRow(true, false), DataRow(true, true)]
        public void EachResourceIsAttemptedOnceAndClearedEvenWhenItFails(bool fail, bool ignore)
        {
            var cleanup = new AddInShutdownCleanup(null);
            object resource = new object(); var owned = resource; int calls = 0;
            var failure = new IOException("owned cleanup failure");
            Action<object> dispose = item => { Assert.AreSame(owned, item); calls++; if (fail) throw failure; };
            cleanup.Resource(ref resource, "owned", dispose, "Dispose", error => ignore);
            cleanup.Resource(ref resource, "owned", dispose);
            Assert.IsNull(resource); Assert.AreEqual(1, calls);
            bool later = false; cleanup.Step("later", () => later = true);
            Assert.IsTrue(later);
            if (fail && !ignore) Assert.AreSame(failure, Assert.ThrowsException<IOException>(cleanup.ThrowIfFailed));
            else cleanup.ThrowIfFailed();
        }

        /// <summary>Every original error survives aggregation, in execution order, after later steps finish.</summary>
        [TestMethod, DataRow(0), DataRow(1), DataRow(2), DataRow(3)]
        public void FailureCollectorPreservesExactExceptionInstancesAndAttemptOrder(int count)
        {
            var cleanup = new AddInShutdownCleanup(null);
            var errors = new List<Exception>(); int finished = 0;
            for (int index = 0; index < count; index++)
            {
                var error = new IOException("failure " + index); errors.Add(error);
                cleanup.Step("failed", () => { throw error; });
            }
            cleanup.Step("finished", () => finished++); Assert.AreEqual(1, finished);
            if (count == 0) cleanup.ThrowIfFailed();
            else if (count == 1) Assert.AreSame(errors[0], Assert.ThrowsException<IOException>(cleanup.ThrowIfFailed));
            else CollectionAssert.AreEqual(errors, Assert.ThrowsException<AggregateException>(cleanup.ThrowIfFailed).InnerExceptions);
        }

        /// <summary>An exception in a native catch policy cannot erase the native failure or skip the next resource.</summary>
        [TestMethod]
        public void FailingCatchPolicyRetainsBothFailuresAndContinuesCleanup()
        {
            var cleanup = new AddInShutdownCleanup(null);
            var primary = new IOException("primary"); var policy = new InvalidOperationException("policy");
            object resource = new object(); bool later = false;
            cleanup.Resource(ref resource, "owned", item => { throw primary; }, "Close", error => { throw policy; });
            cleanup.Step("later", () => later = true);
            Assert.IsNull(resource); Assert.IsTrue(later);
            CollectionAssert.AreEqual(new Exception[] { primary, policy }, Assert.ThrowsException<AggregateException>(cleanup.ThrowIfFailed).InnerExceptions);
        }
    }
}
