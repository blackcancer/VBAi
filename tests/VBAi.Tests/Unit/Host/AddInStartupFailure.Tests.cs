using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks independent failure reporting and rollback without losing the initiating exception.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class AddInStartupFailureTests
    {
        [DataTestMethod]
        [DataRow(false, false, false), DataRow(false, false, true)]
        [DataRow(false, true, false), DataRow(false, true, true)]
        [DataRow(true, false, false), DataRow(true, false, true)]
        [DataRow(true, true, false), DataRow(true, true, true)]
        public void EveryFailureHandlerRunsOnceAndRetainsTheOriginalOrderedFailures(bool logFails, bool reportFails, bool cleanupFails)
        {
            var original = new ApplicationException("connection");
            var log = new Exception("log"); var report = new Exception("report"); var cleanup = new Exception("cleanup");
            var calls = new List<string>();
            Exception observed = null;
            try
            {
                AddInStartupFailure.Rethrow(original,
                    () => { calls.Add("log"); if (logFails) throw log; },
                    () => { calls.Add("report"); if (reportFails) throw report; },
                    () => { calls.Add("cleanup"); if (cleanupFails) throw cleanup; });
            }
            catch (Exception error) { observed = error; }
            CollectionAssert.AreEqual(new[] { "log", "report", "cleanup" }, calls);
            var expected = new[] { original, logFails ? log : null, reportFails ? report : null, cleanupFails ? cleanup : null }.Where(error => error != null).ToArray();
            if (expected.Length == 1) Assert.AreSame(original, observed);
            else CollectionAssert.AreEqual(expected, ((AggregateException)observed).InnerExceptions.ToArray());
        }

        [TestMethod]
        public void TheSameExceptionIsNotRepeatedByMultipleFailureHandlers()
        {
            var original = new Exception("same original"); int calls = 0;
            Action fail = () => { calls++; throw original; };
            Assert.AreSame(original, Assert.ThrowsException<Exception>(() => AddInStartupFailure.Rethrow(original, fail, fail, fail)));
            Assert.AreEqual(3, calls);
        }

        [TestMethod]
        public void MissingFailureOrHandlersAreRejectedBeforeAnyAction()
        {
            Action unexpected = () => Assert.Fail("Missing dependencies cannot start recovery.");
            Assert.ThrowsException<ArgumentNullException>(() => AddInStartupFailure.Rethrow(null, unexpected, unexpected, unexpected));
            Assert.ThrowsException<ArgumentNullException>(() => AddInStartupFailure.Rethrow(new Exception(), null, unexpected, unexpected));
            Assert.ThrowsException<ArgumentNullException>(() => AddInStartupFailure.Rethrow(new Exception(), unexpected, null, unexpected));
            Assert.ThrowsException<ArgumentNullException>(() => AddInStartupFailure.Rethrow(new Exception(), unexpected, unexpected, null));
        }
    }
}
