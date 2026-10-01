using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Guards the prepared sequential native matrix against cleanup after nested uncertain failures.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class ExcelSequentialQualificationTests
    {
        [DataTestMethod]
        [DataRow("timeout")]
        [DataRow("io")]
        [DataRow("cancellation")]
        [DataRow("nested")]
        [DataRow("aggregate")]
        [DataRow("deep-aggregate")]
        public void UncertainDeliveryRemainsUncertainThroughNativeAndAsyncWrappers(string shape)
        {
            Exception error = shape == "timeout" ? (Exception)new TimeoutException() :
                shape == "io" ? new IOException() : shape == "cancellation" ? new OperationCanceledException() :
                shape == "nested" ? new InvalidOperationException("wrapper", new IOException()) :
                shape == "aggregate" ? (Exception)new AggregateException(new ArgumentException(), new TimeoutException()) :
                new AggregateException(new InvalidOperationException("wrapper", new AggregateException(new IOException())));
            Assert.IsTrue(NativeUserFormGitHubTests.HasUncertainSequentialDelivery(error));
        }

        [TestMethod]
        public void KnownSynchronousValidationFailuresAreNotInventedAsPendingDelivery()
        {
            Assert.IsFalse(NativeUserFormGitHubTests.HasUncertainSequentialDelivery(null));
            Assert.IsFalse(NativeUserFormGitHubTests.HasUncertainSequentialDelivery(new ArgumentException()));
            Assert.IsFalse(NativeUserFormGitHubTests.HasUncertainSequentialDelivery(new InvalidOperationException("validation", new FormatException())));
            Assert.IsFalse(NativeUserFormGitHubTests.HasUncertainSequentialDelivery(new AggregateException(new ArgumentException(), new FormatException())));
        }

        [TestMethod]
        public void DualComImportAndCaptureFailureWithPendingRecoveryRetainsUnknownLiveState()
        {
            var primary = new AggregateException(new COMException("mutation"), new COMException("capture"));
            int writes = 0, observations = 0, retained = 0;
            var actual = Assert.ThrowsException<AggregateException>(() => NativeUserFormGitHubTests.ExecuteGuardedImport(
                () => { writes++; throw primary; }, () => { observations++; return true; }, () => retained++));
            Assert.AreSame(primary, actual); Assert.AreEqual(1, writes); Assert.AreEqual(1, observations); Assert.AreEqual(1, retained);
        }

        [TestMethod]
        public void KnownValidationBeforeMutationWithNoRecoveryDoesNotInventUnknownImport()
        {
            var primary = new ArgumentException("preflight"); int retained = 0;
            var actual = Assert.ThrowsException<ArgumentException>(() => NativeUserFormGitHubTests.ExecuteGuardedImport(
                () => { throw primary; }, () => false, () => retained++));
            Assert.AreSame(primary, actual); Assert.AreEqual(0, retained);
        }

        [TestMethod]
        public void UnreadableRecoveryMarkerRetainsHostAndBothOriginalErrors()
        {
            var primary = new COMException("mutation"); var marker = new IOException("marker"); int retained = 0;
            var actual = Assert.ThrowsException<AggregateException>(() => NativeUserFormGitHubTests.ExecuteGuardedImport(
                () => { throw primary; }, () => { throw marker; }, () => retained++));
            CollectionAssert.AreEqual(new Exception[] { primary, marker }, actual.InnerExceptions);
            Assert.AreEqual(1, retained);
        }

        [TestMethod]
        public void RetentionFailureDoesNotReplaceMutationFailureOrRetryIt()
        {
            var primary = new COMException("mutation"); var retention = new IOException("retention"); int writes = 0;
            var actual = Assert.ThrowsException<AggregateException>(() => NativeUserFormGitHubTests.ExecuteGuardedImport(
                () => { writes++; throw primary; }, () => true, () => { throw retention; }));
            CollectionAssert.AreEqual(new Exception[] { primary, retention }, actual.InnerExceptions); Assert.AreEqual(1, writes);
        }

        [TestMethod]
        public void KnownSuccessfulImportNeedsNoRecoveryFailureObservation()
        {
            int writes = 0;
            NativeUserFormGitHubTests.ExecuteGuardedImport(() => writes++, () => { Assert.Fail(); return true; }, () => Assert.Fail());
            Assert.AreEqual(1, writes);
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void FinalEvidenceFailureKeepsPrimaryAndAlwaysRestoresContext(bool hasPrimary)
        {
            var primary = hasPrimary ? new COMException("primary") : null; var recording = new IOException("recording");
            int restored = 0;
            if (hasPrimary)
            {
                var actual = Assert.ThrowsException<AggregateException>(() => NativeUserFormGitHubTests.CompleteSequentialEvidence(
                    primary, () => { throw recording; }, () => restored++));
                CollectionAssert.AreEqual(new Exception[] { primary, recording }, actual.InnerExceptions);
            }
            else Assert.AreSame(recording, Assert.ThrowsException<IOException>(() => NativeUserFormGitHubTests.CompleteSequentialEvidence(
                null, () => { throw recording; }, () => restored++)));
            Assert.AreEqual(1, restored);
        }

        [TestMethod]
        public void ContextFailurePreservesOriginalScenarioAndEvidenceFailures()
        {
            var primary = new COMException("primary"); var recording = new IOException("recording"); var context = new InvalidOperationException("context");
            var actual = Assert.ThrowsException<AggregateException>(() => NativeUserFormGitHubTests.CompleteSequentialEvidence(
                primary, () => { throw recording; }, () => { throw context; }));
            CollectionAssert.AreEqual(new Exception[] { primary, recording, context }, actual.InnerExceptions);
        }
    }
}
