using System;
using System.IO;
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
    }
}
