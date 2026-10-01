using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Preserves a returned mutation's original failure while collecting a single read-only follow-up.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class OfficeMetadataMutationEvidenceTests
    {
        [TestMethod]
        public void ObservationRunsAfterSuccessfulMutationExactlyOnce()
        {
            var order = new List<string>();
            OfficeMetadataMutationEvidence.Run(() => order.Add("mutation"), () => order.Add("read-only"));
            CollectionAssert.AreEqual(new[] { "mutation", "read-only" }, order);
        }

        [TestMethod]
        public void OriginalMutationFailureSurvivesSuccessfulReadbackWithoutReplay()
        {
            int mutations = 0, reads = 0; var original = new InvalidOperationException("original native mutation failure");
            var actual = Assert.ThrowsException<InvalidOperationException>(() => OfficeMetadataMutationEvidence.Run(
                () => { mutations++; throw original; }, () => reads++));
            Assert.AreSame(original, actual); Assert.AreEqual(1, mutations); Assert.AreEqual(1, reads);
        }

        [TestMethod]
        public void IndependentReadFailureOrUncertainDeliveryRefusalDoesNotMaskOriginalMutation()
        {
            int mutations = 0, reads = 0;
            var original = new InvalidOperationException("original"); var refusal = new InvalidOperationException("uncertain delivery: read-only follow-up refused");
            var actual = Assert.ThrowsException<AggregateException>(() => OfficeMetadataMutationEvidence.Run(
                () => { mutations++; throw original; }, () => { reads++; throw refusal; }));
            Assert.AreSame(original, actual.InnerExceptions[0]); Assert.AreSame(refusal, actual.InnerExceptions[1]);
            Assert.AreEqual(1, mutations); Assert.AreEqual(1, reads);
        }
    }
}
