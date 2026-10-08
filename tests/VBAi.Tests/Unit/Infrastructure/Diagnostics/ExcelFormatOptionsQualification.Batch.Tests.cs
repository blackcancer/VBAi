using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;

namespace VBAi.Tests.Unit
{
    public sealed partial class ExcelFormatOptionsQualificationTests
    {
        [TestMethod]
        public void CompleteGroupedMatrixAppliesAndRestoresAllValuesWithTwoConfirmedCommits()
        {
            var probe = new Probe();
            var before = new Dictionary<string, object>(probe.Values);
            string category = probe.Category;
            probe.Create(batchChanges: true).Run();
            var batches = probe.Requests.Where(request => request.ContainsKey("OptionChanges")).ToArray();
            Assert.AreEqual(3, batches.Length, "Apply, expected stale refusal, then restoration.");
            Assert.AreEqual(9, ((System.Collections.ICollection)batches[0]["OptionChanges"]).Count);
            Assert.AreEqual(9, ((System.Collections.ICollection)batches[2]["OptionChanges"]).Count);
            CollectionAssert.AreEquivalent(before.ToArray(), probe.Values.ToArray());
            Assert.AreEqual(category, probe.Category);
            Assert.AreEqual(1, probe.Cleanup); Assert.AreEqual(0, probe.Preserved);
            Assert.IsTrue(probe.Phases.Contains("FormatBatchVerified"));
            Assert.IsTrue(probe.Phases.Contains("StaleVersionBatchVerifiedRefusal"));
            Assert.IsTrue(probe.Phases.Contains("BaselineRestored"));
        }

        [TestMethod]
        public void EmptySizeCatalogueIsClassifiedBeforeTheCompleteGroupedMutation()
        {
            var probe = new Probe { EmptySizes = true };
            probe.Create(batchChanges: true).Run();
            var first = probe.Requests.First(request => request.ContainsKey("OptionChanges"));
            Assert.AreEqual(8, ((System.Collections.ICollection)first["OptionChanges"]).Count);
            Assert.IsTrue(probe.Phases.Contains("EmptySizeVerifiedRefusal"));
            Assert.AreEqual(1, probe.Cleanup); Assert.AreEqual(0, probe.Preserved);
        }

        [TestMethod]
        public void EveryUncertainGroupedReplyStopsTheBankWithoutRestorationOrCleanup()
        {
            var successful = new Probe(); successful.Create(batchChanges: true).Run();
            int requests = successful.Requests.Count;
            foreach (int index in Enumerable.Range(1, requests))
                foreach (string fault in new[] { "Lost", "Timeout", "Null", "MissingOk", "NoData", "TopPending", "Pending", "Uncertain", "DeliveryUncertain", "VerificationPending" })
                {
                    var probe = new Probe { FaultAt = index, Fault = fault };
                    var runner = probe.Create(batchChanges: true);
                    Assert.IsNotNull(Failure(runner.Run), fault + " at " + index);
                    Assert.IsTrue(runner.HostRetained, fault + " at " + index);
                    Assert.AreEqual(index, probe.Requests.Count, "No dispatch after an uncertain reply.");
                    Assert.AreEqual(0, probe.Cleanup);
                }
        }

        [TestMethod]
        public void GroupedQualificationRejectsHistoricalOrPartialMutationSequences()
        {
            foreach (int scenario in Enumerable.Range(0, 4))
            {
                var probe = new Probe();
                Assert.ThrowsException<ArgumentException>(() => probe.Create(batchChanges: true,
                    marginOnly: scenario == 0, fontSizeOnly: scenario == 1,
                    historicalPalettePrefix: scenario == 2 || scenario == 3, historicalFullMatrix: scenario == 3));
                Assert.AreEqual(0, probe.Requests.Count);
            }
        }
    }
}
