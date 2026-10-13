using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using VBAi.Desktop.Helper;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class OllamaOfficeObserverWorkerTests
    {
        [DataTestMethod][DataRow(true, "original", "original")]
        [DataRow(false, null, "original")][DataRow(false, "original", "replacement")]
        public void ExitedOrReplacedParentRefusesFurtherUiAdmission(bool exited, string expected, string actual)
            => Assert.ThrowsException<InvalidOperationException>(() => OllamaOfficeObserverWorker.RequireParentAlive(exited, expected, actual));

        [TestMethod]
        public void SameLiveParentRetainsAdmission()
            => OllamaOfficeObserverWorker.RequireParentAlive(false, "original", "original");

        [DataTestMethod][DataRow("StreamedWhileBusy")][DataRow("CancellationAcknowledged")]
        [DataRow("NextSendCompleted")][DataRow("NativeMarkerVisible")]
        public void EachOracleRecordsOnlyItsOwnObservedStage(string stage)
        {
            var receipt = new OfficeChatObserverReceipt();
            OllamaOfficeObserverWorker.SetPassed(receipt, stage);
            Assert.AreEqual(stage == "StreamedWhileBusy", receipt.StreamedWhileBusy);
            Assert.AreEqual(stage == "CancellationAcknowledged", receipt.CancellationAcknowledged);
            Assert.AreEqual(stage == "NextSendCompleted", receipt.NextSendCompleted);
            Assert.AreEqual(stage == "NativeMarkerVisible", receipt.NativeMarkerVisible);
            Assert.IsFalse(receipt.Completed); Assert.IsFalse(receipt.Settled);
        }

        [TestMethod]
        public void InventedOracleDoesNotCreateAPassOrAuthorizeShutdown()
        {
            var receipt = new OfficeChatObserverReceipt();
            Assert.ThrowsException<ArgumentException>(() => OllamaOfficeObserverWorker.SetPassed(receipt, "looksGood"));
            Assert.IsFalse(receipt.Completed); Assert.IsFalse(receipt.Settled); Assert.IsFalse(receipt.NativeMarkerVisible);
        }
    }
}
