using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Preserves strict operation observations across a re-shown Git modal.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class EmbeddedGitAutomationTests
    {
        [DataTestMethod]
        [DataRow("compare")]
        [DataRow("checkpointCreate")]
        [DataRow("checkpointRestore")]
        public void ReShownCheckpointViewWaitsForCompareWithoutQueryingTheUnmaterializedConnectionTab(string action)
        {
            bool compareEnabled = false;
            int observations = 0;
            Func<string, bool> read = id =>
            {
                observations++;
                if (id != "compare") throw new InvalidOperationException("Exact native leaf absent or ambiguous: connect");
                return compareEnabled;
            };
            string terminal = action == "checkpointCreate" ? "Operation complete: checkpoint_create" :
                action == "checkpointRestore" ? "VBA restored. Check and save the document." :
                "VBA matches the last synchronized state.";
            bool idle = EmbeddedGitAutomation.ObserveIdle(action, read);
            Assert.IsFalse(idle);
            Assert.IsFalse(EmbeddedGitAutomation.HasKnownTerminal(action, idle, true, "Earlier state", terminal));
            compareEnabled = true;
            idle = EmbeddedGitAutomation.ObserveIdle(action, read);
            Assert.IsTrue(idle);
            Assert.IsTrue(EmbeddedGitAutomation.HasKnownTerminal(action, idle, true, "Earlier state", terminal));
            Assert.AreEqual(2, observations);
        }

        [TestMethod]
        public void UnlinkedConnectionErrorRequiresTheActualEnabledConnectControlAndRemainsAFailure()
        {
            int calls = 0;
            bool idle = EmbeddedGitAutomation.ObserveIdle("connect", id => { calls++; return id == "connect"; });
            const string error = "Synthetic failure. Check the connection, account and Git state, then retry.";
            Assert.IsTrue(idle);
            Assert.AreEqual(2, calls);
            Assert.IsTrue(EmbeddedGitAutomation.HasKnownTerminal("connect", idle, true, "Earlier state", error));
            Assert.AreEqual(-1, EmbeddedGitAutomation.ClassifyTerminal("connect", error));
        }

        [TestMethod]
        public void BusyConnectionControlsCannotConfirmAnAlreadyVisibleTerminalLabel()
        {
            bool idle = EmbeddedGitAutomation.ObserveIdle("connect", id => false);
            Assert.IsFalse(idle);
            Assert.IsFalse(EmbeddedGitAutomation.HasKnownTerminal("connect", idle, true, "Earlier state",
                "VBA matches the last synchronized state."));
        }

        [DataTestMethod]
        [DataRow("connect")]
        [DataRow("compare")]
        [DataRow("checkpointCreate")]
        [DataRow("checkpointRestore")]
        public void EnabledCompareNeedsNoConnectionFallbackForAnyPreparedOperation(string action)
        {
            int calls = 0;
            Assert.IsTrue(EmbeddedGitAutomation.ObserveIdle(action, id =>
            {
                calls++;
                Assert.AreEqual("compare", id);
                return true;
            }));
            Assert.AreEqual(1, calls);
        }

        [DataTestMethod]
        [DataRow("connect")]
        [DataRow("compare")]
        [DataRow("checkpointCreate")]
        [DataRow("checkpointRestore")]
        public void MandatoryCompareIdentityRefusalIsPreservedWithoutAnotherControlLookup(string action)
        {
            var expected = new InvalidOperationException("Native owner generation changed.");
            int calls = 0;
            var actual = Assert.ThrowsException<InvalidOperationException>(() =>
                EmbeddedGitAutomation.ObserveIdle(action, id => { calls++; throw expected; }));
            Assert.AreSame(expected, actual);
            Assert.AreEqual(1, calls);
        }

        [TestMethod]
        public void RequiredConnectionFallbackStillRefusesMissingOrAmbiguousConnect()
        {
            var expected = new InvalidOperationException("Exact native leaf absent or ambiguous: connect");
            var actual = Assert.ThrowsException<InvalidOperationException>(() =>
                EmbeddedGitAutomation.ObserveIdle("connect", id =>
                {
                    if (id == "compare") return false;
                    throw expected;
                }));
            Assert.AreSame(expected, actual);
        }

        [TestMethod]
        public void MissingNativeReaderRefusesBeforeAnyObservation()
        {
            Assert.ThrowsException<ArgumentNullException>(() => EmbeddedGitAutomation.ObserveIdle("compare", null));
        }
    }
}
