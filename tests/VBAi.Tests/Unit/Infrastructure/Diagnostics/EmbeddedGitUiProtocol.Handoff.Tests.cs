using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class EmbeddedGitUiHandoffTests
    {
        private const string Session = "11111111111111111111111111111111", Operation = "22222222222222222222222222222222";
        private static string State(string phase, string action = "checkpoint_restore", string operation = Operation, string session = Session)
            => "VBAi.GitSession/1/" + session + "/" + operation + "/" + action + "/" + phase;

        [DataTestMethod, DataRow("AwaitingModalReturn", false, false), DataRow("Executing", false, false)]
        [DataRow("Refused", false, false), DataRow("Succeeded", true, true), DataRow("Failed", true, false)]
        public void OnlyTheCorrelatedBusinessTerminalAllowsThePendingOperationToClose(string phase, bool terminal, bool success)
        {
            var protocol = new EmbeddedGitUiProtocol(_ => { }); protocol.EmitOnce("checkpointRestore", () => { });
            string observed = null; bool actual;
            bool known = EmbeddedGitUiProtocol.HandoffTerminal(Session, "none", "checkpoint_restore", State(phase), ref observed, out actual);
            Assert.AreEqual(terminal, known); Assert.AreEqual(success, actual); Assert.AreEqual(Operation, observed);
            Assert.IsFalse(protocol.CanClose);
            if (known) { protocol.Terminal("checkpointRestore", true, actual); Assert.IsTrue(protocol.CanClose); }
        }

        [DataTestMethod, DataRow("session"), DataRow("operation"), DataRow("action"), DataRow("phase"), DataRow("malformed")]
        public void DifferentOrInventedHandoffEvidenceNeverClearsPendingDelivery(string kind)
        {
            string observed = kind == "operation" ? new string('3', 32) : null;
            string description = kind == "malformed" ? "title only" : State(kind == "phase" ? "Closed" : "Succeeded",
                kind == "action" ? "pull" : "checkpoint_restore", session: kind == "session" ? new string('3', 32) : Session);
            bool success;
            Assert.ThrowsException<InvalidOperationException>(() => EmbeddedGitUiProtocol.HandoffTerminal(Session, "none", "checkpoint_restore", description, ref observed, out success));
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void PriorOperationOrIdleWindowIsNotANewTerminal(bool idle)
        {
            string observed = null; bool success;
            string description = idle ? State("Idle", "none", "none") : State("Succeeded");
            Assert.IsFalse(EmbeddedGitUiProtocol.HandoffTerminal(Session, Operation, "checkpoint_restore", description, ref observed, out success));
            Assert.IsNull(observed);
        }
    }
}
