namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.IO.Pipes;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Web.Script.Serialization;
    using System.Windows.Forms;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>Vérifie le transport named-pipe du pont et la validation des commandes Immediate.</summary>
    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class BridgeServerTests
    {
        /// <summary>Vérifie les connexions successives, les erreurs de requête et la reprise après une erreur.</summary>
        [TestMethod]
        [STATestMethod]
        public void PipeProcessesSuccessValidationAndMalformedJsonAcrossConnections()
        {
            using (var dispatcher = new Control())
            {
                var handle = dispatcher.Handle;
                var vbe = new FakeVbe();
                vbe.VBProjects.Add(new FakeProject { Name = "Disposable", FileName = @"C:\Temp\Disposable.xlsm" });
                int id = Guid.NewGuid().GetHashCode() & int.MaxValue;
                using (var server = new BridgeServer(dispatcher, new VbeSession(vbe), id))
                {
                    server.Start();
                    var status = SendWithMessagePump(id, "{\"Command\":\"status\"}");
                    Assert.AreEqual(true, status["Ok"]);
                    Assert.AreEqual(true, ((IDictionary<string, object>)status["Data"])["Connected"]);
                    var projects = SendWithMessagePump(id, "{\"Command\":\"list_projects\"}");
                    Assert.AreEqual(true, projects["Ok"]);
                    var first = ((object[])projects["Data"])[0] as IDictionary<string, object>;
                    Assert.IsNotNull(first);
                    Assert.AreEqual("Disposable", first["Name"]);
                    var missing = SendWithMessagePump(id, "{}");
                    Assert.AreEqual(false, missing["Ok"]);
                    StringAssert.Contains((string)missing["Error"], "command is required");
                    var malformed = SendWithMessagePump(id, "{invalid json}");
                    Assert.AreEqual(false, malformed["Ok"]);
                    Assert.IsFalse(string.IsNullOrWhiteSpace((string)malformed["Error"]));
                    var unknown = SendWithMessagePump(id, "{\"Command\":\"unknown_command\"}");
                    Assert.AreEqual(false, unknown["Ok"]);
                    StringAssert.Contains((string)unknown["Error"], "Unknown command");
                    var hostError = SendWithMessagePump(id, "{\"Command\":\"list_modules\",\"Project\":\"Disposable\"}");
                    Assert.AreEqual(false, hostError["Ok"]);
                    Assert.IsFalse(string.IsNullOrWhiteSpace((string)hostError["Error"]));
                    var recovered = SendWithMessagePump(id, "{\"Command\":\"status\"}");
                    Assert.AreEqual(true, recovered["Ok"]);
                    var badTree = SendWithMessagePump(id, "{\"Command\":\"debug_item\",\"Pane\":\"locals\",\"Action\":\"expand\"}");
                    Assert.AreEqual(false, badTree["Ok"]);
                    Assert.IsFalse(string.IsNullOrWhiteSpace((string)badTree["Error"]));
                    var badDialog = SendWithMessagePump(id, "{\"Command\":\"respond_debug_dialog\",\"Diagnostic\":\" \"}");
                    Assert.AreEqual(false, badDialog["Ok"]);
                    Assert.IsFalse(string.IsNullOrWhiteSpace((string)badDialog["Error"]));
                }
            }
        }

        /// <summary>Refuse une exécution Immediate sans projet avant tout accès à une fenêtre native.</summary>
        [TestMethod]
        [STATestMethod]
        public void ImmediateCommandRejectsMissingProjectBeforeTouchingNativeWindow()
        {
            using (var dispatcher = new Control())
            {
                var handle = dispatcher.Handle;
                int id = Guid.NewGuid().GetHashCode() & int.MaxValue;
                using (var server = new BridgeServer(dispatcher, new VbeSession(new FakeVbe()), id))
                {
                    server.Start();
                    var result = SendWithMessagePump(id, "{\"Command\":\"immediate_execute\",\"ExpectedMode\":2,\"Text\":\"Debug.Print 1\"}");
                    Assert.AreEqual(false, result["Ok"]);
                    StringAssert.Contains((string)result["Error"], "Project and ExpectedMode");
                }
            }
        }

        /// <summary>Refuse un mode invalide, un projet absent et un mode modifié avant l’exécution.</summary>
        [TestMethod]
        [STATestMethod]
        public void ImmediateCommandRejectsInvalidModeMissingProjectAndChangedModeWithoutExecuting()
        {
            using (var dispatcher = new Control())
            {
                var handle = dispatcher.Handle;
                var vbe = new FakeVbe();
                vbe.VBProjects.Add(new FakeProject { Name = "Disposable", FileName = @"C:\Temp\Disposable.xlsm", Mode = 2 });
                int id = Guid.NewGuid().GetHashCode() & int.MaxValue;
                using (var server = new BridgeServer(dispatcher, new VbeSession(vbe), id))
                {
                    server.Start();
                    var invalidMode = SendWithMessagePump(id, "{\"Command\":\"immediate_execute\",\"Project\":\"Disposable\",\"ExpectedMode\":0}");
                    Assert.AreEqual(false, invalidMode["Ok"]);
                    StringAssert.Contains((string)invalidMode["Error"], "ExpectedMode (1 or 2)");
                    var missingProject = SendWithMessagePump(id, "{\"Command\":\"immediate_execute\",\"Project\":\"Absent\",\"ExpectedMode\":1}");
                    Assert.AreEqual(false, missingProject["Ok"]);
                    StringAssert.Contains((string)missingProject["Error"], "Project selector is absent or ambiguous");
                    var changedMode = SendWithMessagePump(id, "{\"Command\":\"immediate_execute\",\"Project\":\"Disposable\",\"ExpectedMode\":1,\"Text\":\"Debug.Print 1\"}");
                    Assert.AreEqual(false, changedMode["Ok"]);
                    StringAssert.Contains((string)changedMode["Error"], "Project mode changed");
                    var recovered = SendWithMessagePump(id, "{\"Command\":\"status\"}");
                    Assert.AreEqual(true, recovered["Ok"]);
                }
            }
        }
    }
}
