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

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class BridgeServerTests
    {
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

                    var hostError = SendWithMessagePump(id,
                        "{\"Command\":\"list_modules\",\"Project\":\"Disposable\"}");
                    Assert.AreEqual(false, hostError["Ok"]);
                    Assert.IsFalse(string.IsNullOrWhiteSpace((string)hostError["Error"]));

                    var recovered = SendWithMessagePump(id, "{\"Command\":\"status\"}");
                    Assert.AreEqual(true, recovered["Ok"]);

                    var badTree = SendWithMessagePump(id,
                        "{\"Command\":\"debug_item\",\"Pane\":\"locals\",\"Action\":\"expand\"}");
                    Assert.AreEqual(false, badTree["Ok"]);
                    Assert.IsFalse(string.IsNullOrWhiteSpace((string)badTree["Error"]));

                    var badDialog = SendWithMessagePump(id,
                        "{\"Command\":\"respond_debug_dialog\",\"Diagnostic\":\" \"}");
                    Assert.AreEqual(false, badDialog["Ok"]);
                    Assert.IsFalse(string.IsNullOrWhiteSpace((string)badDialog["Error"]));
                }
            }
        }

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
                    var result = SendWithMessagePump(id,
                        "{\"Command\":\"immediate_execute\",\"ExpectedMode\":2,\"Text\":\"Debug.Print 1\"}");
                    Assert.AreEqual(false, result["Ok"]);
                    StringAssert.Contains((string)result["Error"], "Project and ExpectedMode");
                }
            }
        }

        private static IDictionary<string, object> SendWithMessagePump(int processId, string request)
        {
            var pending = Task.Run(() => {
                using (var pipe = new NamedPipeClientStream(".", "CodexVBE." + processId, PipeDirection.InOut))
                {
                    pipe.Connect(5000);
                    using (var writer = new StreamWriter(pipe, new UTF8Encoding(false), 4096, true) { AutoFlush = true })
                    using (var reader = new StreamReader(pipe, new UTF8Encoding(false), false, 4096, true))
                    {
                        writer.WriteLine(request);
                        return (IDictionary<string, object>)new JavaScriptSerializer().DeserializeObject(reader.ReadLine());
                    }
                }
            });
            var deadline = DateTime.UtcNow.AddSeconds(10);
            while (!pending.IsCompleted && DateTime.UtcNow < deadline)
            {
                Application.DoEvents();
                Thread.Sleep(5);
            }
            Assert.IsTrue(pending.IsCompleted, "The VBE bridge did not complete a pipe request.");
            return pending.GetAwaiter().GetResult();
        }

        public sealed class FakeVbe
        {
            public List<FakeProject> VBProjects { get; } = new List<FakeProject>();
        }

        public sealed class FakeProject
        {
            public string Name { get; set; }
            public string FileName { get; set; }
            public int Mode { get; set; } = 2;
        }
    }
}
