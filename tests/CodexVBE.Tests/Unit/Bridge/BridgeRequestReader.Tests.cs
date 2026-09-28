namespace CodexVBE.Tests.Unit
{
    using System;
    using System.IO;
    using System.IO.Pipes;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Web.Script.Serialization;
    using System.Windows.Forms;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed class BridgeRequestBudgetTests
    {
        [TestMethod]
        public async Task ReaderAcceptsExactByteBudgetAndCrLfWithUnicode()
        {
            var bytes = Encoding.UTF8.GetBytes("é漢\r\n");
            using (var stream = new MemoryStream(bytes))
                Assert.AreEqual("é漢", await BridgeRequestReader.ReadAsync(stream, bytes.Length - 1, TimeSpan.FromSeconds(1)));
        }

        [TestMethod]
        public async Task ReaderRejectsOversizeBeforeNewlineAndRejectsIncompleteFrame()
        {
            using (var stream = new MemoryStream(Encoding.UTF8.GetBytes("12345")))
                await Assert.ThrowsExceptionAsync<IOException>(() => BridgeRequestReader.ReadAsync(stream, 4, TimeSpan.FromSeconds(1)));
            using (var stream = new MemoryStream(Encoding.UTF8.GetBytes("{}")))
                await Assert.ThrowsExceptionAsync<EndOfStreamException>(() => BridgeRequestReader.ReadAsync(stream, 4, TimeSpan.FromSeconds(1)));
        }

        [DataTestMethod]
        [DataRow("silent")]
        [DataRow("partial")]
        [DataRow("interrupted")]
        [DataRow("oversized")]
        [DataRow("malformed")]
        [DataRow("invalid-utf8")]
        public void HealthyClientResumesAfterIncompleteOrInvalidClient(string scenario)
        {
            int id = Guid.NewGuid().GetHashCode() & int.MaxValue;
            int executed = 0;
            using (var dispatcher = new Control())
            using (var server = new BridgeServer(dispatcher, null, id))
            {
                server.MaxRequestBytes = 256;
                server.RequestReadTimeout = TimeSpan.FromMilliseconds(250);
                server.Native.Capture = include => { Interlocked.Increment(ref executed); return new { Healthy = true }; };
                server.Start();
                using (var bad = Connect(id))
                {
                    byte[] bytes = scenario == "partial" || scenario == "interrupted" ? Encoding.UTF8.GetBytes("{\"Command\":") :
                        scenario == "oversized" ? Encoding.UTF8.GetBytes(new string('x', 257)) :
                        scenario == "malformed" ? Encoding.UTF8.GetBytes("{broken}\n") :
                        scenario == "invalid-utf8" ? new byte[] { 0xc3, 0x28, 10 } : new byte[0];
                    if (bytes.Length > 0) bad.Write(bytes, 0, bytes.Length);
                    if (scenario == "interrupted") bad.Dispose();
                    else
                    {
                        var end = Task.Run(() =>
                        {
                            using (var reader = new StreamReader(bad, Encoding.UTF8, false, 1024, true))
                                return reader.ReadLine();
                        });
                        Assert.IsTrue(end.Wait(4000), "The incomplete client retained the listener.");
                        if (scenario == "malformed")
                            Assert.IsFalse(new JavaScriptSerializer().Deserialize<Response>(end.Result).Ok);
                        else Assert.IsNull(end.Result, "Invalid/incomplete frames must not be dispatched.");
                    }
                    Assert.AreEqual(0, executed);
                    using (var good = Connect(id))
                    {
                        var request = Encoding.UTF8.GetBytes("{\"Command\":\"debug_windows\"}\n");
                        good.Write(request, 0, request.Length);
                        var reply = Task.Run(() =>
                        {
                            using (var reader = new StreamReader(good, Encoding.UTF8, false, 1024, true))
                                return reader.ReadLine();
                        });
                        Assert.IsTrue(reply.Wait(4000), "A healthy client could not resume after " + scenario);
                        Assert.IsTrue(new JavaScriptSerializer().Deserialize<Response>(reply.Result).Ok);
                        Assert.AreEqual(1, executed);
                    }
                }
            }
        }

        private static NamedPipeClientStream Connect(int id)
        {
            var client = new NamedPipeClientStream(".", "CodexVBE." + id, PipeDirection.InOut, PipeOptions.Asynchronous);
            try { client.Connect(4000); return client; }
            catch { client.Dispose(); throw; }
        }
    }
}
