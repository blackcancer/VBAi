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
namespace CodexVBE.Tests.Unit
{
    [Microsoft.VisualStudio.TestTools.UnitTesting.TestClass]
    [Microsoft.VisualStudio.TestTools.UnitTesting.TestCategory("Unit")]
    public sealed class BridgeRequestReaderBoundaryTests
    {
        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public async System.Threading.Tasks.Task ReaderValidatesBudgetsAndReadsEmptyUtf8OrFragmentedFrames()
        {
            using(var stream=new System.IO.MemoryStream()) {
                await Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsExceptionAsync<System.ArgumentOutOfRangeException>(()=>BridgeRequestReader.ReadAsync(stream,0,System.TimeSpan.FromSeconds(1)));
                await Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsExceptionAsync<System.ArgumentOutOfRangeException>(()=>BridgeRequestReader.ReadAsync(stream,1,System.TimeSpan.Zero));
            }
            foreach(var text in new[]{"\n","text\n","\r\n"}) using(var stream=new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes(text)))
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(text.TrimEnd('\r','\n'),await BridgeRequestReader.ReadAsync(stream,10,System.TimeSpan.FromSeconds(1)));
            using(var stream=new System.IO.MemoryStream(new byte[]{0xc3,0x28,10})) {
                var error=await Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsExceptionAsync<System.IO.IOException>(()=>BridgeRequestReader.ReadAsync(stream,10,System.TimeSpan.FromSeconds(1)));
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsInstanceOfType(error.InnerException,typeof(System.Text.DecoderFallbackException));
            }
            foreach(bool fail in new[]{false,true}) using(var stream=new DeadlineStream(fail)) {
                var error=await Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsExceptionAsync<System.IO.IOException>(()=>BridgeRequestReader.ReadAsync(stream,10,System.TimeSpan.FromMilliseconds(20)));
                Microsoft.VisualStudio.TestTools.UnitTesting.StringAssert.Contains(error.Message,"timed out");
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsTrue(stream.Disposed);
            }
        }

        private sealed class DeadlineStream : System.IO.Stream
        {
            private readonly bool fail;
            private readonly System.Threading.Tasks.TaskCompletionSource<int> completion=new System.Threading.Tasks.TaskCompletionSource<int>();
            private byte[] buffer;
            public bool Disposed;
            public DeadlineStream(bool fail) { this.fail=fail; }
            public override System.Threading.Tasks.Task<int> ReadAsync(byte[] buffer,int offset,int count,System.Threading.CancellationToken cancellationToken) { this.buffer=buffer; return completion.Task; }
            protected override void Dispose(bool disposing) { Disposed=true; if(fail) completion.TrySetException(new System.IO.IOException("Closed read")); else {buffer[0]=10;completion.TrySetResult(1);}base.Dispose(disposing); }
            public override bool CanRead=>true;public override bool CanSeek=>false;public override bool CanWrite=>false;
            public override long Length=>throw new System.NotSupportedException();public override long Position {get=>throw new System.NotSupportedException();set=>throw new System.NotSupportedException();}
            public override int Read(byte[] buffer,int offset,int count)=>throw new System.NotSupportedException();
            public override void Flush()=>throw new System.NotSupportedException();public override long Seek(long offset,System.IO.SeekOrigin origin)=>throw new System.NotSupportedException();
            public override void SetLength(long value)=>throw new System.NotSupportedException();public override void Write(byte[] buffer,int offset,int count)=>throw new System.NotSupportedException();
        }
    }
}
