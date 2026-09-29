using System;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using VBAi;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class BridgeResponseBudgetTests
    {
        [TestMethod]
        public async Task WriterSendsUtf8AndBoundsAnUnresponsiveTransport()
        {
            using (var memory = new MemoryStream())
            {
                await BridgeResponseWriter.WriteAsync(memory, "é漢", TimeSpan.FromSeconds(1));
                Assert.AreEqual("é漢\n", Encoding.UTF8.GetString(memory.ToArray()));
                await Assert.ThrowsExceptionAsync<ArgumentOutOfRangeException>(() => BridgeResponseWriter.WriteAsync(memory, "", TimeSpan.Zero));
            }
            using (var blocked = new BlockedWriter())
            {
                await Assert.ThrowsExceptionAsync<IOException>(() => BridgeResponseWriter.WriteAsync(blocked, "{}", TimeSpan.FromMilliseconds(50)));
                Assert.IsTrue(blocked.Disposed);
            }
        }

        [TestMethod]
        public void HealthyClientResumesAfterAClientStopsReadingWithoutReplayingItsCommand()
        {
            int id = Guid.NewGuid().GetHashCode() & int.MaxValue;
            int executed = 0;
            bool networkDenied = false;
            using (var dispatcher = new Control())
            using (var server = new BridgeServer(dispatcher, null, id))
            {
                var open = server.OpenPipe;
                server.OpenPipe = security =>
                {
                    networkDenied = security.GetAccessRules(true, true, typeof(SecurityIdentifier)).OfType<PipeAccessRule>()
                        .Any(rule => rule.AccessControlType == AccessControlType.Deny &&
                            rule.IdentityReference.Value == new SecurityIdentifier(WellKnownSidType.NetworkSid, null).Value);
                    return open(security);
                };
                server.ResponseWriteTimeout = TimeSpan.FromMilliseconds(150);
                server.Native.Capture = include => Interlocked.Increment(ref executed) == 1
                    ? (object)new { Text = new string('x', 2 * 1024 * 1024) } : new { Healthy = true };
                server.Start();
                using (var stalled = Connect(id))
                {
                    Send(stalled);
                    Assert.IsTrue(SpinWait.SpinUntil(() => Volatile.Read(ref executed) == 1, 4000));
                    using (var healthy = Connect(id))
                    {
                        Send(healthy);
                        var reply = Task.Run(() => new StreamReader(healthy, Encoding.UTF8, false, 1024, true).ReadLine());
                        Assert.IsTrue(reply.Wait(4000));
                        Assert.IsTrue(new JavaScriptSerializer().Deserialize<Response>(reply.Result).Ok);
                        Assert.AreEqual(2, executed);
                        Assert.IsTrue(networkDenied);
                    }
                }
            }
        }

        [TestMethod]
        public void UnserializableResponseReturnsAnExplicitUnknownResultWithoutStoppingTheServer()
        {
            int id = Guid.NewGuid().GetHashCode() & int.MaxValue;
            using (var dispatcher = new Control())
            using (var server = new BridgeServer(dispatcher, null, id))
            {
                server.Native.Capture = include => new RecursivePayload();
                server.Start();
                using (var client = Connect(id))
                {
                    Send(client);
                    var response = Task.Run(() => new StreamReader(client, Encoding.UTF8, false, 1024, true).ReadLine());
                    Assert.IsTrue(response.Wait(4000));
                    var result = new JavaScriptSerializer().Deserialize<Response>(response.Result);
                    Assert.IsFalse(result.Ok);
                    StringAssert.Contains(result.Error, "may already have completed");
                }
            }
        }

        private sealed class RecursivePayload { public RecursivePayload Value => this; }
        private static NamedPipeClientStream Connect(int id)
        {
            var client = new NamedPipeClientStream(".", "VBAi." + id, PipeDirection.InOut, PipeOptions.Asynchronous);
            try { client.Connect(4000); return client; } catch { client.Dispose(); throw; }
        }
        private static void Send(Stream stream)
        {
            byte[] request = Encoding.UTF8.GetBytes("{\"Command\":\"debug_windows\"}\n");
            stream.Write(request, 0, request.Length);
        }
        private sealed class BlockedWriter : MemoryStream
        {
            internal bool Disposed;
            private readonly TaskCompletionSource<bool> pending = new TaskCompletionSource<bool>();
            public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) => pending.Task;
            protected override void Dispose(bool disposing)
            {
                Disposed = true;
                pending.TrySetException(new IOException("closed"));
                base.Dispose(disposing);
            }
        }
    }
}
