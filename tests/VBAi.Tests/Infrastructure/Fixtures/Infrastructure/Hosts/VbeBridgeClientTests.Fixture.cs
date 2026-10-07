using System;
using System.IO;
using System.IO.Pipes;
using System.Text;
using System.Threading;

namespace VBAi.Tests.Infrastructure
{
    /// <summary>Serveur de pipe unique possédé, jamais connecté à un canal Office réel.</summary>
    internal sealed class BridgeClientPipeFixture : IDisposable
    {
        internal readonly string Name = "VBAi.Tests." + Guid.NewGuid().ToString("N");
        internal readonly ManualResetEventSlim Ready = new ManualResetEventSlim(), Received = new ManualResetEventSlim();
        private readonly ManualResetEventSlim release = new ManualResetEventSlim();
        private readonly Thread worker;
        private NamedPipeServerStream server;
        internal int Count;
        internal string Payload;
        internal Exception Failure;
        private volatile bool closing;

        /// <summary>Crée au besoin le serveur après un délai et traite une seule émission.</summary>
        internal BridgeClientPipeFixture(string reply, bool disconnect = false, bool hold = false, int startDelay = 0, string pipeName = null)
        {
            if (pipeName != null) Name = pipeName;
            worker = new Thread(() =>
            {
                try
                {
                    if (startDelay > 0 && release.Wait(startDelay)) return;
                    if (closing) return;
                    using (var owned = new NamedPipeServerStream(Name, PipeDirection.InOut, 1, PipeTransmissionMode.Byte, PipeOptions.Asynchronous))
                    {
                        server = owned; Ready.Set(); owned.WaitForConnection();
                        using (var reader = new StreamReader(owned, new UTF8Encoding(false), false, 4096, true))
                        using (var writer = new StreamWriter(owned, new UTF8Encoding(false), 4096, true) { AutoFlush = true })
                        {
                            Payload = reader.ReadLine(); if (Payload != null) Interlocked.Increment(ref Count); Received.Set();
                            if (hold) release.Wait();
                            else if (!disconnect) writer.WriteLine(reply);
                        }
                    }
                }
                catch (Exception ex) { if (!closing) Failure = ex; }
                finally { Ready.Set(); }
            })
            { IsBackground = true, Name = "OwnedBridgeClientPipe" };
            worker.Start();
        }

        /// <summary>Ferme uniquement le pipe unique de cette fixture et libère son thread.</summary>
        public void Dispose()
        {
            closing = true; release.Set(); server?.Dispose();
            if (!worker.Join(3000)) throw new InvalidOperationException("The owned pipe fixture did not stop.");
            Ready.Dispose(); Received.Dispose(); release.Dispose();
        }
    }
}