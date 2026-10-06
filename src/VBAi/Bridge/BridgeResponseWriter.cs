using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace VBAi
{

    /// <summary>Sends a single UTF-8 response without allowing a stalled client to retain the listener.</summary>
    internal static class BridgeResponseWriter
    {

        /// <summary>Writes async for bridge response writer.</summary>
        /// <param name="stream">stream that supplies the stream for this operation.</param>
        /// <param name="payload">Text that supplies the payload value. Use the format required by the calling operation.</param>
        /// <param name="timeout">Duration that supplies the timeout for this operation.</param>
        /// <returns>task produced by the operation for write async on bridge response writer.</returns>
        internal static async Task WriteAsync(Stream stream, string payload, TimeSpan timeout)
        {
            if (timeout <= TimeSpan.Zero) throw new ArgumentOutOfRangeException(nameof(timeout));
            byte[] bytes = new UTF8Encoding(false).GetBytes(payload + "\n");
            using (var deadline = new CancellationTokenSource())
            {
                Task write = stream.WriteAsync(bytes, 0, bytes.Length);
                Task delay = Task.Delay(timeout, deadline.Token);
                if (await Task.WhenAny(write, delay).ConfigureAwait(false) != write)
                {
                    stream.Dispose();
                    // Observe a later I/O failure without waiting again on a stalled transport.
                    _ = write.ContinueWith(task => { var error = task.Exception; }, CancellationToken.None,
                        TaskContinuationOptions.OnlyOnFaulted | TaskContinuationOptions.ExecuteSynchronously, TaskScheduler.Default);
                    throw new IOException("The bridge response transmission timed out. The command may already have completed; do not automatically retry edits.");
                }
                deadline.Cancel();
                await write.ConfigureAwait(false);
            }
        }
    }
}
