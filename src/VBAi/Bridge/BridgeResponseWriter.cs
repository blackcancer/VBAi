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

        /// <summary>Writes one UTF-8 payload followed by LF; on timeout it closes the stream and reports that the command outcome may be uncertain.</summary>
        /// <param name="stream">Connected pipe stream used for the one response write.</param>
        /// <param name="payload">Serialized response body without its line terminator.</param>
        /// <param name="timeout">Maximum write wait; must be positive.</param>
        /// <returns>A task that completes when the write finishes or faults on timeout or I/O failure.</returns>
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
