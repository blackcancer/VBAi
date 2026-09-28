using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace CodexVBE
{
    /// <summary>Lit une seule trame UTF-8 terminée par LF avec limites de taille et de durée totale.</summary>
    internal static class BridgeRequestReader
    {
                /// <summary>Ferme la connexion si sa réception reste incomplète au terme du délai.</summary>
                /// <param name="stream">The stream used by this operation.</param>
                /// <param name="maxBytes">The max bytes used by this operation.</param>
                /// <param name="timeout">The timeout used by this operation.</param>
                /// <returns>The result produced by this operation.</returns>
        internal static async Task<string> ReadAsync(Stream stream, int maxBytes, TimeSpan timeout)
        {
            if (maxBytes < 1) throw new ArgumentOutOfRangeException(nameof(maxBytes));
            if (timeout <= TimeSpan.Zero) throw new ArgumentOutOfRangeException(nameof(timeout));
            using (var deadline = new CancellationTokenSource())
            {
                Task<string> read = ReadFrameAsync(stream, maxBytes);
                Task delay = Task.Delay(timeout, deadline.Token);
                if (await Task.WhenAny(read, delay).ConfigureAwait(false) != read)
                {
                    // Disposing the asynchronous pipe releases the outstanding read; a silent
                    // peer must not retain a listener or a worker after its budget expires.
                    stream.Dispose();
                    try { await read.ConfigureAwait(false); } catch (Exception) { }
                    throw new IOException("The bridge request reception timed out.");
                }
                deadline.Cancel();
                return await read.ConfigureAwait(false);
            }
        }

        /// <summary>Reads and decodes one bounded UTF-8 frame terminated by a newline.</summary>
        /// <param name="stream">The stream used by this operation.</param>
        /// <param name="maxBytes">The max bytes used by this operation.</param>
        /// <returns>The result produced by this operation.</returns>
        private static async Task<string> ReadFrameAsync(Stream stream, int maxBytes)
        {
            var buffer = new byte[Math.Min(4096, maxBytes)];
            using (var frame = new MemoryStream())
            {
                while (true)
                {
                    // Only one byte beyond the limit is read to distinguish an exact-size
                    // request followed by LF from an over-budget request without a terminator.
                    int count = await stream.ReadAsync(buffer, 0,
                        (int)Math.Min(buffer.Length, (long)maxBytes - frame.Length + 1)).ConfigureAwait(false);
                    if (count == 0) throw new EndOfStreamException("The bridge request ended before its newline.");
                    int end = Array.IndexOf(buffer, (byte)'\n', 0, count);
                    int payload = end < 0 ? count : end;
                    if (frame.Length + payload > maxBytes) throw new IOException("The bridge request exceeds its byte limit.");
                    frame.Write(buffer, 0, payload);
                    if (end < 0) continue;
                    byte[] bytes = frame.GetBuffer();
                    int length = (int)frame.Length;
                    if (length > 0 && bytes[length - 1] == '\r') length--;
                    try { return new UTF8Encoding(false, true).GetString(bytes, 0, length); }
                    catch (DecoderFallbackException error) { throw new IOException("The bridge request is not valid UTF-8.", error); }
                }
            }
        }
    }
}
