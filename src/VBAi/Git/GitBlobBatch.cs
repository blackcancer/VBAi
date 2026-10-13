using System;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace VBAi
{
    /// <summary>Decodes ordered Git blob replies without converting their binary bodies to text.</summary>
    internal static class GitBlobBatch
    {
        /// <summary>Allows full SHA-1 or SHA-256 object identities, never revision expressions or paths.</summary>
        private static readonly Regex ObjectId = new Regex(@"\A(?:[0-9a-f]{40}|[0-9a-f]{64})\z", RegexOptions.CultureInvariant);

        /// <summary>Immutable identity and size obtained from the reviewed Git tree.</summary>
        internal sealed class Entry
        {
            /// <summary>Full object identity requested from Git.</summary>
            internal readonly string Id;
            /// <summary>Expected binary body length from the tree listing.</summary>
            internal readonly long Size;
            /// <summary>Captures a request; validation happens before launching the batch process.</summary>
            /// <param name="id">Full lowercase object identity from the tree listing.</param>
            /// <param name="size">Expected body length; checked before a batch is launched.</param>
            internal Entry(string id, long size) { Id = id; Size = size; }
        }

        /// <summary>Validates the entire immutable request before sending any object to Git.</summary>
        /// <param name="entries">Ordered object identities and sizes read from the reviewed tree.</param>
        internal static void Validate(Entry[] entries)
        {
            if (entries == null) throw new ArgumentNullException(nameof(entries));
            // Preserve the existing reader's admission bound; package validation may impose a tighter one.
            if (entries.Length > 2049) throw new InvalidOperationException(UiText.Get("The VBA repository exceeds 32 MB."));
            long total = 0;
            foreach (var entry in entries)
            {
                if (entry == null || entry.Id == null || !ObjectId.IsMatch(entry.Id) || entry.Size < 0 || entry.Size > VbaGitSnapshot.MaxBytes)
                    throw new InvalidOperationException("The Git tree contains an invalid blob identity or size.");
                total += entry.Size;
                if (total > VbaGitSnapshot.MaxBytes) throw new InvalidOperationException(UiText.Get("The VBA repository exceeds 32 MB."));
            }
        }

        /// <summary>Checks each header, exact binary body length and delimiter against its corresponding request.</summary>
        /// <param name="output">Unmodified binary output from one cat-file batch process.</param>
        /// <param name="entries">Ordered identities and sizes used for that exact request.</param>
        /// <returns>Independent binary bodies, in request order.</returns>
        internal static byte[][] Read(byte[] output, Entry[] entries)
        {
            if (output == null) throw new ArgumentNullException(nameof(output));
            Validate(entries);
            var blobs = new byte[entries.Length][];
            int offset = 0;
            for (int index = 0; index < entries.Length; index++)
            {
                var entry = entries[index];
                string expected = entry.Id + " blob " + entry.Size.ToString(CultureInfo.InvariantCulture);
                int headerLength = expected.Length;
                if (output.Length - offset <= headerLength || output[offset + headerLength] != (byte)'\n' ||
                    !string.Equals(Encoding.ASCII.GetString(output, offset, headerLength), expected, StringComparison.Ordinal))
                    throw new InvalidOperationException("The Git batch header does not match the requested blob.");
                offset += headerLength + 1;
                int length = (int)entry.Size;
                if (output.Length - offset <= length || output[offset + length] != (byte)'\n')
                    throw new InvalidOperationException("The Git batch body is truncated or has an invalid delimiter.");
                blobs[index] = new byte[length];
                Buffer.BlockCopy(output, offset, blobs[index], 0, length);
                offset += length + 1;
            }
            if (offset != output.Length) throw new InvalidOperationException("The Git batch reply contains unexpected trailing data.");
            return blobs;
        }
    }
}
