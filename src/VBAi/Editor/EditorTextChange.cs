using System;
using System.Linq;
using System.Text;

namespace VBAi
{
    /// <summary>A UTF-16 edit against one specific Monaco model revision.</summary>
    internal sealed class EditorTextChange
    {
        /// <summary>Offset in the original model, in UTF-16 code units.</summary>
        public int rangeOffset { get; set; }
        /// <summary>Number of original code units replaced.</summary>
        public int rangeLength { get; set; }
        /// <summary>Replacement content.</summary>
        public string text { get; set; }
        /// <summary>Applies a non-overlapping change batch atomically; malformed batches request a full snapshot.</summary>
        internal static bool TryApply(string source, EditorTextChange[] changes, out string result)
        {
            result = source;
            if (changes == null || changes.Length == 0 || changes.Length > 10000) return false;
            long length = source.Length;
            int boundary = source.Length;
            var ordered = changes.OrderByDescending(c => c?.rangeOffset ?? -1).ToArray();
            foreach (var change in ordered)
            {
                if (change == null || change.text == null || change.rangeOffset < 0 || change.rangeLength < 0 ||
                    change.rangeOffset > boundary || change.rangeLength > boundary - change.rangeOffset) return false;
                length += (long)change.text.Length - change.rangeLength;
                if (change.text.IndexOf('\0') >= 0) return false;
                boundary = change.rangeOffset;
            }
            if (length > EditorDocument.MaxLength) return false;
            // Validate the complete atomic batch before allocating or publishing its result.
            // Copy each unchanged span once, preserving right-to-left ordering for tied insertions.
            var builder = new StringBuilder((int)length);
            int position = 0;
            for (int i = ordered.Length - 1; i >= 0; i--)
            {
                var change = ordered[i];
                builder.Append(source, position, change.rangeOffset - position);
                builder.Append(change.text);
                position = change.rangeOffset + change.rangeLength;
            }
            builder.Append(source, position, source.Length - position);
            result = builder.ToString();
            return true;
        }
    }
}
