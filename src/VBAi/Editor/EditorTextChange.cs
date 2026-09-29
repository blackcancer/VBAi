using System;
using System.Linq;

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
            foreach (var change in changes.OrderByDescending(c => c?.rangeOffset ?? -1))
            {
                if (change == null || change.text == null || change.rangeOffset < 0 || change.rangeLength < 0 ||
                    change.rangeOffset > boundary || change.rangeLength > boundary - change.rangeOffset) return false;
                length += (long)change.text.Length - change.rangeLength;
                if (length > EditorDocument.MaxLength || change.text.IndexOf('\0') >= 0) return false;
                boundary = change.rangeOffset;
            }
            foreach (var change in changes.OrderByDescending(c => c.rangeOffset))
                result = result.Remove(change.rangeOffset, change.rangeLength).Insert(change.rangeOffset, change.text);
            return true;
        }
    }
}
