using System;
using System.Collections.Generic;

namespace VBAi
{
    /// <summary>Resolves links from the reference snapshot that belongs to a transcript entry.</summary>
    internal static class ChatTranscriptReferences
    {
        /// <summary>Uses explicit entry attachments, its preceding turn prompt, or a preceding legacy prompt.</summary>
        /// <param name="entries">Ordered entries from this conversation only.</param>
        /// <param name="entry">Entry being rendered; a foreign entry never inherits conversation references.</param>
        /// <returns>Unambiguous captured references, without reading live projects or future messages.</returns>
        internal static IDictionary<string, VbeChatReference> Resolve(IReadOnlyList<ChatEntry> entries, ChatEntry entry)
        {
            VbeChatReference[] snapshot = entry?.References;
            if (entry != null && snapshot == null && entry.Speaker != "Vous" && entries != null)
            {
                int index = -1;
                for (int i = 0; i < entries.Count; i++)
                    if (ReferenceEquals(entries[i], entry)) { index = i; break; }
                for (int i = index - 1; i >= 0; i--)
                {
                    var prompt = entries[i];
                    if (prompt == null || prompt.Speaker != "Vous") continue;
                    if (!string.IsNullOrEmpty(entry.TurnId))
                    {
                        if (!string.Equals(prompt.TurnId, entry.TurnId, StringComparison.Ordinal)) continue;
                    }
                    else if (!string.IsNullOrEmpty(prompt.TurnId))
                    {
                        // A legacy entry must not borrow a newer structured turn's context.
                        break;
                    }
                    snapshot = prompt.References;
                    break;
                }
            }
            return Unique(snapshot);
        }

        /// <summary>Rejects conflicting destinations sharing a visible token in one captured snapshot.</summary>
        /// <param name="snapshot">Captured attachments; null and empty both provide no links.</param>
        /// <returns>One reference per token, excluding every ambiguous token.</returns>
        private static IDictionary<string, VbeChatReference> Unique(IEnumerable<VbeChatReference> snapshot)
        {
            var result = new Dictionary<string, VbeChatReference>(StringComparer.Ordinal);
            var ambiguous = new HashSet<string>(StringComparer.Ordinal);
            if (snapshot == null) return result;
            foreach (var reference in snapshot)
            {
                if (reference == null || string.IsNullOrWhiteSpace(reference.Project)) continue;
                string token = reference.Token;
                if (ambiguous.Contains(token)) continue;
                if (!result.TryGetValue(token, out var previous)) result.Add(token, reference);
                else if (!SameSnapshot(previous, reference))
                {
                    result.Remove(token);
                    ambiguous.Add(token);
                }
            }
            return result;
        }

        /// <summary>Compares captured target and freshness fields without consulting current host state.</summary>
        /// <param name="first">First reference with this token.</param>
        /// <param name="second">Another reference with this token.</param>
        /// <returns>True only when the recorded destination, range and revision are identical.</returns>
        private static bool SameSnapshot(VbeChatReference first, VbeChatReference second) =>
            string.Equals(first.Project, second.Project, StringComparison.Ordinal) &&
            string.Equals(first.Module, second.Module, StringComparison.Ordinal) &&
            string.Equals(first.Name, second.Name, StringComparison.Ordinal) &&
            string.Equals(first.DeclarationScope, second.DeclarationScope, StringComparison.Ordinal) &&
            string.Equals(first.Kind, second.Kind, StringComparison.Ordinal) &&
            string.Equals(first.Sha256, second.Sha256, StringComparison.Ordinal) &&
            first.ProcKind == second.ProcKind && first.StartLine == second.StartLine &&
            first.EndLine == second.EndLine && first.DeclarationColumn == second.DeclarationColumn;
    }
}
