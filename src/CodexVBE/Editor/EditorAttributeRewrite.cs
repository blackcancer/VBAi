using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace CodexVBE
{
    internal static class EditorAttributeRewrite
    {
        // Only a single physical declaration line can be mapped unambiguously.
        internal static string Prepare(string exported, string before, Tuple<int, int, string> patch)
        {
            if (patch.Item2 != 1 || patch.Item3.Contains("\n")) return null;
            var lines = EditorDocument.Normalize(exported).Split('\n');
            var visible = lines.Select((text, index) => new { text, index })
                .Where(x => !Regex.IsMatch(x.text, @"^\s*Attribute\s", RegexOptions.IgnoreCase)).ToArray();
            if (string.Join("\n", visible.Select(x => x.text)).TrimEnd('\n') != before.TrimEnd('\n')) return null;
            if (patch.Item1 < 1 || patch.Item1 > visible.Length) return null;
            const string declaration = @"^\s*(?:(?:Public|Private|Friend|Static)\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_][A-Za-z0-9_]*)\s*\([^\r\n]*\)(?:\s+As\s+[A-Za-z_][A-Za-z0-9_.]*(?:\(\))?)?\s*$";
            var old = Regex.Match(visible[patch.Item1 - 1].text, declaration, RegexOptions.IgnoreCase);
            var next = Regex.Match(patch.Item3, declaration, RegexOptions.IgnoreCase);
            if (!old.Success || !next.Success || !string.Equals(old.Groups[1].Value, next.Groups[1].Value, StringComparison.OrdinalIgnoreCase)) return null;
            // A parameter/signature edit can invalidate parameter attributes. Only rename the owner.
            string oldSignature = visible[patch.Item1 - 1].text.Remove(old.Groups[2].Index, old.Groups[2].Length);
            string newSignature = patch.Item3.Remove(next.Groups[2].Index, next.Groups[2].Length);
            if (!string.Equals(oldSignature, newSignature, StringComparison.OrdinalIgnoreCase)) return null;
            int physical = visible[patch.Item1 - 1].index;
            var owner = new Regex(@"^(\s*Attribute\s+)" + Regex.Escape(old.Groups[2].Value) + @"(?=\.)", RegexOptions.IgnoreCase);
            bool found = false;
            // Attributes belong to this declaration, not other accessors with the same name.
            for (int i = physical + 1; i < lines.Length && Regex.IsMatch(lines[i], @"^\s*Attribute\s", RegexOptions.IgnoreCase); i++)
            {
                if (!owner.IsMatch(lines[i])) return null;
                lines[i] = owner.Replace(lines[i], m => m.Groups[1].Value + next.Groups[2].Value);
                found = true;
            }
            if (!found) return null;
            lines[physical] = patch.Item3;
            return string.Join("\r\n", lines);
        }
        internal static string Metadata(string source) => string.Join("\n", EditorDocument.Normalize(source).Split('\n')
            .Where(line => Regex.IsMatch(line, @"^\s*Attribute\s", RegexOptions.IgnoreCase)));
    }
}
