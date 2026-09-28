using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace CodexVBE
{
    internal static class EditorAttributeRewrite
    {
        private sealed class Procedure
        {
            internal string Name, Kind;
            internal int First, Last;
            internal readonly List<string> Attributes = new List<string>();
        }
        private static readonly Regex Attribute = new Regex(@"^\s*Attribute\s+", RegexOptions.IgnoreCase);
        private static List<Procedure> Procedures(string source)
        {
            var result = new List<Procedure>();
            foreach (var tokens in VbaDeclarationIndex.Statements(source))
            {
                int p = 0;
                while (p < tokens.Count && new[] { "public", "private", "friend", "static" }.Contains(tokens[p].Text.ToLowerInvariant())) p++;
                if (p >= tokens.Count) continue;
                string kind = tokens[p].Text.ToLowerInvariant();
                if (!new[] { "sub", "function", "property" }.Contains(kind)) continue;
                int name = p + (kind == "property" ? 2 : 1);
                if (name >= tokens.Count || !Regex.IsMatch(tokens[name].Text, @"^[A-Za-z_][A-Za-z0-9_]*$")) continue;
                if (!tokens.Any(t => t.Text == "(") || !tokens.Any(t => t.Text == ")")) return null;
                result.Add(new Procedure { Name = tokens[name].Text, Kind = kind == "property" ? kind + ":" + tokens[p + 1].Text.ToLowerInvariant() : kind,
                    First = tokens[0].Line, Last = tokens[tokens.Count - 1].Line });
            }
            return result;
        }
        // Export headers describe the designer/class. Only the code section is reloaded into
        // the existing CodeModule; the component and its designer are never replaced.
        internal static string CodeSection(string exported)
        {
            string text = EditorDocument.Normalize(exported);
            var start = Regex.Match(text, @"(?im)^Attribute VB_Name\s*=");
            return start.Success ? text.Substring(start.Index) : text;
        }
        internal static string Prepare(string exported, string before, Tuple<int, int, string> patch)
        {
            before = EditorDocument.Normalize(before);
            var code = CodeSection(exported).Split('\n');
            var visible = code.Where(line => !Attribute.IsMatch(line)).ToArray();
            if (string.Join("\n", visible).TrimEnd('\n') != before.TrimEnd('\n')) return null;
            var lines = before.Split('\n').ToList();
            if (patch.Item1 < 1 || patch.Item2 < 0 || patch.Item1 - 1 + patch.Item2 > lines.Count) return null;
            lines.RemoveRange(patch.Item1 - 1, patch.Item2);
            if (patch.Item3.Length != 0) lines.InsertRange(patch.Item1 - 1, EditorDocument.Normalize(patch.Item3).Split('\n'));
            string after = string.Join("\n", lines);
            // A conditional procedure can have identical names and different attributes.
            if (Regex.IsMatch(before + "\n" + after, @"(?im)^\s*#\s*(If|Else|End)\b")) return null;
            var old = Procedures(before); var next = Procedures(after);
            if (old == null || next == null || old.Count != next.Count) return null;
            var globals = new List<string>(); int visibleLine = 0;
            foreach (string line in code)
            {
                if (!Attribute.IsMatch(line)) { visibleLine++; continue; }
                var member = Regex.Match(line, @"^\s*Attribute\s+([^\s.]+)\.", RegexOptions.IgnoreCase);
                if (!member.Success) { globals.Add(line); continue; }
                var owner = old.LastOrDefault(d => d.Last <= visibleLine);
                if (owner == null || owner.Last != visibleLine || !string.Equals(owner.Name, member.Groups[1].Value, StringComparison.OrdinalIgnoreCase)) return null;
                owner.Attributes.Add(line);
            }
            bool touches = old.Any(d => d.Attributes.Count > 0 && (patch.Item2 > 0
                ? patch.Item1 <= d.Last && patch.Item1 + patch.Item2 - 1 >= d.First
                : patch.Item1 > d.First && patch.Item1 <= d.Last));
            if (!touches) return null;
            var used = new HashSet<Procedure>();
            for (int i = 0; i < old.Count; i++)
            {
                var source = old[i];
                var matches = next.Where(d => d.Kind == source.Kind && string.Equals(d.Name, source.Name, StringComparison.OrdinalIgnoreCase)).ToArray();
                Procedure target;
                if (matches.Length == 1) target = matches[0];
                else if (matches.Length == 0 && next[i].Kind == source.Kind && !old.Any(d => string.Equals(d.Name, next[i].Name, StringComparison.OrdinalIgnoreCase))) target = next[i];
                else return null;
                if (!used.Add(target)) return null;
                foreach (string attribute in source.Attributes)
                {
                    var owner = Regex.Match(attribute, @"^(\s*Attribute\s+)([^\s.]+)\.(.*)$", RegexOptions.IgnoreCase);
                    // Parameter metadata can become invalid after a signature change. Preserve it
                    // only when the signature (apart from the procedure name/line wrapping) is unchanged.
                    if (!owner.Groups[3].Value.StartsWith("VB_", StringComparison.OrdinalIgnoreCase))
                    {
                        string Signature(Procedure d, string text) => string.Join(" ", VbaDeclarationIndex.Statements(text).First(s => s[0].Line == d.First).Select(t => t.Text == d.Name ? "$owner" : t.Text));
                        if (!string.Equals(Signature(source, before), Signature(target, after), StringComparison.OrdinalIgnoreCase)) return null;
                    }
                    target.Attributes.Add(owner.Groups[1].Value + target.Name + "." + owner.Groups[3].Value);
                }
            }
            var output = new List<string>(globals);
            for (int i = 0; i < lines.Count; i++)
            {
                output.Add(lines[i]);
                foreach (var declaration in next.Where(d => d.Last == i + 1)) output.AddRange(declaration.Attributes);
            }
            return string.Join("\r\n", output);
        }
        internal static bool HasMultilineAttributes(string exported)
        {
            string code = CodeSection(exported);
            var procedures = Procedures(string.Join("\n", code.Split('\n').Where(line => !Attribute.IsMatch(line))));
            return procedures == null || procedures.Any(d => d.First != d.Last && Regex.IsMatch(code, @"(?im)^\s*Attribute\s+" + Regex.Escape(d.Name) + @"\."));
        }
        internal static string Metadata(string source) => string.Join("\n", EditorDocument.Normalize(source).Split('\n').Where(line => Attribute.IsMatch(line)));
    }
}
