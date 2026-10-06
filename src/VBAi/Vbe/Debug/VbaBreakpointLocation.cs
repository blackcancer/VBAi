using System;
using System.Linq;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace VBAi
{

    /// <summary>Filters non-executable source locations before invoking the native breakpoint command.</summary>
    internal static class VbaBreakpointLocation
    {

        /// <summary>Checks the first physical line of a statement inside a procedure.</summary>
        /// <remarks>This is a source filter, not a substitute for the VBA compiler or a breakpoint inventory.</remarks>
        /// <param name="source">VBA module source text, including line continuations and conditional-compilation directives.</param>
        /// <param name="line">One-based physical source line requested for a native breakpoint toggle.</param>
        /// <returns>True only for a known-active executable statement inside a Sub/Function/Property body.</returns>
        internal static bool CanRequest(string source, int line)
        {
            if (string.IsNullOrEmpty(source) || line < 1) return false;
            string[] lines = source.Replace("\r", "").Split('\n');
            if (line > lines.Length) return false;
            var activeLines = KnownActiveLines(lines);
            if (!activeLines[line - 1]) return false;
            bool procedure = false;
            foreach (var tokens in VbaDeclarationIndex.Statements(source))
            {
                if (tokens.Count == 0) continue;
                if (tokens[0].Line > line) break;
                if (!activeLines[tokens[0].Line - 1]) continue;
                int first = 0;
                // Numeric line numbers precede statements without a colon.
                while (first < tokens.Count && tokens[first].Text.All(char.IsDigit)) first++;
                if (first == tokens.Count) continue;
                string word = tokens[first].Text.ToUpperInvariant();
                int signature = first;
                while (signature < tokens.Count && new[] { "PUBLIC", "PRIVATE", "FRIEND", "STATIC" }.Contains(tokens[signature].Text.ToUpperInvariant())) signature++;
                if (signature < tokens.Count && new[] { "SUB", "FUNCTION", "PROPERTY" }.Contains(tokens[signature].Text.ToUpperInvariant()))
                { procedure = true; continue; }
                bool end = word == "END" && first + 1 < tokens.Count &&
                    new[] { "SUB", "FUNCTION", "PROPERTY" }.Contains(tokens[first + 1].Text.ToUpperInvariant());
                bool inProcedure = procedure;
                if (end) procedure = false;
                if (!inProcedure || tokens[0].Line != line) continue;
                if (new[] { "DIM", "CONST", "STATIC", "PUBLIC", "PRIVATE", "FRIEND", "DECLARE", "TYPE", "ENUM", "OPTION", "ATTRIBUTE", "#" }.Contains(word)) continue;
                // A lone identifier followed by ':' is a label, not a procedure call.
                int after = tokens[first].Column - 1 + tokens[first].Text.Length;
                if (tokens.Count == first + 1 && lines[line - 1].Substring(Math.Min(after, lines[line - 1].Length)).TrimStart().StartsWith(":")) continue;
                return true;
            }
            return false;
        }

        /// <summary>Nested conditional-compilation state used to fail closed on unknown constants.</summary>
        private sealed class ConditionalBranch
        {

            /// <summary>Whether the enclosing branch is active, an earlier arm was taken, or any arm is unknown.</summary>
            internal bool Parent, Taken, Unknown;
        }

        /// <summary>Unknown compiler constants cannot safely authorize a native toggle.</summary>
        /// <param name="lines">Physical module lines without carriage returns.</param>
        /// <returns>Per-line activity mask; malformed directives and conditions using unknown constants are inactive.</returns>
        private static bool[] KnownActiveLines(string[] lines)
        {
            var result = new bool[lines.Length];
            var stack = new Stack<ConditionalBranch>();
            bool active = true;
            for (int i = 0; i < lines.Length; i++)
            {
                string text = lines[i].Trim();
                var condition = Regex.Match(text, @"^#(If|ElseIf)\s+(.+?)\s+Then\s*(?:'.*)?$", RegexOptions.IgnoreCase);
                if (condition.Success)
                {
                    bool first = condition.Groups[1].Value.Equals("If", StringComparison.OrdinalIgnoreCase);
                    if (first) stack.Push(new ConditionalBranch { Parent = active });
                    if (stack.Count == 0) { active = false; continue; }
                    var branch = stack.Peek();
                    string expression = condition.Groups[2].Value.Trim();
                    bool known = expression.Equals("True", StringComparison.OrdinalIgnoreCase) || expression == "-1" || expression == "1" ||
                        expression.Equals("False", StringComparison.OrdinalIgnoreCase) || expression == "0";
                    bool taken = known && expression != "0" && !expression.Equals("False", StringComparison.OrdinalIgnoreCase);
                    active = branch.Parent && !branch.Taken && !branch.Unknown && known && taken;
                    branch.Unknown |= !known;
                    branch.Taken |= taken;
                    continue;
                }
                if (Regex.IsMatch(text, @"^#Else\s*(?:'.*)?$", RegexOptions.IgnoreCase))
                {
                    if (stack.Count == 0) { active = false; continue; }
                    var branch = stack.Peek(); active = branch.Parent && !branch.Taken && !branch.Unknown; branch.Taken = true; continue;
                }
                if (Regex.IsMatch(text, @"^#End\s+If\s*(?:'.*)?$", RegexOptions.IgnoreCase))
                {
                    active = stack.Count != 0 && stack.Pop().Parent; continue;
                }
                // An unrecognized directive is never a statement. Unknown/malformed #If blocks fail closed.
                if (Regex.IsMatch(text, @"^#If\b", RegexOptions.IgnoreCase))
                { stack.Push(new ConditionalBranch { Parent = active, Unknown = true }); active = false; continue; }
                result[i] = active && !text.StartsWith("#", StringComparison.Ordinal);
            }
            return result;
        }
    }
}
