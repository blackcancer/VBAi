using System;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace CodexVBE
{
    // Text transformations are explicit about scope. Identifier replacement is lexical,
    // not a claim of semantic binding across VBA/COM projects.
    internal static class VbaTextEdits
    {
        internal static string Transform(string source, Request request)
        {
            string[] lines = CodeRollback.Lines(source);
            if (request.StartLine < 1 || request.Count < 1 || request.StartLine > lines.Length ||
                request.Count > lines.Length - request.StartLine + 1)
                throw new ArgumentException("StartLine and Count must select existing lines.");
            if ((request.Text ?? "").Length > 262144 || (request.Query ?? "").Length > 4096)
                throw new ArgumentException("Replacement text is too large.");
            int first = request.StartLine - 1;
            string selected = string.Join("\r\n", lines.Skip(first).Take(request.Count));
            string replacement;
            switch (request.Action)
            {
                case "replace":
                    if (string.IsNullOrEmpty(request.Query) || request.Text == null || request.Query.Contains("\n") || request.Query.Contains("\r"))
                        throw new ArgumentException("A single-line Query and Text are required.");
                    string pattern = Regex.Escape(request.Query);
                    if (request.WholeWord) pattern = @"(?<![\p{L}\p{N}_])" + pattern + @"(?![\p{L}\p{N}_])";
                    replacement = Regex.Replace(selected, pattern, m => request.Text,
                        request.MatchCase ? RegexOptions.None : RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
                    break;
                case "replace_identifier":
                    ValidateIdentifier(request.Query); ValidateIdentifier(request.NewName);
                    replacement = ReplaceIdentifier(selected, request.Query, request.NewName);
                    break;
                case "comment":
                    replacement = string.Join("\r\n", lines.Skip(first).Take(request.Count).Select(line => "'" + line));
                    break;
                case "uncomment":
                    replacement = string.Join("\r\n", lines.Skip(first).Take(request.Count).Select(line => Regex.Replace(line, @"^(\s*)'", "$1")));
                    break;
                case "indent": case "unindent":
                    int size = request.InsertIndex ?? 4;
                    if (size < 1 || size > 16) throw new ArgumentException("Indent size must be between 1 and 16.");
                    replacement = string.Join("\r\n", lines.Skip(first).Take(request.Count).Select(line => request.Action == "indent"
                        ? new string(' ', size) + line : line.StartsWith("\t", StringComparison.Ordinal) ? line.Substring(1)
                        : line.Substring(Math.Min(size, line.TakeWhile(c => c == ' ').Count()))));
                    break;
                default: throw new ArgumentException("Use replace, replace_identifier, comment, uncomment, indent or unindent.");
            }
            return string.Join("\r\n", lines.Take(first).Concat(CodeRollback.Lines(replacement)).Concat(lines.Skip(first + request.Count)));
        }

        internal static void ValidateIdentifier(string value)
        {
            if (!Regex.IsMatch(value ?? "", @"^[A-Za-z][A-Za-z0-9_]{0,254}$"))
                throw new ArgumentException("A VBA identifier is required.");
            if (Regex.IsMatch(value, @"^(?:Rem|If|Then|Else|End|Sub|Function|Property|Dim|As|Set|Let|Get|Public|Private|Friend|Static|Const|For|Next|Do|Loop|While|With|Select|Case|Option|ByVal|ByRef|Me|Nothing|True|False|New|Call|Exit)$", RegexOptions.IgnoreCase))
                throw new ArgumentException("A reserved VBA word is not a replacement identifier.");
        }

        internal static string ReplaceIdentifier(string text, string oldName, string newName)
        {
            var output = new StringBuilder();

            for (int i = 0; i < text.Length;)
            {
                char c = text[i];
                if (c == '\r' || c == '\n' || c == ':') { output.Append(c); i++; continue; }
                if (c == '\'' || (i + 3 <= text.Length && string.Equals(text.Substring(i, 3), "Rem", StringComparison.OrdinalIgnoreCase) && (i + 3 == text.Length || char.IsWhiteSpace(text[i + 3]))))
                { while (i < text.Length && text[i] != '\r' && text[i] != '\n') output.Append(text[i++]); continue; }
                if (c == '"' || c == '[' || c == '#')
                {
                    // Quoted strings, bracket expressions and date/directive text are opaque.
                    char end = c == '[' ? ']' : c;
                    output.Append(text[i++]);
                    while (i < text.Length && text[i] != '\r' && text[i] != '\n')
                    {
                        char next = text[i++]; output.Append(next);
                        if (next != end) continue;
                        if (end == '"' && i < text.Length && text[i] == '"') { output.Append(text[i++]); continue; }
                        break;
                    }
                     continue;
                }
                if (char.IsLetter(c) || c == '_')
                {
                    int start = i++;
                    while (i < text.Length && (char.IsLetterOrDigit(text[i]) || text[i] == '_')) i++;
                    string token = text.Substring(start, i - start);
                    output.Append(string.Equals(token, oldName, StringComparison.OrdinalIgnoreCase) ? newName : token);

                }
                else { output.Append(c); i++; }
            }
            return output.ToString();
        }
    }
}
