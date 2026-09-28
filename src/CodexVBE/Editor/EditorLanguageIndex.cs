using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace CodexVBE
{
    internal sealed class EditorSource
    {
        public string Module { get; set; }
        public string Text { get; set; }
        public int ComponentType { get; set; }
    }
    internal sealed class EditorSymbol
    {
        public string Name { get; set; }
        public string Module { get; set; }
        public string Kind { get; set; }
        public string Scope { get; set; }
        public string TypeName { get; set; }
        public string Declaration { get; set; }
        public bool Private { get; set; }
        public bool Conditional { get; set; }
        public bool External { get; set; }
        public string Library { get; set; }
        public int Line { get; set; }
        public int EndLine { get; set; }
        public int Column { get; set; }
        public string[] Parameters { get; set; } = Array.Empty<string>();
    }
    /// <summary>Pure snapshot analysis. No COM or UI access; executed by the synchronization worker.</summary>
    internal static class EditorLanguageIndex
    {
        internal static EditorSymbol[] Build(EditorSource[] sources)
        {
            var result = new List<EditorSymbol>();
            foreach (var source in sources)
            {
                string[] lines = EditorDocument.Normalize(source.Text).Split('\n');
                result.Add(new EditorSymbol { Name = source.Module, Module = source.Module, Kind = source.ComponentType == 1 ? "Module" : "Class", Scope = "Module", Line = 1, Column = 1, EndLine = lines.Length, Declaration = source.Module });
                var procedures = new List<EditorSymbol>();
                int conditionalDepth = 0;
                foreach (var statement in VbaDeclarationIndex.Statements(source.Text))
                {
                    if (statement.Count > 1 && statement[0].Text == "#")
                    {
                        if (statement[1].Text.Equals("If", StringComparison.OrdinalIgnoreCase)) conditionalDepth++;
                        else if (statement[1].Text.Equals("End", StringComparison.OrdinalIgnoreCase)) conditionalDepth = Math.Max(0, conditionalDepth - 1);
                        continue;
                    }
                    int p = 0;
                    while (p < statement.Count && new[] { "public", "private", "friend", "static" }.Contains(statement[p].Text.ToLowerInvariant())) p++;
                    if (p >= statement.Count) continue;
                    string kind = statement[p].Text.ToLowerInvariant();
                    if (kind == "end" && procedures.Count > 0 && p + 1 < statement.Count && new[] { "sub", "function", "property" }.Contains(statement[p + 1].Text.ToLowerInvariant()))
                    { procedures.Last().EndLine = statement.Last().Line; continue; }
                    if (kind != "sub" && kind != "function" && kind != "property") continue;
                    int n = p + (kind == "property" ? 2 : 1);
                    if (n >= statement.Count) continue;
                    var token = statement[n];
                    string declaration = string.Join("\n", lines.Skip(statement[0].Line - 1).Take(statement.Last().Line - statement[0].Line + 1)).Replace("_\n", " ").Trim();
                    var symbol = new EditorSymbol { Name = token.Text, Module = source.Module, Kind = kind == "property" ? "Property" : "Procedure", Scope = "Module", Line = token.Line, Column = token.Column, EndLine = lines.Length, Declaration = declaration, Private = statement.Take(p).Any(t => t.Text.Equals("Private", StringComparison.OrdinalIgnoreCase)) };
                    var type = Regex.Match(declaration, @"\)\s+As\s+([\w.]+)", RegexOptions.IgnoreCase);
                    symbol.TypeName = type.Success ? type.Groups[1].Value : "Variant";
                    symbol.Conditional = conditionalDepth > 0;
                    var parameters = VbaDeclarationIndex.Read(declaration).Where(d => d.Kind == "Parameter").Select(d => d.Name + " As " + d.TypeName).ToArray();
                    symbol.Parameters = parameters; procedures.Add(symbol); result.Add(symbol);
                }
                foreach (var d in VbaDeclarationIndex.Read(source.Text))
                {
                    var procedure = procedures.FirstOrDefault(s => s.Name.Equals(d.Scope, StringComparison.OrdinalIgnoreCase) && d.Line >= s.Line && d.Line <= s.EndLine);
                    string declaration = lines[d.Line - 1].Trim();
                    result.Add(new EditorSymbol { Name = d.Name, Module = source.Module, Kind = d.Kind, Scope = d.Scope, TypeName = d.TypeName, Line = d.Line, Column = d.Column, EndLine = procedure?.EndLine ?? lines.Length, Conditional = d.Conditional, Declaration = declaration,
                        Private = Regex.IsMatch(declaration, @"^(Private|Dim)\b", RegexOptions.IgnoreCase) });
                }
            }
            return result.ToArray();
        }
    }
}
