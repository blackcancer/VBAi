using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;

namespace CodexVBE
{
    internal sealed partial class VbeDebug
    {
        /// <summary>Catalogue syntaxique des macros et procédures publiques standard, sans exécuter de code.</summary>
        public object ListMacros(Request request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Project) || request.Offset < 0 || request.Offset > 100000 ||
                request.Limit < 0 || request.Limit > 500 || (request.Query?.Length ?? 0) > 256)
                throw new ArgumentException("Project, Offset >= 0, Limit 0..500 and Query <= 256 characters are required.");
            dynamic project = GetProject(request.Project);
            var entries = new List<object>();
            var identities = new List<object>();
            foreach (dynamic component in project.VBComponents)
            {
                if ((int)component.Type != 1) continue;
                dynamic module = component.CodeModule;
                int count = (int)module.CountOfLines;
                if (count > 200000) throw new InvalidOperationException("A module exceeds the macro catalogue source limit.");
                string source = count == 0 ? "" : (string)module.Lines[1, count];
                string name = (string)component.Name, sha = Hash(source);
                identities.Add(new { Module = name, Sha256 = sha });
                var statements = VbaDeclarationIndex.Statements(source).Where(s => s.Count > 0).ToArray();
                var declarations = VbaDeclarationIndex.Read(source);
                bool optionPrivate = statements.Any(s => s.Count == 3 && s[0].Text.Equals("Option", StringComparison.OrdinalIgnoreCase) &&
                    s[1].Text.Equals("Private", StringComparison.OrdinalIgnoreCase) && s[2].Text.Equals("Module", StringComparison.OrdinalIgnoreCase));
                int conditional = 0;
                foreach (var statement in statements)
                {
                    if (statement[0].Text == "#")
                    {
                        if (statement.Count > 1 && statement[1].Text.Equals("If", StringComparison.OrdinalIgnoreCase)) conditional++;
                        if (statement.Count > 2 && statement[1].Text.Equals("End", StringComparison.OrdinalIgnoreCase) && statement[2].Text.Equals("If", StringComparison.OrdinalIgnoreCase)) conditional = Math.Max(0, conditional - 1);
                        continue;
                    }
                    int first = 0;
                    bool isPrivate = false;
                    while (first < statement.Count && new[] { "Public", "Private", "Friend", "Static", "Global" }.Contains(statement[first].Text, StringComparer.OrdinalIgnoreCase))
                    { isPrivate |= statement[first].Text.Equals("Private", StringComparison.OrdinalIgnoreCase) || statement[first].Text.Equals("Friend", StringComparison.OrdinalIgnoreCase); first++; }
                    if (isPrivate || first + 1 >= statement.Count ||
                        !new[] { "Sub", "Function" }.Contains(statement[first].Text, StringComparer.OrdinalIgnoreCase)) continue;
                    string procedure = statement[first + 1].Text.TrimEnd('$', '%', '&', '!', '#', '@');
                    var parameters = declarations.Where(d => d.Kind == "Parameter" && d.Line >= statement[0].Line && d.Line <= statement.Last().Line &&
                        d.Scope.Equals(procedure, StringComparison.OrdinalIgnoreCase)).ToArray();
                    bool macro = statement[first].Text.Equals("Sub", StringComparison.OrdinalIgnoreCase) && parameters.Length == 0 && !optionPrivate && conditional == 0;
                    if (!string.IsNullOrEmpty(request.Query) && (name + "." + procedure).IndexOf(request.Query, StringComparison.OrdinalIgnoreCase) < 0) continue;
                    entries.Add(new { Module = name, Procedure = procedure, Kind = statement[first].Text, Line = statement[0].Line,
                        Sha256 = sha, Parameters = parameters.Select(p => new { p.Name, p.TypeName }).ToArray(),
                        NativeMacroCandidate = macro, Conditional = conditional > 0, OptionPrivateModule = optionPrivate,
                        NativeVisibilityVerified = false, ExecutionTool = conditional == 0 ? "run_procedure" : null });
                    if (entries.Count > 10000) throw new InvalidOperationException("The macro catalogue exceeds 10000 procedures.");
                }
            }
            int limit = request.Limit == 0 ? 100 : request.Limit;
            return new { Project = request.Project, Mode = (int)project.Mode,
                CatalogVersion = Hash(new JavaScriptSerializer().Serialize(identities)),
                Macros = entries.Skip(request.Offset).Take(limit).ToArray(), Total = entries.Count, request.Offset, Limit = limit,
                HasMore = request.Offset + limit < entries.Count,
                Coverage = "Public standard-module Sub/Function declarations. NativeMacroCandidate is syntactic, not host chooser qualification." };
        }
    }
}
