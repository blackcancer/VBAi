using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CodexVBE
{
    /// <summary>Renomme un paramètre d'une procédure privée standard et les arguments nommés de ses appels locaux.</summary>
    internal static class VbaParameterRename
    {
        /// <summary>Prépare une édition atomique du module entier sans modifier les appels de membres homonymes.</summary>
        /// <param name="source">Source vivante complète du module standard.</param>
        /// <param name="request">Déclaration de paramètre, module et procédure identifiés.</param>
        /// <param name="first">Première ligne de la déclaration VBIDE.</param>
        /// <param name="last">Dernière ligne de la procédure VBIDE.</param>
        /// <returns>Source avec déclaration, usages locaux et arguments nommés liés modifiés.</returns>
        internal static string Transform(string source, Request request, int first, int last)
        {
            var statements = VbaDeclarationIndex.Statements(source).ToArray();
            var header = statements.FirstOrDefault(s => s.Count > 0 && s[0].Line == first);
            if (header == null || header.Count < 3 || !Same(header[0].Text, "Private") ||
                (!Same(header[1].Text, "Sub") && !Same(header[1].Text, "Function")) ||
                !Same(Bare(header[2].Text), request.Procedure) || request.ProcKind != 0)
                throw new InvalidOperationException("Parameter renaming requires an exact Private Sub/Function in a standard module.");
            var declarations = VbaDeclarationIndex.Read(source);
            if (declarations.Any(d => Same(d.Name, request.Procedure) || Same(d.Name, request.Module)))
                throw new InvalidOperationException("The procedure or module name is shadowed by a declaration; call binding is ambiguous.");
            // A conditional/repeated declaration can bind a caller to a different signature.
            int signatures = statements.Count(s => s.Any(t => Same(t.Text, "Sub") || Same(t.Text, "Function")) &&
                s.Any(t => Same(Bare(t.Text), request.Procedure)));
            if (signatures != 1 || statements.Any(s => s.Count > 0 && s[0].Text == "#"))
                throw new InvalidOperationException("Conditional or ambiguous procedure signatures need a project-wide refactoring plan.");
            string locallyRenamed = VbaLocalRename.TransformParameter(source, request, first, last);
            var offsets = new List<int> { 0 };
            for (int i = 0; i < locallyRenamed.Length; i++) if (locallyRenamed[i] == '\n') offsets.Add(i + 1);
            var edits = new HashSet<int>();
            foreach (var statement in VbaDeclarationIndex.Statements(locallyRenamed))
            {
                if (statement.Any(t => Same(t.Text, "Sub") || Same(t.Text, "Function") || Same(t.Text, "Property"))) continue;
                for (int call = 0; call < statement.Count; call++)
                {
                    if (!Same(Bare(statement[call].Text), request.Procedure)) continue;
                    if (call > 0 && (statement[call - 1].Text == "." || statement[call - 1].Text.EndsWith("!", StringComparison.Ordinal)))
                    {
                        if (call < 2 || !Same(statement[call - 2].Text, request.Module)) continue;
                        if (call >= 3 && statement[call - 3].Text == ".")
                            throw new InvalidOperationException("Project-qualified private calls require a project-wide binding analysis.");
                    }
                    int start = call + 1;
                    if (start >= statement.Count) continue;
                    bool parenthesized = statement[start].Text == "(";
                    if (!parenthesized && !(call == 0 || Same(statement[call - 1].Text, "Call") ||
                        Same(statement[call - 1].Text, "Then") || Same(statement[call - 1].Text, "Else") ||
                        (call >= 2 && statement[call - 1].Text == "." && Same(statement[call - 2].Text, request.Module)))) continue;
                    int depth = 0;
                    if (parenthesized) start++;
                    for (int i = start; i < statement.Count; i++)
                    {
                        string word = statement[i].Text;
                        if (word == "(") { depth++; continue; }
                        if (word == ")") { if (depth == 0) break; depth--; continue; }
                        if (!parenthesized && depth == 0 && Same(word, "Else")) break;
                        if (depth == 0 && i + 2 < statement.Count && Same(Bare(word), request.Query) &&
                            statement[i + 1].Text == ":" && statement[i + 2].Text == "=")
                            edits.Add(offsets[statement[i].Line - 1] + statement[i].Column - 1);
                    }
                }
            }
            var output = new StringBuilder(locallyRenamed);
            foreach (int offset in edits.OrderByDescending(x => x))
            { output.Remove(offset, request.Query.Length); output.Insert(offset, request.NewName); }
            return output.ToString();
        }

                /// <summary>Compare deux noms VBA sans distinction de casse.</summary>
                /// <param name="left">Premier nom.</param>
                /// <param name="right">Second nom.</param>
                /// <returns><see langword="true"/> si les noms sont égaux sans tenir compte de la casse.</returns>
        private static bool Same(string left, string right) => string.Equals(left, right, StringComparison.OrdinalIgnoreCase);
                /// <summary>Retire le suffixe de type d’un identifiant lexical.</summary>
                /// <param name="text">Jeton VBA.</param>
                /// <returns>Identifiant sans son suffixe de type.</returns>
        private static string Bare(string text) => text.TrimEnd('$', '%', '&', '!', '#', '@', '^');
    }
}
