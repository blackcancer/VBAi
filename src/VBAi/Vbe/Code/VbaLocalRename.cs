using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace VBAi
{

    /// <summary>Renommage lié à une déclaration locale explicite, dans une seule procédure VBA.</summary>
    internal static class VbaLocalRename
    {

        /// <summary>Prépare le texte sans changer membres qualifiés, types, labels, chaînes ni commentaires.</summary>
        /// <param name="source">Module vivant entier.</param>
        /// <param name="request">Déclaration exacte, nom attendu et nouveau nom.</param>
        /// <param name="first">Ligne de début de procédure obtenue de VBIDE.</param>
        /// <param name="last">Dernière ligne de sa plage VBIDE.</param>
        /// <returns>Texte du module après substitution des seules utilisations locales identifiées.</returns>
        internal static string Transform(string source, Request request, int first, int last)
            => TransformCore(source, request, first, last, false);

        /// <summary>Renomme le paramètre et ses utilisations dans la procédure ; les appelants sont traités séparément.</summary>
        /// <param name="source">Source VBA complète de la procédure et du module.</param>
        /// <param name="request">Déclaration de paramètre et nouveau nom demandé.</param>
        /// <param name="first">Première ligne de la procédure telle que fournie par VBIDE.</param>
        /// <param name="last">Dernière ligne de la plage VBIDE.</param>
        /// <returns>Source où la déclaration et les usages locaux du paramètre sont renommés.</returns>
        internal static string TransformParameter(string source, Request request, int first, int last)
            => TransformCore(source, request, first, last, true);

        /// <summary>Applique les contrôles lexicaux partagés à une déclaration locale ou de paramètre.</summary>
        /// <param name="source">Code VBA complet contenant la procédure.</param>
        /// <param name="request">Déclaration précise, ancien nom et nom de remplacement.</param>
        /// <param name="first">Première ligne de la procédure.</param>
        /// <param name="last">Dernière ligne de la plage VBIDE.</param>
        /// <param name="parameter">Indique si la cible attendue est un paramètre.</param>
        /// <returns>Source modifiée après substitution des références résolues localement.</returns>
        /// <exception cref="ArgumentException">Le nom ou la plage de déclaration est invalide.</exception>
        /// <exception cref="InvalidOperationException">La liaison est ambiguë ou la déclaration ne peut pas être renommée sûrement.</exception>
        private static string TransformCore(string source, Request request, int first, int last, bool parameter)
        {
            VbaTextEdits.ValidateIdentifier(request.NewName);
            if (Regex.IsMatch(request.NewName, @"^(?:Boolean|Byte|Integer|Long|LongLong|LongPtr|Single|Double|Currency|Date|String|Variant|Object|Type|Enum|Declare|PtrSafe|Optional|ParamArray|WithEvents|ReDim|Preserve|Erase|Stop|Debug|Print|GoTo|GoSub|Resume|Error|On|Until|Wend|To|Step|Each|And|Or|Xor|Not|Is|Like|Mod|Implements|RaiseEvent|Event|AddressOf|DefBool|DefByte|DefInt|DefLng|DefLngLng|DefLngPtr|DefSng|DefDbl|DefCur|DefDate|DefStr|DefObj|DefVar)$", RegexOptions.IgnoreCase))
                throw new ArgumentException("The replacement is a reserved VBA word.");
            if (first < 1 || last < first || request.StartLine < first || request.StartLine > last || request.StartColumn < 1)
                throw new ArgumentException("The declaration must belong to the inspected procedure range.");
            var declarations = VbaDeclarationIndex.Read(source);
            var matches = declarations.Where(x => x.Line == request.StartLine && x.Column == request.StartColumn &&
                x.Name.Equals(request.Query ?? "", StringComparison.OrdinalIgnoreCase)).ToArray();
            if (matches.Length != 1 || (parameter ? matches[0].Kind != "Parameter" : matches[0].Kind != "Variable" && matches[0].Kind != "Constant") ||
                !matches[0].Scope.Equals(request.Procedure ?? "", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Select one explicit local variable/constant declaration; parameters and project symbols require a wider refactoring plan.");
            var target = matches[0];
            var statements = VbaDeclarationIndex.Statements(source).ToArray();
            var header = statements.FirstOrDefault(x => x.Count > 0 && x[0].Line == first);
            if (header == null || !header.Any(x => x.Text.Equals("Sub", StringComparison.OrdinalIgnoreCase) || x.Text.Equals("Function", StringComparison.OrdinalIgnoreCase) || x.Text.Equals("Property", StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException("The VBIDE range must start at an exact procedure declaration.");
            string ending = header.Any(x => x.Text.Equals("Property", StringComparison.OrdinalIgnoreCase)) ? "Property" : header.Any(x => x.Text.Equals("Function", StringComparison.OrdinalIgnoreCase)) ? "Function" : "Sub";
            var endStatement = statements.FirstOrDefault(x => x.Count == 2 && x[0].Line > first && x[0].Line <= last && x[0].Text.Equals("End", StringComparison.OrdinalIgnoreCase) && x[1].Text.Equals(ending, StringComparison.OrdinalIgnoreCase));
            if (endStatement == null || request.StartLine >= endStatement[0].Line) throw new InvalidOperationException("The procedure's matching End statement is required.");
            last = endStatement[0].Line;
            if (target.Conditional || declarations.Count(x => x.Line >= first && x.Line <= last && x.Name.Equals(target.Name, StringComparison.OrdinalIgnoreCase)) != 1)
                throw new InvalidOperationException("Conditional or ambiguous local declarations cannot be renamed safely.");
            if (statements.Any(x => x.Count > 0 && Regex.IsMatch(x[0].Text, @"^Def(?:Bool|Byte|Int|Lng|LngLng|LngPtr|Sng|Dbl|Cur|Date|Str|Obj|Var)$", RegexOptions.IgnoreCase)) ||
                statements.Any(x => x.Count > 0 && x[0].Line >= first && x[0].Line <= last && x[0].Text == "#"))
                throw new InvalidOperationException("Conditional compilation and DefType modules need an explicit semantic analysis.");
            var offsets = new List<int> { 0 };
            for (int i = 0; i < source.Length; i++) if (source[i] == '\n') offsets.Add(i + 1);
            var replacements = new List<int>();
            bool sameName = target.Name.Equals(request.NewName, StringComparison.OrdinalIgnoreCase);
            foreach (var statement in statements)
            {
                for (int i = 0; i < statement.Count; i++)
                {
                    var token = statement[i];
                    if (token.Line < first || token.Line > last) continue;
                    string name = token.Text.TrimEnd('$', '%', '&', '!', '#', '@', '^');
                    if (!sameName && name.Equals(request.NewName, StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException("The new name already occurs in the procedure and could change binding.");
                    if (!name.Equals(target.Name, StringComparison.OrdinalIgnoreCase)) continue;
                    string previous = i == 0 ? "" : statement[i - 1].Text;
                    int offset = offsets[token.Line - 1] + token.Column - 1;
                    int following = offset + token.Text.Length;
                    while (following < source.Length && (source[following] == ' ' || source[following] == '\t')) following++;
                    bool member = previous == "." || previous == "!" || previous.EndsWith("!", StringComparison.Ordinal);
                    bool type = previous.Equals("AddressOf", StringComparison.OrdinalIgnoreCase) || previous.Equals("As", StringComparison.OrdinalIgnoreCase) || previous.Equals("New", StringComparison.OrdinalIgnoreCase) ||
                        (previous.Equals("Is", StringComparison.OrdinalIgnoreCase) && statement.Any(x => x.Text.Equals("TypeOf", StringComparison.OrdinalIgnoreCase)));
                    bool colon = following < source.Length && source[following] == ':';
                    bool labelOrNamedArgument = colon && ((i == 0 && statement.Count == 1) || (following + 1 < source.Length && source[following + 1] == '='));
                    bool labelTarget = previous.Equals("GoTo", StringComparison.OrdinalIgnoreCase) || previous.Equals("GoSub", StringComparison.OrdinalIgnoreCase) || previous.Equals("Resume", StringComparison.OrdinalIgnoreCase);
                    if (!member && !type && !labelOrNamedArgument && !labelTarget) replacements.Add(offset);
                }
            }
            int declarationOffset = offsets[target.Line - 1] + target.Column - 1;
            if (!replacements.Contains(declarationOffset)) throw new InvalidOperationException("The local declaration was not resolved in the procedure.");
            var output = new StringBuilder(source);
            foreach (int offset in replacements.OrderByDescending(x => x)) { output.Remove(offset, target.Name.Length); output.Insert(offset, request.NewName); }
            return output.ToString();
        }
    }
}
