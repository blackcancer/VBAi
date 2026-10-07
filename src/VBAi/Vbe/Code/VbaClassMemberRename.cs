using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace VBAi
{

    /// <summary>Lie les références internes explicites de membres privés de classe à une déclaration unique.</summary>
    internal static class VbaClassMemberRename
    {

        /// <summary>Signature fermée, position physique et visibilité réellement écrite.</summary>
        private sealed class Member
        {

            /// <summary>Nom et catégorie syntaxique du membre.</summary>
            internal string Name, Kind, Access;

            /// <summary>Type d’accesseur VBIDE et dernière ligne de sa déclaration fermée.</summary>
            internal int ProcKind, EndLine;

            /// <summary>Jeton lexical qui porte le nom dans la source.</summary>
            internal VbaDeclarationIndex.Token Token;
        }

        /// <summary>Prépare un plan limité aux liaisons privées directes ou Me, sans déduire un type de receveur.</summary>
        /// <param name="project">Nom canonique du projet inspecté.</param>
        /// <param name="modules">Snapshots complets des composants et de leurs sources.</param>
        /// <param name="request">Classe, déclaration exacte, ancien nom, nouveau nom et empreinte attendue.</param>
        /// <returns>Plan sans mutation avec les positions privées résolues.</returns>
        /// <exception cref="ArgumentException">Le projet, les snapshots ou les identifiants de nom sont invalides.</exception>
        /// <exception cref="InvalidOperationException">La liaison est publique, ambiguë, masquée ou sort du périmètre privé pris en charge.</exception>
        internal static VbaProcedureRename.Plan Prepare(string project, IEnumerable<VbaProcedureRename.ModuleSnapshot> modules, Request request)
        {
            if (request == null || string.IsNullOrWhiteSpace(project) || modules == null)
                throw new ArgumentException("Project, complete snapshots and an exact declaration are required.");
            VbaTextEdits.ValidateIdentifier(request.Query); VbaTextEdits.ValidateIdentifier(request.NewName);
            if (Regex.IsMatch(request.NewName, @"^(?:Boolean|Byte|Integer|Long|LongLong|LongPtr|Single|Double|Currency|Date|String|Variant|Object|Type|Enum|Declare|PtrSafe|Optional|ParamArray|WithEvents|ReDim|Preserve|Erase|Stop|Debug|Print|GoTo|GoSub|Resume|Error|On|Until|Wend|To|Step|Each|And|Or|Xor|Not|Is|Like|Mod|Implements|RaiseEvent|Event|AddressOf|DefBool|DefByte|DefInt|DefLng|DefSng|DefDbl|DefCur|DefDate|DefStr|DefObj|DefVar)$", RegexOptions.IgnoreCase))
                throw new ArgumentException("The replacement is a reserved VBA word.");
            var all = modules.ToArray();
            if (all.Length == 0 || all.Length > 1000 || all.Any(x => x == null || string.IsNullOrWhiteSpace(x.Name) || x.Source == null) ||
                all.GroupBy(x => x.Name, StringComparer.OrdinalIgnoreCase).Any(g => g.Count() != 1))
                throw new InvalidOperationException("The complete project component snapshot is unreadable or ambiguous.");
            var target = all.SingleOrDefault(x => Same(x.Name, request.Module));
            if (target == null || target.Type != 2)
                throw new InvalidOperationException("An ordinary class module (VBIDE Type=2) is required; forms and host components are excluded.");
            if (!Same(VbaProcedureRename.Digest(target.Source), request.ExpectedSha256))
                throw new InvalidOperationException("The exact class SHA changed since inspection.");
            if (request.Query.Contains("_") || request.NewName.Contains("_") ||
                new[] { "main", "autoopen", "autoclose" }.Any(x => Same(x, request.Query) || Same(x, request.NewName)))
                throw new InvalidOperationException("Conventional event/interface callbacks require a separate consumer plan.");
            if (all.Any(x => Same(x.Name, request.Query) || Same(x.Name, request.NewName)) || Same(project, request.Query) || Same(project, request.NewName))
                throw new InvalidOperationException("The member collides with a component or project identity.");
            var syntax = all.ToDictionary(x => x.Name, x => VbaDeclarationIndex.Statements(x.Source).ToArray(), StringComparer.OrdinalIgnoreCase);
            foreach (var module in all) Guard(module, syntax[module.Name]);
            var members = ReadMembers(syntax[target.Name]);
            var family = members.Where(x => Same(x.Name, request.Query)).ToArray();
            var chosen = family.Where(x => x.Token.Line == request.StartLine && x.Token.Column == request.StartColumn && x.ProcKind == request.ProcKind).ToArray();
            if (chosen.Length != 1 || family.Any(x => x.Access != "Private"))
                throw new InvalidOperationException("Select the exact explicit Private class member. Public/Friend and implicit Public require a typed consumer plan.");
            bool property = chosen[0].Kind == "Property";
            if (family.Any(x => (x.Kind == "Property") != property) || (!property && family.Length != 1) ||
                (property && family.GroupBy(x => x.ProcKind).Any(g => g.Count() != 1)))
                throw new InvalidOperationException("The member or property accessor family is ambiguous.");
            if (!Same(request.Query, request.NewName) && members.Any(x => Same(x.Name, request.NewName)))
                throw new InvalidOperationException("The replacement already names another class member.");
            if (VbaDeclarationIndex.Read(target.Source).Any(x => Same(x.Name, request.Query) || Same(x.Name, request.NewName) || Same(x.Name, "Me")))
                throw new InvalidOperationException("A class field, parameter, local, type or external declaration shadows the binding.");
            var lineOffsets = all.ToDictionary(x => x.Name, x => Offsets(x.Source), StringComparer.OrdinalIgnoreCase);
            var positions = new HashSet<int>();
            foreach (var module in all)
                foreach (var statement in syntax[module.Name])
                    for (int i = 0; i < statement.Count; i++)
                    {
                        var token = statement[i];
                        if (module == target && !Same(request.Query, request.NewName) && Same(Bare(token.Text), request.NewName))
                            throw new InvalidOperationException("The replacement occurs in the class and its existing binding is not proven.");
                        if (!Same(Bare(token.Text), request.Query)) continue;
                        int offset = lineOffsets[module.Name][token.Line - 1] + token.Column - 1;
                        if (Excluded(statement, i, module.Source, offset)) continue;
                        if (module != target)
                        {
                            if (IsMember(statement, i)) throw new InvalidOperationException("An external qualified receiver has unresolved member binding.");
                            continue;
                        }
                        if (!Same(token.Text, request.Query)) throw new InvalidOperationException("Suffix or bang binding is not qualified.");
                        if (family.Any(x => x.Token.Line == token.Line && x.Token.Column == token.Column)) { positions.Add(offset); continue; }
                        bool qualified = IsMember(statement, i);
                        int start = i;
                        if (qualified)
                        {
                            if (i < 2 || statement[i - 1].Text != "." || !Same(statement[i - 2].Text, "Me") ||
                                (i > 2 && IsMember(statement, i - 2)))
                                throw new InvalidOperationException("Only direct Me qualification proves this private member receiver.");
                            start = i - 2;
                        }
                        var owner = members.SingleOrDefault(x => token.Line > x.Token.Line && token.Line < x.EndLine) ?? throw new InvalidOperationException("The reference is outside a closed class procedure.");
                        bool statementStart = start == 0 || new[] { "call", "then", "else" }.Contains(statement[start - 1].Text.ToLowerInvariant());
                        bool assignmentStart = start == 0 || new[] { "let", "set", "then", "else" }.Contains(statement[start - 1].Text.ToLowerInvariant());
                        bool assignment = assignmentStart && i + 1 < statement.Count && statement[i + 1].Text == "=";
                        if (chosen[0].Kind == "Sub" && (!statementStart || assignment))
                            throw new InvalidOperationException("The Sub reference is not a resolved direct call.");
                        if (chosen[0].Kind == "Function" && assignment && (qualified || owner != chosen[0]))
                            throw new InvalidOperationException("A Function name can be assigned only as its own return value.");
                        positions.Add(offset);
                    }
            var output = new StringBuilder(target.Source);
            foreach (int position in positions.OrderByDescending(x => x))
            { output.Remove(position, request.Query.Length); output.Insert(position, request.NewName); }
            var edits = new List<VbaProcedureRename.ModuleEdit>();
            if (!string.Equals(target.Source, output.ToString(), StringComparison.Ordinal))
                edits.Add(new VbaProcedureRename.ModuleEdit(target.Name, target.Source, output.ToString(), positions.Count));
            return new VbaProcedureRename.Plan(VbaProcedureRename.Version(project, all), edits);
        }

        /// <summary>Refuse toute construction qui peut cacher une liaison ou un consommateur dynamique.</summary>
        /// <param name="module">Module de classe analysé.</param>
        /// <param name="statements">Instructions lexicales du module.</param>
        /// <exception cref="InvalidOperationException">Une déclaration implicite, conditionnelle, dynamique ou non résolue empêche le renommage sûr.</exception>
        private static void Guard(VbaProcedureRename.ModuleSnapshot module, IList<List<VbaDeclarationIndex.Token>> statements)
        {
            if (statements.All(x => x.Count == 0)) return;
            if (!statements.Any(x => x.Count == 2 && Same(x[0].Text, "Option") && Same(x[1].Text, "Explicit")))
                throw new InvalidOperationException("Every nonempty component requires Option Explicit.");
            var offsets = Offsets(module.Source);
            foreach (var token in statements.SelectMany(x => x))
            {
                if (token.Text == "#" || Regex.IsMatch(token.Text, @"^(?:Def(?:Bool|Byte|Int|Lng|LngLng|LngPtr|Sng|Dbl|Cur|Date|Str|Obj|Var)|Implements|Attribute|AddressOf|CallByName|Run|OnAction|OnTime|Evaluate|With|WithEvents|RaiseEvent)$", RegexOptions.IgnoreCase))
                    throw new InvalidOperationException("Conditional, attribute, interface, callback, dynamic or With binding is not qualified.");
                if (token.Text == "<literal>" && module.Source[offsets[token.Line - 1] + token.Column - 1] == '[')
                    throw new InvalidOperationException("Bracket expressions hide unresolved consumers.");
            }
        }

        /// <summary>Lit les signatures explicites et chaque accesseur avec son terminateur exact.</summary>
        /// <param name="statements">Instructions lexicales du module de classe.</param>
        /// <returns>Membres fermés avec visibilité, type et jeton du nom.</returns>
        /// <exception cref="InvalidOperationException">Une signature est illisible, imbriquée ou sans terminateur correspondant.</exception>
        private static Member[] ReadMembers(IList<List<VbaDeclarationIndex.Token>> statements)
        {
            var result = new List<Member>(); Member active = null;
            foreach (var tokens in statements)
            {
                if (tokens.Count < 2) continue;
                if (Same(tokens[0].Text, "End") && new[] { "sub", "function", "property" }.Contains(tokens[1].Text.ToLowerInvariant()))
                {
                    if (active == null || !Same(active.Kind, tokens[1].Text)) throw new InvalidOperationException("A procedure terminator is unmatched.");
                    active.EndLine = tokens[1].Line; active = null; continue;
                }
                int first = 0; string access = "Public";
                while (first < tokens.Count && new[] { "public", "private", "friend", "static" }.Contains(tokens[first].Text.ToLowerInvariant()))
                { if (!Same(tokens[first].Text, "Static")) access = tokens[first].Text; first++; }
                if (first >= tokens.Count || !new[] { "sub", "function", "property" }.Contains(tokens[first].Text.ToLowerInvariant())) continue;
                if (active != null) throw new InvalidOperationException("Nested or unterminated class procedures are unsupported.");
                bool property = Same(tokens[first].Text, "Property"); int name = first + (property ? 2 : 1), kind = 0;
                if (name >= tokens.Count) throw new InvalidOperationException("The class member declaration is unreadable.");
                if (property)
                {
                    string accessor = tokens[first + 1].Text.ToLowerInvariant();
                    kind = accessor == "get" ? 3 : accessor == "let" ? 1 : accessor == "set" ? 2 : -1;
                    if (kind < 0) throw new InvalidOperationException("The property accessor kind is unreadable.");
                }
                active = new Member
                {
                    Name = Bare(tokens[name].Text),
                    Token = tokens[name],
                    ProcKind = kind,
                    Kind = property ? "Property" : Same(tokens[first].Text, "Sub") ? "Sub" : "Function",
                    Access = Same(access, "Private") ? "Private" : access
                };
                result.Add(active);
            }
            if (active != null) throw new InvalidOperationException("The matching End statement is required.");
            return result.ToArray();
        }

        /// <summary>Exclut types, labels, sauts et noms des arguments nommés.</summary>
        /// <param name="tokens">Jetons de l’instruction.</param>
        /// <param name="i">Index du jeton candidat.</param>
        /// <param name="source">Source exacte du module.</param>
        /// <param name="offset">Offset de début du jeton candidat.</param>
        /// <returns><see langword="true"/> si le jeton doit être exclu des usages de membre.</returns>
        private static bool Excluded(IList<VbaDeclarationIndex.Token> tokens, int i, string source, int offset)
        {
            if (i > 0 && new[] { "as", "new", "goto", "gosub", "resume" }.Contains(tokens[i - 1].Text.ToLowerInvariant())) return true;
            int next = offset + tokens[i].Text.Length;
            while (next < source.Length && (source[next] == ' ' || source[next] == '\t')) next++;
            return next < source.Length && source[next] == ':' && (i == 0 && tokens.Count == 1 || next + 1 < source.Length && source[next + 1] == '=');
        }

        /// <summary>Reconnaît point, leading-dot et notation bang sans inventer le receveur.</summary>
        /// <param name="tokens">Jetons de l’instruction.</param>
        /// <param name="i">Index du nom candidat.</param>
        /// <returns><see langword="true"/> si le nom est utilisé comme membre qualifié.</returns>
        private static bool IsMember(IList<VbaDeclarationIndex.Token> tokens, int i) => i > 0 && (tokens[i - 1].Text == "." || tokens[i - 1].Text.EndsWith("!", StringComparison.Ordinal));

        /// <summary>Compare les noms VBA indépendamment de leur casse.</summary>
        /// <param name="left">Premier nom.</param>
        /// <param name="right">Second nom.</param>
        /// <returns><see langword="true"/> si les identifiants sont égaux sans sensibilité à la casse.</returns>
        private static bool Same(string left, string right) => string.Equals(left, right, StringComparison.OrdinalIgnoreCase);

        /// <summary>Permet de repérer puis refuser les suffixes non qualifiés.</summary>
        /// <param name="text">Nom lexical d’un jeton.</param>
        /// <returns>Nom débarrassé de son suffixe de type éventuel.</returns>
        private static string Bare(string text) => text.TrimEnd('$', '%', '&', '!', '#', '@', '^');

        /// <summary>Offsets physiques sans normalisation de la source.</summary>
        /// <param name="source">Source avec ses fins de ligne d’origine.</param>
        /// <returns>Offsets absolus de chaque début de ligne.</returns>
        private static List<int> Offsets(string source)
        { var result = new List<int> { 0 }; for (int i = 0; i < source.Length; i++) if (source[i] == '\n') result.Add(i + 1); return result; }
    }
}
