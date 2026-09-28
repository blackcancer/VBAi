using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace CodexVBE
{
    /// <summary>Prépare un renommage conservateur de procédure standard dans un instantané complet du projet.</summary>
    internal static class VbaProcedureRename
    {
        /// <summary>Source et identité immuables d'un composant VBIDE inspecté.</summary>
        internal sealed class ModuleSnapshot
        {
                        /// <summary>Nom exact du composant.</summary>
                        /// <value>Nom utilisé pour associer les edits au module.</value>
            internal string Name { get; }
                        /// <summary>Type VBIDE du composant; 1 désigne un module standard.</summary>
                        /// <value>Valeur numérique VBComponent.Type.</value>
            internal int Type { get; }
                        /// <summary>Source entière observée.</summary>
                        /// <value>Contenu du CodeModule au moment du snapshot.</value>
            internal string Source { get; }
                        /// <summary>Construit une observation sans conserver de collection mutable de l'appelant.</summary>
                        /// <param name="name">Nom du composant.</param>
                        /// <param name="type">Type numérique VBIDE.</param>
                        /// <param name="source">Source complète du module.</param>
            internal ModuleSnapshot(string name, int type, string source)
            { Name = name; Type = type; Source = source; }
        }

        /// <summary>Édition complète préparée pour un module, avec empreinte de la source attendue.</summary>
        internal sealed class ModuleEdit
        {
                        /// <summary>Identité du composant.</summary>
                        /// <value>Nom du module à écrire.</value>
            public string Module { get; }
                        /// <summary>Source avant toute mutation.</summary>
                        /// <value>Texte source dont l’empreinte est ExpectedSha256.</value>
            public string Before { get; }
                        /// <summary>Source après remplacement des seules références résolues.</summary>
                        /// <value>Texte source proposé pour le module.</value>
            public string After { get; }
                        /// <summary>SHA256 de Before.</summary>
                        /// <value>Empreinte exigée avant d’appliquer cette édition.</value>
            public string ExpectedSha256 { get; }
                        /// <summary>Nombre de positions remplacées, déclaration incluse.</summary>
                        /// <value>Nombre de jetons substitués dans le module.</value>
            public int Replacements { get; }
                        /// <summary>Construit une édition immutable.</summary>
                        /// <param name="module">Nom du module concerné.</param>
                        /// <param name="before">Source avant renommage.</param>
                        /// <param name="after">Source après renommage.</param>
                        /// <param name="count">Nombre de références remplacées.</param>
            internal ModuleEdit(string module, string before, string after, int count)
            { Module = module; Before = before; After = after; Replacements = count; ExpectedSha256 = Digest(before); }
        }

        /// <summary>Plan immutable, sans lecture COM ni écriture dans l'éditeur.</summary>
        internal sealed class Plan
        {
                        /// <summary>Version du catalogue complet, des types et de toutes les sources.</summary>
                        /// <value>Empreinte de l’instantané complet utilisé pour préparer le plan.</value>
            public string SourceVersion { get; }
                        /// <summary>Éditions calculées; les modules inchangés sont absents.</summary>
                        /// <value>Collection en lecture seule des modifications à appliquer.</value>
            public ReadOnlyCollection<ModuleEdit> Edits { get; }
                        /// <summary>Construit le plan en copiant sa collection d'éditions.</summary>
                        /// <param name="version">Empreinte de l’instantané source.</param>
                        /// <param name="edits">Éditions calculées.</param>
            internal Plan(string version, IList<ModuleEdit> edits)
            { SourceVersion = version; Edits = new ReadOnlyCollection<ModuleEdit>(edits.ToArray()); }
        }

        /// <summary>Signature syntaxique et plage fermée d'une procédure.</summary>
        private sealed class Signature
        {
            /// <summary>Nom et catégorie d’une déclaration de procédure.</summary>
            internal string Name, Kind, Access;
            /// <summary>Jeton portant le nom dans la source.</summary>
            internal VbaDeclarationIndex.Token NameToken;
            /// <summary>Dernière ligne de la procédure complète.</summary>
            internal int EndLine;
        }

                /// <summary>Versionne toutes les sources, même les composants sans édition, indépendamment de l'ordre de collecte.</summary>
                /// <param name="project">Nom canonique du projet.</param>
                /// <param name="modules">Composants triés logiquement avant calcul de l’empreinte.</param>
                /// <returns>SHA-256 du nom de projet et de toutes les identités/types/sources.</returns>
        internal static string Version(string project, IEnumerable<ModuleSnapshot> modules)
        {
            var text = new StringBuilder(); Append(text, project ?? "");
            foreach (var module in modules.OrderBy(x => x.Name, StringComparer.OrdinalIgnoreCase))
            { Append(text, module.Name); Append(text, module.Type.ToString(System.Globalization.CultureInfo.InvariantCulture)); Append(text, module.Source); }
            return Digest(text.ToString());
        }

        /// <summary>Résout déclaration, appels directs et appels qualifiés vérifiés dans le seul projet fourni.</summary>
        /// <param name="project">Nom VBA du projet utilisé pour qualifier les appels.</param>
        /// <param name="modules">Tous les composants, sources complètes et types VBIDE.</param>
        /// <param name="request">Module, Query, NewName, StartLine/StartColumn exacts et SHA du module cible.</param>
        /// <returns>Plan sans mutation; les interfaces, callbacks et résolutions ambiguës sont refusés.</returns>
        internal static Plan Prepare(string project, IEnumerable<ModuleSnapshot> modules, Request request)
        {
            if (request == null || string.IsNullOrWhiteSpace(project) || modules == null)
                throw new ArgumentException("Project, complete module snapshots and request are required.");
            VbaTextEdits.ValidateIdentifier(request.Query); VbaTextEdits.ValidateIdentifier(request.NewName);
            if (Regex.IsMatch(request.NewName, @"^(?:Boolean|Byte|Integer|Long|LongLong|LongPtr|Single|Double|Currency|Date|String|Variant|Object|Type|Enum|Declare|PtrSafe|Optional|ParamArray|WithEvents|ReDim|Preserve|Erase|Stop|Debug|Print|GoTo|GoSub|Resume|Error|On|Until|Wend|To|Step|Each|And|Or|Xor|Not|Is|Like|Mod|Implements|RaiseEvent|Event|AddressOf|DefBool|DefByte|DefInt|DefLng|DefSng|DefDbl|DefCur|DefDate|DefStr|DefObj|DefVar)$", RegexOptions.IgnoreCase))
                throw new ArgumentException("The replacement is a reserved VBA word.");
            var snapshots = modules.ToArray();
            if (snapshots.Length == 0 || snapshots.Length > 1000 || snapshots.Any(x => x == null || string.IsNullOrWhiteSpace(x.Name) || x.Source == null) ||
                snapshots.GroupBy(x => x.Name, StringComparer.OrdinalIgnoreCase).Any(x => x.Count() != 1))
                throw new InvalidOperationException("The complete component catalogue is empty, duplicate or unreadable.");
            var targetModule = snapshots.SingleOrDefault(x => Same(x.Name, request.Module));
            if (targetModule == null || targetModule.Type != 1 || request.ProcKind != 0)
                throw new InvalidOperationException("Select a Sub/Function in one standard module; properties and class interfaces are unsupported.");
            if (string.IsNullOrWhiteSpace(request.ExpectedSha256) || !Same(Digest(targetModule.Source), request.ExpectedSha256))
                throw new InvalidOperationException("The target module SHA changed since inspection.");
            if (request.Query.Contains("_") || request.NewName.Contains("_") ||
                new[] { "autoopen", "autoclose", "main" }.Contains(request.Query.ToLowerInvariant()) ||
                new[] { "autoopen", "autoclose", "main" }.Contains(request.NewName.ToLowerInvariant()))
                throw new InvalidOperationException("Host entry points, event/interface callbacks and conventional callback names require an explicit consumer plan.");
            if (snapshots.Any(x => Same(x.Name, request.Query) || Same(x.Name, request.NewName)) || Same(project, request.Query) || Same(project, request.NewName))
                throw new InvalidOperationException("The procedure name collides with a project or component identity.");
            var syntax = snapshots.ToDictionary(x => x.Name, x => VbaDeclarationIndex.Statements(x.Source).ToArray(), StringComparer.OrdinalIgnoreCase);
            var signatures = snapshots.ToDictionary(x => x.Name, x => Signatures(syntax[x.Name]), StringComparer.OrdinalIgnoreCase);
            var targetMatches = signatures[targetModule.Name].Where(x => Same(x.Name, request.Query) &&
                x.NameToken.Line == request.StartLine && x.NameToken.Column == request.StartColumn).ToArray();
            if (targetMatches.Length != 1 || (targetMatches[0].Kind != "Sub" && targetMatches[0].Kind != "Function") ||
                (targetMatches[0].Access != "Public" && targetMatches[0].Access != "Private"))
                throw new InvalidOperationException("Select the exact unique Public/Private Sub or Function declaration.");
            var target = targetMatches[0];
            if (signatures.Values.SelectMany(x => x).Count(x => Same(x.Name, request.Query)) != 1 ||
                (!Same(request.Query, request.NewName) && signatures.Values.SelectMany(x => x).Any(x => Same(x.Name, request.NewName))))
                throw new InvalidOperationException("An existing procedure or property makes the rename ambiguous.");
            var edits = new List<ModuleEdit>();
            foreach (var module in snapshots)
            {
                var statements = syntax[module.Name];
                var offsets = LineOffsets(module.Source);
                GuardModule(module, statements, offsets, request, project);
                var positions = new HashSet<int>();
                foreach (var statement in statements)
                    for (int i = 0; i < statement.Count; i++)
                    {
                        var token = statement[i]; string word = Bare(token.Text);
                        if (!Same(request.Query, request.NewName) && Same(word, request.NewName) && !IsMember(statement, i))
                            throw new InvalidOperationException("The new name already occurs and could change an existing binding.");
                        if (!Same(word, request.Query)) continue;
                        int position = offsets[token.Line - 1] + token.Column - 1;
                        if (module == targetModule && token.Line == target.NameToken.Line && token.Column == target.NameToken.Column)
                        { positions.Add(position); continue; }
                        if (Excluded(statement, i, module.Source, position)) continue;
                        bool qualified = IsMember(statement, i);
                        if (qualified)
                        {
                            if (statement[i - 1].Text != "." || i < 2 || !Same(statement[i - 2].Text, targetModule.Name)) continue;
                            if (i >= 3 && statement[i - 3].Text == ".")
                            {
                                if (i < 4 || !Same(statement[i - 4].Text, project) || (i >= 5 && statement[i - 5].Text == "."))
                                    throw new InvalidOperationException("The module qualification belongs to an unresolved object or external project.");
                            }
                        }
                        if (target.Access == "Private" && module != targetModule)
                            throw new InvalidOperationException("A private procedure cannot bind an external-module caller.");
                        int start = qualified ? i - 2 : i;
                        if (qualified && i >= 4 && statement[i - 3].Text == ".") start = i - 4;
                        bool assignmentPosition = start == 0 || Same(statement[start - 1].Text, "Let") || Same(statement[start - 1].Text, "Set") ||
                            Same(statement[start - 1].Text, "Then") || Same(statement[start - 1].Text, "Else");
                        bool assignment = assignmentPosition && i + 1 < statement.Count && statement[i + 1].Text == "=";
                        bool returnAssignment = assignment && !qualified && module == targetModule && token.Line > target.NameToken.Line && token.Line < target.EndLine;
                        if (returnAssignment && target.Kind != "Function")
                            throw new InvalidOperationException("A Sub name cannot be assigned a return value.");
                        bool callStatement = start == 0 || Same(statement[start - 1].Text, "Call") || Same(statement[start - 1].Text, "Then") || Same(statement[start - 1].Text, "Else");
                        if (target.Kind == "Sub" && !callStatement)
                            throw new InvalidOperationException("The Sub reference is not a resolved direct call.");
                        if (!returnAssignment && assignment)
                            throw new InvalidOperationException("An assignment to the procedure name has ambiguous binding.");
                        positions.Add(position);
                    }
                var output = new StringBuilder(module.Source);
                foreach (int position in positions.OrderByDescending(x => x))
                { output.Remove(position, request.Query.Length); output.Insert(position, request.NewName); }
                if (!string.Equals(output.ToString(), module.Source, StringComparison.Ordinal))
                    edits.Add(new ModuleEdit(module.Name, module.Source, output.ToString(), positions.Count));
            }
            return new Plan(Version(project, snapshots), edits);
        }

                /// <summary>Refuse les constructs dont les consommateurs ou déclarations ne sont pas entièrement résolus.</summary>
                /// <param name="module">Snapshot de module analysé.</param>
                /// <param name="statements">Instructions lexicales du module.</param>
                /// <param name="offsets">Offsets de début de ligne dans la source.</param>
                /// <param name="request">Ancien nom et nom demandé.</param>
                /// <param name="project">Nom du projet dont les références sont résolues.</param>
                /// <exception cref="InvalidOperationException">Le code contient des consommateurs implicites, dynamiques ou ambigus.</exception>
        private static void GuardModule(ModuleSnapshot module, IList<List<VbaDeclarationIndex.Token>> statements, IList<int> offsets, Request request, string project)
        {
            if (statements.All(x => x.Count == 0)) return;
            if (!statements.Any(x => x.Count == 2 && Same(x[0].Text, "Option") && Same(x[1].Text, "Explicit")))
                throw new InvalidOperationException("Option Explicit is required in every component to reject implicit shadows.");
            foreach (var statement in statements)
                foreach (var token in statement)
                {
                    if (token.Text == "#" || Regex.IsMatch(token.Text, @"^(?:Def(?:Bool|Byte|Int|Lng|LngLng|LngPtr|Sng|Dbl|Cur|Date|Str|Obj|Var)|Implements|AddressOf|CallByName|Run|OnAction|OnTime|Evaluate)$", RegexOptions.IgnoreCase))
                        throw new InvalidOperationException("Conditional, implicit-type, interface, callback or dynamic consumers require an explicit semantic plan.");
                    if (token.Text == "<literal>" && module.Source[offsets[token.Line - 1] + token.Column - 1] == '[')
                        throw new InvalidOperationException("Bracket expressions can hide procedure consumers and are unsupported.");
                }
            if (VbaDeclarationIndex.Read(module.Source).Any(x => Same(x.Name, request.Query) || Same(x.Name, request.NewName) || Same(x.Name, request.Module) || Same(x.Name, project)))
                throw new InvalidOperationException("A variable, parameter, type, enum or external declaration shadows a binding used by this rename.");
        }

                /// <summary>Identifie les signatures et exige leur fermeture; ne traite pas les accesseurs comme des fonctions.</summary>
                /// <param name="statements">Instructions lexicales du composant.</param>
                /// <returns>Signatures Sub, Function et Property avec positions et accès.</returns>
                /// <exception cref="InvalidOperationException">Une signature est imbriquée, incomplète ou sans terminateur correspondant.</exception>
        private static Signature[] Signatures(IList<List<VbaDeclarationIndex.Token>> statements)
        {
            var signatures = new List<Signature>(); Signature active = null;
            foreach (var tokens in statements)
            {
                if (tokens.Count < 2) continue;
                if (Same(tokens[0].Text, "End") && new[] { "sub", "function", "property" }.Contains(tokens[1].Text.ToLowerInvariant()))
                {
                    if (active == null || !Same(active.Kind, tokens[1].Text)) throw new InvalidOperationException("A procedure terminator has no matching signature.");
                    active.EndLine = tokens[1].Line; active = null; continue;
                }
                int first = 0; string access = "Public";
                while (first < tokens.Count && new[] { "public", "private", "friend", "static" }.Contains(tokens[first].Text.ToLowerInvariant()))
                { if (!Same(tokens[first].Text, "Static")) access = tokens[first].Text; first++; }
                if (first >= tokens.Count || !new[] { "sub", "function", "property" }.Contains(tokens[first].Text.ToLowerInvariant())) continue;
                if (active != null) throw new InvalidOperationException("Nested or unterminated procedures are unsupported.");
                int name = first + (Same(tokens[first].Text, "Property") ? 2 : 1);
                if (name >= tokens.Count) throw new InvalidOperationException("A procedure name is unreadable.");
                active = new Signature { Name = Bare(tokens[name].Text), Kind = Same(tokens[first].Text, "Property") ? "Property" : Same(tokens[first].Text, "Sub") ? "Sub" : "Function",
                    Access = Same(access, "Public") ? "Public" : Same(access, "Private") ? "Private" : access, NameToken = tokens[name] };
                signatures.Add(active);
            }
            if (active != null) throw new InvalidOperationException("The matching End statement is required.");
            return signatures.ToArray();
        }

                /// <summary>Exclut types, labels, arguments nommés et cibles de saut qui ne sont pas des appels.</summary>
                /// <param name="tokens">Jetons de l’instruction source.</param>
                /// <param name="index">Index du jeton candidat.</param>
                /// <param name="source">Source entière permettant de lire le caractère suivant.</param>
                /// <param name="position">Offset physique du jeton dans la source.</param>
                /// <returns><see langword="true"/> si le jeton est un usage exclu du renommage.</returns>
        private static bool Excluded(IList<VbaDeclarationIndex.Token> tokens, int index, string source, int position)
        {
            string previous = index == 0 ? "" : tokens[index - 1].Text;
            if (new[] { "as", "new", "goto", "gosub", "resume" }.Contains(previous.ToLowerInvariant())) return true;
            int following = position + tokens[index].Text.Length;
            while (following < source.Length && (source[following] == ' ' || source[following] == '\t')) following++;
            return following < source.Length && source[following] == ':' &&
                ((index == 0 && tokens.Count == 1) || (following + 1 < source.Length && source[following + 1] == '='));
        }
                /// <summary>Reconnaît les accès de membres, y compris la notation bang.</summary>
                /// <param name="tokens">Jetons de l’instruction.</param>
                /// <param name="index">Index du nom candidat.</param>
                /// <returns><see langword="true"/> si le nom est qualifié comme membre.</returns>
        private static bool IsMember(IList<VbaDeclarationIndex.Token> tokens, int index) => index > 0 &&
            (tokens[index - 1].Text == "." || tokens[index - 1].Text.EndsWith("!", StringComparison.Ordinal));
                /// <summary>Retire uniquement les suffixes de type de nom VBA.</summary>
                /// <param name="name">Nom lexical potentiellement suffixé.</param>
                /// <returns>Nom sans suffixe final de type.</returns>
        private static string Bare(string name) => name.TrimEnd('$', '%', '&', '!', '#', '@', '^');
                /// <summary>Compare les identités VBA sans distinction de casse.</summary>
                /// <param name="left">Première identité.</param>
                /// <param name="right">Seconde identité.</param>
                /// <returns><see langword="true"/> si elles sont identiques sans sensibilité à la casse.</returns>
        private static bool Same(string left, string right) => string.Equals(left, right, StringComparison.OrdinalIgnoreCase);
                /// <summary>Calcule les offsets physiques des débuts de lignes.</summary>
                /// <param name="source">Texte source avec fins de ligne intactes.</param>
                /// <returns>Offsets de début de chaque ligne dans le texte.</returns>
        private static List<int> LineOffsets(string source)
        { var offsets = new List<int> { 0 }; for (int i = 0; i < source.Length; i++) if (source[i] == '\n') offsets.Add(i + 1); return offsets; }
                /// <summary>Encode une chaîne avec longueur pour éviter une concaténation de version ambiguë.</summary>
                /// <param name="output">Tampon recevant la valeur encadrée par sa longueur.</param>
                /// <param name="value">Texte à ajouter.</param>
        private static void Append(StringBuilder output, string value) => output.Append(value.Length).Append(':').Append(value);
                /// <summary>SHA256 du texte UTF8 exact, sans normalisation des lignes.</summary>
                /// <param name="source">Texte source dont les octets UTF-8 exacts sont empreintés.</param>
                /// <returns>SHA-256 hexadécimal minuscule.</returns>
        internal static string Digest(string source)
        { using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(source))).Replace("-", "").ToLowerInvariant(); }
    }
}
