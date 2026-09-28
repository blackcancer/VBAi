using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace CodexVBE
{
    /// <summary>Index syntaxique des déclarations VBA avec positions physiques, sans résolution de références COM.</summary>
    internal static class VbaDeclarationIndex
    {
        /// <summary>Déclaration navigable du texte source vivant.</summary>
        internal sealed class Declaration
        {
                        /// <summary>Nom déclaré, sans suffixe de type VBA.</summary>
                        /// <value>Identifiant nettoyé de son suffixe de type éventuel.</value>
            public string Name { get; set; }
                        /// <summary>Catégorie syntaxique de la déclaration.</summary>
                        /// <value>Type de membre tel que Variable, Constant, Procedure ou Field.</value>
            public string Kind { get; set; }
                        /// <summary>Nom de la procédure ou du type contenant la déclaration ; null au niveau module.</summary>
                        /// <value>Portée lexicale, ou <see langword="null"/> pour une déclaration de module.</value>
            public string Scope { get; set; }
                        /// <summary>Type écrit dans la déclaration ou déduit de son suffixe.</summary>
                        /// <value>Nom du type VBA ou Variant par défaut.</value>
            public string TypeName { get; set; }
                        /// <summary>Ligne physique à base un du nom déclaré.</summary>
                        /// <value>Numéro de ligne dans le texte source d’origine.</value>
            public int Line { get; set; }
                        /// <summary>Colonne physique à base un du nom déclaré.</summary>
                        /// <value>Position du premier caractère du nom.</value>
            public int Column { get; set; }
                        /// <summary>Indique une déclaration située dans une branche de compilation conditionnelle.</summary>
                        /// <value><see langword="true"/> lorsque la déclaration apparaît dans une branche #If.</value>
            public bool Conditional { get; set; }
        }
        /// <summary>Jeton lexical hors commentaires et littéraux avec ses coordonnées physiques.</summary>
        internal sealed class Token
        {
            /// <summary>Texte exact du jeton, y compris un suffixe de type.</summary>
            public string Text;
            /// <summary>Ligne et colonne physiques à base un.</summary>
            public int Line, Column;
        }
        /// <summary>Recense variables, constantes, paramètres, types, champs et membres d’énumération.</summary>
        /// <param name="source">Texte du module, chaînes/commentaires exclus des déclarations.</param>
        /// <returns>Déclarations dans leur ordre source, branches conditionnelles marquées.</returns>
        internal static Declaration[] Read(string source)
        {
            var result = new List<Declaration>();
            string procedure = null, container = null, containerKind = null;
            int conditionalDepth = 0;
            foreach (var tokens in Statements(source ?? ""))
            {
                if (tokens.Count == 0) continue;
                string Word(int i) => i < tokens.Count ? tokens[i].Text.ToLowerInvariant() : "";
                if (Word(0) == "#")
                {
                    if (Word(1) == "if") conditionalDepth++;
                    else if (Word(1) == "end" && Word(2) == "if") conditionalDepth = Math.Max(0, conditionalDepth - 1);
                    continue;
                }
                if (Word(0) == "end")
                {
                    if (Word(1) == "sub" || Word(1) == "function" || Word(1) == "property") procedure = null;
                    if (Word(1) == "type" || Word(1) == "enum") { container = null; containerKind = null; }
                    continue;
                }
                int first = 0;
                while (Word(first) == "public" || Word(first) == "private" || Word(first) == "friend" || Word(first) == "global" || Word(first) == "static") first++;
                string keyword = Word(first);
                if (keyword == "sub" || keyword == "function" || keyword == "property")
                {
                    int name = first + (keyword == "property" ? 2 : 1);
                    if (name >= tokens.Count || !IsName(tokens[name].Text)) continue;
                    procedure = Bare(tokens[name].Text);
                    int opening = tokens.FindIndex(name + 1, t => t.Text == "(");
                    if (opening >= 0)
                    {
                        int depth = 0, closing = -1;
                        // Locate the signature close independently from array/default-expression parentheses.
                        depth = 0;
                        for (int i = opening; i < tokens.Count; i++)
                        { if (tokens[i].Text == "(") depth++; else if (tokens[i].Text == ")" && --depth == 0) { closing = i; break; } }
                        if (closing > opening)
                            AddList(tokens.GetRange(opening + 1, closing - opening - 1), "Parameter", procedure, conditionalDepth > 0, result);
                    }
                    continue;
                }
                if (keyword == "event" || keyword == "declare")
                {
                    int name = first + 1;
                    if (keyword == "declare")
                    {
                        if (Word(name) == "ptrsafe") name++;
                        if (Word(name) != "sub" && Word(name) != "function") continue;
                        name++;
                    }
                    if (name < tokens.Count && IsName(tokens[name].Text))
                        Add(tokens[name], keyword == "event" ? "Event" : "ExternalProcedure", "Module", "", conditionalDepth > 0, result);
                    continue;
                }
                if (keyword == "type" || keyword == "enum")
                {
                    if (first + 1 < tokens.Count && IsName(tokens[first + 1].Text))
                    {
                        container = Bare(tokens[first + 1].Text); containerKind = keyword;
                        Add(tokens[first + 1], keyword == "type" ? "Type" : "Enum", "Module", "", conditionalDepth > 0, result);
                    }
                    continue;
                }
                if (container != null)
                {
                    if (IsName(tokens[0].Text)) AddList(tokens, containerKind == "type" ? "Field" : "EnumMember", container, conditionalDepth > 0, result);
                    continue;
                }
                bool isConst = keyword == "const";
                bool explicitVariable = keyword == "dim" || keyword == "withevents" || isConst;
                bool modifiedVariable = first > 0 && keyword != "declare" && keyword != "event";
                if (!explicitVariable && !modifiedVariable) continue;
                if (explicitVariable) first++;
                if (Word(first) == "withevents") first++;
                AddList(tokens.Skip(first).ToList(), isConst ? "Constant" : "Variable", procedure ?? "Module", conditionalDepth > 0, result);
            }
            return result.ToArray();
        }
                /// <summary>Découpe les déclarateurs en respectant tableaux et expressions parenthésées.</summary>
                /// <param name="tokens">Jetons formant la liste de déclarations.</param>
                /// <param name="kind">Catégorie affectée aux noms extraits.</param>
                /// <param name="scope">Portée qui contient ces noms.</param>
                /// <param name="conditional">Indique si la liste est dans une compilation conditionnelle.</param>
                /// <param name="result">Collection enrichie avec les déclarations trouvées.</param>
        private static void AddList(List<Token> tokens, string kind, string scope, bool conditional, List<Declaration> result)
        {
            int start = 0, depth = 0;
            for (int i = 0; i <= tokens.Count; i++)
            {
                if (i < tokens.Count && tokens[i].Text == "(") depth++;
                if (i < tokens.Count && tokens[i].Text == ")") depth--;
                if (i < tokens.Count && (tokens[i].Text != "," || depth != 0)) continue;
                var part = tokens.Skip(start).Take(i - start).ToList(); start = i + 1;
                int name = 0;
                while (name < part.Count && new[] { "optional", "byval", "byref", "paramarray", "withevents" }.Contains(part[name].Text.ToLowerInvariant())) name++;
                if (name >= part.Count || !IsName(part[name].Text)) continue;
                int asIndex = part.FindIndex(name + 1, t => t.Text.Equals("As", StringComparison.OrdinalIgnoreCase));
                string type = "Variant";
                if (asIndex >= 0)
                {
                    int from = asIndex + 1;
                    if (from < part.Count && part[from].Text.Equals("New", StringComparison.OrdinalIgnoreCase)) from++;
                    var names = part.Skip(from).TakeWhile(t => t.Text != "=" && t.Text != "*").Select(t => t.Text);
                    type = string.Concat(names);
                }
                else if (part[name].Text.Length > Bare(part[name].Text).Length)
                    type = SuffixType(part[name].Text.Last());
                Add(part[name], kind, scope, type, conditional, result);
            }
        }
        /// <summary>Ajoute un symbole en séparant son nom de son éventuel suffixe de type.</summary>
        /// <param name="token">Jeton portant le nom déclaré.</param>
        /// <param name="kind">Catégorie de déclaration.</param>
        /// <param name="scope">Portée du symbole.</param>
        /// <param name="type">Type déclaré ou inféré.</param>
        /// <param name="conditional">Indique une branche conditionnelle.</param>
        /// <param name="output">Collection de sortie.</param>
        private static void Add(Token token, string kind, string scope, string type, bool conditional, List<Declaration> output) =>
            output.Add(new Declaration { Name = Bare(token.Text), Kind = kind, Scope = scope, TypeName = type,
                Line = token.Line, Column = token.Column, Conditional = conditional });
        /// <summary>Retire un suffixe de type VBA éventuel du nom lexical.</summary>
        /// <param name="text">Jeton ou identifiant à normaliser.</param>
        /// <returns>Nom sans suffixe de type final.</returns>
        private static string Bare(string text) => text.TrimEnd('$', '%', '&', '!', '#', '@', '^');
        /// <summary>Vérifie la forme lexicale d’un identifiant VBA, suffixe de type facultatif compris.</summary>
        /// <param name="text">Texte du jeton à valider.</param>
        /// <returns><see langword="true"/> si le jeton correspond à un nom VBA reconnu.</returns>
        private static bool IsName(string text) => Regex.IsMatch(text, @"^\p{L}[\p{L}\p{N}_]*[$%&!#@^]?$");
        /// <summary>Convertit un caractère suffixe VBA en nom de type correspondant.</summary>
        /// <param name="suffix">Suffixe placé à la fin d’un identifiant.</param>
        /// <returns>Type VBA associé, ou Variant pour un suffixe non reconnu.</returns>
        private static string SuffixType(char suffix)
        {
            switch (suffix) { case '$': return "String"; case '%': return "Integer"; case '&': return "Long";
                case '!': return "Single"; case '#': return "Double"; case '@': return "Currency"; case '^': return "LongLong"; default: return "Variant"; }
        }
                /// <summary>Lexeur de déclarations : conserve les positions, les continuations et les séparateurs.</summary>
                /// <param name="source">Texte VBA à découper.</param>
                /// <returns>Listes de jetons séparées par des fins de ligne ou des séparateurs d’instruction.</returns>
        internal static IEnumerable<List<Token>> Statements(string source)
        {
            var statement = new List<Token>();
            int line = 1, column = 1;
            for (int i = 0; i < source.Length;)
            {
                char c = source[i];
                if (c == '\r') { i++; continue; }
                if (c == '\n')
                {
                    bool continued = statement.Count > 0 && statement.Last().Text == "_";
                    if (continued) statement.RemoveAt(statement.Count - 1);
                    else { yield return statement; statement = new List<Token>(); }
                    i++; line++; column = 1; continue;
                }
                if (char.IsWhiteSpace(c)) { i++; column++; continue; }
                if (c == ':' && (i + 1 >= source.Length || source[i + 1] != '='))
                { yield return statement; statement = new List<Token>(); i++; column++; continue; }
                int tokenLine = line, tokenColumn = column, start = i;
                if (c == '\'') { while (i < source.Length && source[i] != '\n') { i++; column++; } continue; }
                if (c == '"' || c == '[' || (c == '#' && statement.Count > 0))
                {
                    char close = c == '[' ? ']' : c;
                    i++; column++;
                    while (i < source.Length && source[i] != '\n')
                    {
                        char next = source[i++]; column++;
                        if (next != close) continue;
                        if (close == '"' && i < source.Length && source[i] == '"') { i++; column++; continue; }
                        break;
                    }
                    statement.Add(new Token { Text = "<literal>", Line = tokenLine, Column = tokenColumn }); continue;
                }
                if (char.IsLetter(c) || c == '_')
                {
                    i++; column++;
                    while (i < source.Length && (char.IsLetterOrDigit(source[i]) || source[i] == '_')) { i++; column++; }
                    if (i < source.Length && "$%&!#@^".Contains(source[i])) { i++; column++; }
                    string text = source.Substring(start, i - start);
                    if (statement.Count == 0 && text.Equals("Rem", StringComparison.OrdinalIgnoreCase))
                    { while (i < source.Length && source[i] != '\n') { i++; column++; } continue; }
                    statement.Add(new Token { Text = text, Line = tokenLine, Column = tokenColumn }); continue;
                }
                statement.Add(new Token { Text = c.ToString(), Line = tokenLine, Column = tokenColumn }); i++; column++;
            }
            if (statement.Count > 0) yield return statement;
        }
    }
}
