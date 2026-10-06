using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace VBAi
{

    /// <summary>Instantané du code et du type d’un composant VBA.</summary>
    internal sealed class EditorSource
    {

        /// <summary>Nom du composant dont la source est capturée.</summary>
        /// <value>Nom VBComponent.Name.</value>
        public string Module { get; set; }

        /// <summary>Code source du composant.</summary>
        /// <value>Texte lu depuis son CodeModule.</value>
        public string Text { get; set; }

        /// <summary>Type numérique du composant VBIDE.</summary>
        /// <value>Valeur VBComponent.Type.</value>
        public int ComponentType { get; set; }
    }

    /// <summary>Symbole déclaré ou déduit d’une source VBA indexée par l’éditeur.</summary>
    internal sealed class EditorSymbol
    {

        /// <summary>Nom du symbole.</summary>
        /// <value>Nom du module, membre ou paramètre.</value>
        public string Name { get; set; }

        /// <summary>Module qui contient le symbole.</summary>
        /// <value>Nom du composant source.</value>
        public string Module { get; set; }

        /// <summary>Catégorie de déclaration.</summary>
        /// <value>Par exemple Module, Procedure ou Property.</value>
        public string Kind { get; set; }

        /// <summary>Portée dans laquelle le symbole est déclaré.</summary>
        /// <value>Nom de portée produit par l’analyseur VBA.</value>
        public string Scope { get; set; }

        /// <summary>Type VBA déduit de la déclaration.</summary>
        /// <value>Nom du type ou Variant lorsque la déclaration ne précise pas de type.</value>
        public string TypeName { get; set; }

        /// <summary>Texte de déclaration affiché pour le symbole.</summary>
        /// <value>Déclaration complète pour une procédure, ou ligne source pour les autres symboles.</value>
        public string Declaration { get; set; }

        /// <summary>Indique si la déclaration est privée ou locale.</summary>
        /// <value>Résultat de l’analyse de ses modificateurs de portée.</value>
        public bool Private { get; set; }

        /// <summary>Indique que le symbole se trouve dans une compilation conditionnelle.</summary>
        /// <value>État conditionnel relevé dans sa source.</value>
        public bool Conditional { get; set; }

        /// <summary>Indique que la déclaration est une référence externe.</summary>
        /// <value>Valeur fournie par l’index qui construit le symbole.</value>
        public bool External { get; set; }

        /// <summary>Gets or sets the library.</summary>
        /// <value>Current library exposed by editor symbol.</value>
        public string Library { get; set; }

        /// <summary>Description IntelliSense du membre fournie par la bibliothèque de types.</summary>
        /// <value>Documentation native du membre, éventuellement vide.</value>
        public string Documentation { get; set; }

        /// <summary>Description de la bibliothèque qui définit le symbole externe.</summary>
        /// <value>Texte de documentation de la référence chargée.</value>
        public string LibraryDescription { get; set; }

        /// <summary>Fichier de la référence chargée dont les métadonnées ont été lues.</summary>
        /// <value>Chemin local de la bibliothèque, sans instanciation de ses objets.</value>
        public string LibraryPath { get; set; }

        /// <summary>Fichier d’aide déclaré par la bibliothèque pour ce membre.</summary>
        /// <value>Chemin informatif ; aucune ouverture automatique.</value>
        public string HelpFile { get; set; }

        /// <summary>Identifiant de rubrique d’aide du membre.</summary>
        /// <value>Contexte natif, ou zéro si aucune rubrique n’est déclarée.</value>
        public int HelpContext { get; set; }

        /// <summary>Indique le membre appelé implicitement pour indexer une collection.</summary>
        /// <value>Valeur du DISPID zéro ou du drapeau de liaison par défaut.</value>
        public bool DefaultMember { get; set; }

        /// <summary>Indique un membre public d’un module de bibliothèque accessible sans qualification.</summary>
        /// <value>Vrai pour les fonctions et constantes des modules statiques.</value>
        public bool Global { get; set; }

        /// <summary>Première ligne de la déclaration.</summary>
        /// <value>Numéro de ligne indexé à partir de un.</value>
        public int Line { get; set; }

        /// <summary>Dernière ligne du symbole ou du corps de procédure qui le contient.</summary>
        /// <value>Numéro de ligne indexé à partir de un.</value>
        public int EndLine { get; set; }

        /// <summary>Colonne de début de la déclaration.</summary>
        /// <value>Position de colonne fournie par l’analyseur VBA.</value>
        public int Column { get; set; }

        /// <summary>Paramètres associés à une procédure.</summary>
        /// <value>Descriptions sous la forme « nom As type ».</value>
        public string[] Parameters { get; set; } = Array.Empty<string>();
    }

    /// <summary>Pure snapshot analysis. No COM or UI access; executed by the synchronization worker.</summary>
    internal static class EditorLanguageIndex
    {

        /// <summary>Construit les symboles des modules, procédures, propriétés et déclarations de leurs sources.</summary>
        /// <param name="sources">Instantanés de code lus depuis le projet.</param>
        /// <returns>Symboles avec leur emplacement, portée, type et déclaration source.</returns>
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
