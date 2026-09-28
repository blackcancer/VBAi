using System;
using System.Linq;
using System.Text;

namespace CodexVBE
{
    /// <summary>Indique la nature d’une ligne dans une comparaison de code.</summary>
    internal enum CodeDiffKind
    {
        /// <summary>Ligne de contexte présente dans les deux versions.</summary>
        Context,
        /// <summary>Ligne présente dans la version initiale et supprimée dans la version modifiée.</summary>
        Removed,
        /// <summary>Ligne ajoutée dans la version modifiée.</summary>
        Added,
        /// <summary>Ligne informative qui délimite un bloc ou signale l’absence de différence.</summary>
        Notice
    }

    /// <summary>Décrit une ligne typée d’une comparaison de code, avec ses numéros de ligne éventuels.</summary>
    internal sealed class CodeDiffLine
    {
        /// <summary>Obtient ou définit la catégorie de la ligne.</summary>
        public CodeDiffKind Kind { get; set; }
        /// <summary>Obtient ou définit le numéro de ligne dans la version initiale, ou nul si absent.</summary>
        public int? OldLine { get; set; }
        /// <summary>Obtient ou définit le numéro de ligne dans la version modifiée, ou nul si absent.</summary>
        public int? NewLine { get; set; }
        /// <summary>Obtient ou définit le texte de la ligne sans son préfixe d’affichage.</summary>
        public string Text { get; set; }
        /// <summary>Obtient le signe « + » pour une ligne ajoutée, « − » pour une ligne supprimée, ou une chaîne vide sinon.</summary>
        public string Sign { get { return Kind == CodeDiffKind.Added ? "+" : Kind == CodeDiffKind.Removed ? "−" : ""; } }
    }

    /// <summary>Conserve dans une session l’avant et l’après d’une modification réellement appliquée au code VBIDE.</summary>
    internal sealed class CodeChange
    {
        /// <summary>Obtient ou définit le nom du projet modifié.</summary>
        public string Project { get; set; }
        /// <summary>Obtient ou définit le nom du module modifié.</summary>
        public string Module { get; set; }
        /// <summary>Obtient ou définit le code avant la modification.</summary>
        public string Before { get; set; }
        /// <summary>Obtient ou définit le code après la modification.</summary>
        public string After { get; set; }
        /// <summary>Obtient ou définit le SHA-256 du code initial.</summary>
        public string BeforeSha256 { get; set; }
        /// <summary>Obtient ou définit le SHA-256 du code modifié.</summary>
        public string AfterSha256 { get; set; }
        /// <summary>Obtient ou définit le nombre de lignes après modification.</summary>
        public int AfterLineCount { get; set; }
        /// <summary>Obtient ou définit la date et l’heure locale de la modification.</summary>
        public DateTime Time { get; set; }
        /// <summary>Obtient ou définit si l’intégralité de la modification a été restaurée.</summary>
        public bool Restored { get; set; }
        /// <summary>Obtient ou définit l’identifiant du tour ayant effectué la modification.</summary>
        public string TurnId { get; set; }
        /// <summary>Obtient ou définit les index des blocs déjà restaurés partiellement.</summary>
        public System.Collections.Generic.List<int> RestoredHunks { get; set; } = new System.Collections.Generic.List<int>();
        /// <summary>Obtient le diff textuel formaté à partir des versions conservées.</summary>
        [System.Web.Script.Serialization.ScriptIgnore] public string Diff { get { return FormatDiff(Before, After); } }
        /// <summary>Obtient les lignes structurées du diff.</summary>
        [System.Web.Script.Serialization.ScriptIgnore] public CodeDiffLine[] Rows { get { return BuildRows(Before, After); } }
        /// <summary>Obtient le libellé horodaté du changement, avec le projet et le module.</summary>
        [System.Web.Script.Serialization.ScriptIgnore] public string Label { get { return Time.ToString("HH:mm:ss") + "  " + Project + "." + Module +
            (Restored ? UiText.Get("  (restored)") : ""); } }
        /// <summary>Retourne le libellé d’affichage du changement.</summary>
        /// <returns>Valeur de <see cref="Label"/>.</returns>
        public override string ToString() { return Label; }

        /// <summary>Crée un changement vide pour la désérialisation.</summary>
        public CodeChange() { }
        /// <summary>Crée l’enregistrement d’une modification avec ses textes, empreintes et compte de lignes.</summary>
        /// <param name="project">Nom du projet modifié.</param>
        /// <param name="module">Nom du module modifié.</param>
        /// <param name="before">Code avant modification.</param>
        /// <param name="beforeSha256">Empreinte SHA-256 du code initial.</param>
        /// <param name="after">Code après modification.</param>
        /// <param name="afterSha256">Empreinte SHA-256 du code résultant.</param>
        /// <param name="afterLineCount">Nombre de lignes dans le code résultant.</param>
        public CodeChange(string project, string module, string before, string beforeSha256,
            string after, string afterSha256, int afterLineCount)
        {
            Project = project;
            Module = module;
            Before = before;
            BeforeSha256 = beforeSha256;
            After = after;
            AfterSha256 = afterSha256;
            AfterLineCount = afterLineCount;
            Time = DateTime.Now;
        }

        /// <summary>Formate le diff entre le texte courant et le code qui résulterait d’une requête de remplacement.</summary>
        /// <param name="before">Code courant du module.</param>
        /// <param name="request">Requête décrivant la plage à remplacer et le texte de remplacement.</param>
        /// <returns>Diff textuel, ou message informatif lorsqu’il n’existe aucune différence.</returns>
        /// <exception cref="ArgumentException">La plage de remplacement est invalide ou son texte est nul.</exception>
        public static string Preview(string before, Request request)
        {
            return FormatDiff(before, ProposedCode(before, request));
        }

        /// <summary>Construit les lignes structurées du diff prévisionnel d’une requête.</summary>
        /// <param name="before">Code courant du module.</param>
        /// <param name="request">Requête décrivant le remplacement à simuler.</param>
        /// <returns>Lignes du diff, y compris une ligne informative si le code ne change pas.</returns>
        /// <exception cref="ArgumentException">La plage de remplacement est invalide ou son texte est nul.</exception>
        public static CodeDiffLine[] PreviewRows(string before, Request request)
        {
            return BuildRows(before, ProposedCode(before, request));
        }

        /// <summary>Applique en mémoire une requête de remplacement après validation de sa plage.</summary>
        /// <param name="before">Code initial.</param>
        /// <param name="request">Description du remplacement.</param>
        /// <returns>Code résultant avec des séparateurs LF.</returns>
        /// <exception cref="ArgumentException">La plage dépasse le code ou le texte de remplacement est nul.</exception>
        private static string ProposedCode(string before, Request request)
        {
            string[] lines = Lines(before);
            if (request.StartLine < 1 || request.Count < 0 || request.StartLine > lines.Length + 1 ||
                request.Count > lines.Length - request.StartLine + 1 || request.Text == null)
                throw new ArgumentException(UiText.Get("Invalid replacement range."));
            string[] inserted = Lines(request.Text);
            string[] result = lines.Take(request.StartLine - 1).Concat(inserted)
                .Concat(lines.Skip(request.StartLine - 1 + request.Count)).ToArray();
            return string.Join("\n", result);
        }

        /// <summary>Formate les lignes structurées du diff sous forme de texte préfixé par +, − ou espace.</summary>
        /// <param name="before">Version initiale.</param>
        /// <param name="after">Version modifiée.</param>
        /// <returns>Diff avec fins de ligne CRLF, ou message informatif si les versions sont identiques.</returns>
        public static string FormatDiff(string before, string after)
        {
            var rows = BuildRows(before, after);
            if (rows.Length == 1 && rows[0].Kind == CodeDiffKind.Notice) return rows[0].Text;
            var diff = new StringBuilder();
            foreach (var row in rows)
                diff.Append(row.Kind == CodeDiffKind.Added ? '+' :
                    row.Kind == CodeDiffKind.Removed ? '-' : ' ')
                    .Append(row.Text).Append("\r\n");
            return diff.ToString();
        }

        /// <summary>Construit les lignes de comparaison et jusqu’à trois lignes de contexte autour de chaque bloc modifié.</summary>
        /// <param name="before">Version initiale.</param>
        /// <param name="after">Version modifiée.</param>
        /// <returns>Lignes typées du diff, ou une ligne informative lorsque les versions sont identiques.</returns>
        public static CodeDiffLine[] BuildRows(string before, string after)
        {
            var oldLines = CodeRollback.Lines(before); var newLines = CodeRollback.Lines(after);
            var hunks = CodeRollback.Hunks(before, after);
            var rows = new System.Collections.Generic.List<CodeDiffLine>();
            foreach (var hunk in hunks)
            {
                rows.Add(new CodeDiffLine { Kind = CodeDiffKind.Notice, Text = "Bloc " + (hunk.Index + 1) });
                for (int i = Math.Max(0, hunk.BeforeStart - 3); i < hunk.BeforeStart; i++)
                    rows.Add(new CodeDiffLine { Kind = CodeDiffKind.Context, OldLine = i + 1, NewLine = hunk.AfterStart - hunk.BeforeStart + i + 1, Text = oldLines[i] });
                for (int i = 0; i < hunk.Before.Length; i++)
                    rows.Add(new CodeDiffLine { Kind = CodeDiffKind.Removed, OldLine = hunk.BeforeStart + i + 1, Text = hunk.Before[i] });
                for (int i = 0; i < hunk.After.Length; i++)
                    rows.Add(new CodeDiffLine { Kind = CodeDiffKind.Added, NewLine = hunk.AfterStart + i + 1, Text = hunk.After[i] });
                for (int i = hunk.AfterStart + hunk.After.Length; i < Math.Min(newLines.Length, hunk.AfterStart + hunk.After.Length + 3); i++)
                    rows.Add(new CodeDiffLine { Kind = CodeDiffKind.Context, OldLine = hunk.BeforeStart + hunk.Before.Length + i - hunk.AfterStart - hunk.After.Length + 1, NewLine = i + 1, Text = newLines[i] });
            }
            return rows.Count == 0 ? new[] { new CodeDiffLine { Kind = CodeDiffKind.Notice, Text = UiText.Get("No code difference.") } } : rows.ToArray();
        }

        /// <summary>Découpe le code en lignes, normalise les séparateurs et retire la ligne vide finale due au terminateur.</summary>
        /// <param name="code">Code à découper.</param>
        /// <returns>Lignes sans leurs séparateurs ; tableau vide si le code est nul ou vide.</returns>
        private static string[] Lines(string code)
        {
            if (string.IsNullOrEmpty(code)) return new string[0];
            string[] lines = code.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            if (lines.Length > 0 && lines[lines.Length - 1].Length == 0)
                Array.Resize(ref lines, lines.Length - 1);
            return lines;
        }
    }
}
