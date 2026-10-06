using System;
using System.Collections.Generic;
using System.Linq;

namespace VBAi
{

    /// <summary>Représente une zone contiguë différente entre deux versions d’un module.</summary>
    internal sealed class CodeHunk
    {

        /// <summary>Obtient ou définit l’index de la zone dans la séquence de différences.</summary>
        /// <value>Index de la zone dans la séquence de différences.</value>
        public int Index { get; set; }

        /// <summary>Obtient ou définit la première ligne concernée dans la version initiale, indexée à partir de zéro.</summary>
        /// <value>Première ligne de l’ancienne version, indexée à partir de zéro.</value>
        public int BeforeStart { get; set; }

        /// <summary>Obtient ou définit la première ligne concernée dans la version modifiée, indexée à partir de zéro.</summary>
        /// <value>Première ligne de la nouvelle version, indexée à partir de zéro.</value>
        public int AfterStart { get; set; }

        /// <summary>Obtient ou définit les lignes de la version initiale remplacées par cette zone.</summary>
        /// <value>Lignes remplacées dans l’ancienne version.</value>
        public string[] Before { get; set; }

        /// <summary>Obtient ou définit les lignes de la version modifiée correspondant à cette zone.</summary>
        /// <value>Lignes insérées dans la nouvelle version.</value>
        public string[] After { get; set; }
    }

    /// <summary>Calcule les différences textuelles et restaure des blocs de code en vérifiant leur contexte.</summary>
    internal static class CodeRollback
    {

        /// <summary>Découpe un texte en lignes, normalise les fins de ligne et retourne un tableau vide pour une entrée vide.</summary>
        /// <param name="text">Texte à découper.</param>
        /// <returns>Les lignes du texte sans leurs séparateurs.</returns>
        public static string[] Lines(string text) { return string.IsNullOrEmpty(text) ? new string[0] : text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n'); }

        /// <summary>Construit les zones modifiées entre une version initiale et une version modifiée.</summary>
        /// <param name="before">Texte initial.</param>
        /// <param name="after">Texte modifié.</param>
        /// <returns>Les zones ordonnées par position ; les textes identiques ne produisent aucune zone.</returns>
        public static CodeHunk[] Hunks(string before, string after)
        {
            var a = Lines(before); var b = Lines(after);
            // Bound memory for very large modules; one conservative hunk remains reversible.
            if ((long)(a.Length + 1) * (b.Length + 1) > 2000000)
            {
                int prefix = 0, suffix = 0;
                while (prefix < a.Length && prefix < b.Length && a[prefix] == b[prefix]) prefix++;
                while (suffix < a.Length - prefix && suffix < b.Length - prefix && a[a.Length - suffix - 1] == b[b.Length - suffix - 1]) suffix++;
                return prefix == a.Length && prefix == b.Length ? new CodeHunk[0] : new[] { new CodeHunk {
                    BeforeStart = prefix, AfterStart = prefix, Before = a.Skip(prefix).Take(a.Length - prefix - suffix).ToArray(),
                    After = b.Skip(prefix).Take(b.Length - prefix - suffix).ToArray() } };
            }
            var lengths = new int[a.Length + 1, b.Length + 1];
            for (int i = a.Length - 1; i >= 0; i--)
                for (int j = b.Length - 1; j >= 0; j--)
                    lengths[i, j] = a[i] == b[j] ? 1 + lengths[i + 1, j + 1] : Math.Max(lengths[i + 1, j], lengths[i, j + 1]);
            var result = new List<CodeHunk>(); int x = 0, y = 0;
            while (x < a.Length || y < b.Length)
            {
                if (x < a.Length && y < b.Length && a[x] == b[y]) { x++; y++; continue; }
                int startX = x, startY = y;
                while (x < a.Length || y < b.Length)
                {
                    if (x < a.Length && y < b.Length && a[x] == b[y]) break;
                    if (y == b.Length || (x < a.Length && lengths[x + 1, y] >= lengths[x, y + 1])) x++; else y++;
                }
                result.Add(new CodeHunk { Index = result.Count, BeforeStart = startX, AfterStart = startY,
                    Before = a.Skip(startX).Take(x - startX).ToArray(), After = b.Skip(startY).Take(y - startY).ToArray() });
            }
            return result.ToArray();
        }

        /// <summary>Restaure une ou toutes les zones encore modifiées d’un changement, en refusant les contextes ambigus ou modifiés.</summary>
        /// <param name="change">Changement à restaurer, avec ses versions et l’état des zones déjà restaurées.</param>
        /// <param name="current">Contenu actuel du module.</param>
        /// <param name="onlyHunk">Index facultatif de la seule zone à restaurer.</param>
        /// <returns>Le contenu résultant, avec des fins de ligne CRLF.</returns>
        /// <exception cref="InvalidOperationException">Aucune zone sélectionnée ne reste à restaurer, ou le bloc ne peut pas être retrouvé sans ambiguïté.</exception>
        public static string Apply(CodeChange change, string current, int? onlyHunk = null)
        {
            var lines = Lines(current).ToList(); var original = Lines(change.After).ToList();
            var all = Hunks(change.Before, change.After);
            foreach (var restored in all.Where(x => change.RestoredHunks.Contains(x.Index)).Reverse())
            { original.RemoveRange(restored.AfterStart, restored.After.Length); original.InsertRange(restored.AfterStart, restored.Before); }
            var hunks = all.Where(x => !change.RestoredHunks.Contains(x.Index) && (!onlyHunk.HasValue || x.Index == onlyHunk.Value)).ToArray();
            if (hunks.Length == 0) throw new InvalidOperationException("Ce bloc est déjà annulé.");
            bool exact = lines.SequenceEqual(original);
            foreach (var hunk in hunks.Reverse())
            {
                int expected = hunk.AfterStart + all.Where(x => change.RestoredHunks.Contains(x.Index) && x.Index < hunk.Index).Sum(x => x.Before.Length - x.After.Length);
                int position = expected;
                if (!exact)
                {
                    // Match the edited block with up to three context lines on each side.
                    var left = original.Skip(Math.Max(0, expected - 3)).Take(Math.Min(3, expected)).ToArray();
                    var right = original.Skip(expected + hunk.After.Length).Take(3).ToArray();
                    var needle = left.Concat(hunk.After).Concat(right).ToArray();
                    var matches = new List<int>();
                    for (int i = 0; needle.Length > 0 && i <= lines.Count - needle.Length; i++)
                        if (lines.Skip(i).Take(needle.Length).SequenceEqual(needle)) matches.Add(i + left.Length);
                    if (matches.Count != 1) throw new InvalidOperationException("Conflit dans " + change.Module + ", bloc " + (hunk.Index + 1) + " : le code ou son contexte a changé. Aucun remplacement forcé.");
                    position = matches[0];
                }
                lines.RemoveRange(position, hunk.After.Length); lines.InsertRange(position, hunk.Before);
                original.RemoveRange(expected, hunk.After.Length); original.InsertRange(expected, hunk.Before);
            }
            return string.Join("\r\n", lines);
        }
    }
}
