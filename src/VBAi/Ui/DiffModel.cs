using System;
using System.Collections.Generic;
using System.Linq;
namespace VBAi
{

    /// <summary>Représente une ligne du diff côte à côte ou unifiée.</summary>
    internal sealed class DiffRow
    {

        /// <summary>Texte de la ligne provenant de la version antérieure.</summary>
        /// <value>Texte à gauche, ou null si cette ligne ne figure pas dans l’ancienne version.</value>
        public string Left { get; set; }

        /// <summary>Texte de la ligne provenant de la nouvelle version.</summary>
        /// <value>Texte à droite, ou null si cette ligne ne figure pas dans la nouvelle version.</value>
        public string Right { get; set; }

        /// <summary>Numéro de ligne dans l’ancienne version.</summary>
        /// <value>Numéro de ligne, ou null si la ligne est une insertion.</value>
        public int? Old { get; set; }

        /// <summary>Numéro de ligne dans la nouvelle version.</summary>
        /// <value>Numéro de ligne, ou null si la ligne est une suppression.</value>
        public int? New { get; set; }

        /// <summary>Index du groupe de changements auquel la ligne appartient.</summary>
        /// <value>-1 pour le contexte inchangé ; sinon index du groupe.</value>
        public int Hunk { get; set; } = -1;

        /// <summary>Indique si la ligne représente une zone de contexte repliée.</summary>
        /// <value>Vrai pour une ligne de pliage.</value>
        public bool Fold { get; set; }

        /// <summary>Texte choisi pour la vue unifiée.</summary>
        /// <value>Texte droit s’il existe, sinon texte gauche.</value>
        public string Unified { get { return Right ?? Left; } }
    }

    /// <summary>Construit les lignes visibles à partir des groupes de changements calculés par CodeRollback.</summary>
    internal static class DiffModel
    {

        /// <summary>Associe le contexte inchangé et les changements dans une vue côte à côte ou unifiée, puis replie le contexte éloigné.</summary>
        /// <param name="before">Texte de la version antérieure.</param>
        /// <param name="after">Texte de la nouvelle version.</param>
        /// <param name="unified">Vrai pour placer les suppressions et insertions dans une colonne unique.</param>
        /// <param name="collapse">Vrai pour remplacer le contexte éloigné par des lignes de pliage.</param>
        /// <returns>Lignes ordonnées du diff à afficher.</returns>
        internal static List<DiffRow> Build(string before, string after, bool unified, bool collapse)
        {
            var all = new List<DiffRow>();
            var left = CodeRollback.Lines(before); var right = CodeRollback.Lines(after); int x = 0, y = 0;
            foreach (var hunk in CodeRollback.Hunks(before, after))
            {
                while (x < hunk.BeforeStart && y < hunk.AfterStart) all.Add(new DiffRow { Left = left[x], Right = right[y], Old = ++x, New = ++y });
                if (unified)
                {
                    foreach (string line in hunk.Before) all.Add(new DiffRow { Left = line, Old = ++x, Hunk = hunk.Index });
                    foreach (string line in hunk.After) all.Add(new DiffRow { Right = line, New = ++y, Hunk = hunk.Index });
                }
                else for (int i = 0; i < Math.Max(hunk.Before.Length, hunk.After.Length); i++)
                {
                    var row = new DiffRow { Hunk = hunk.Index };
                    if (i < hunk.Before.Length) { row.Left = left[x]; row.Old = ++x; }
                    if (i < hunk.After.Length) { row.Right = right[y]; row.New = ++y; }
                    all.Add(row);
                }
            }
            while (x < left.Length && y < right.Length) all.Add(new DiffRow { Left = left[x], Right = right[y], Old = ++x, New = ++y });
            if (!collapse) return all;
            var keep = new bool[all.Count];
            for (int i = 0; i < all.Count; i++) if (all[i].Hunk >= 0) for (int j = Math.Max(0, i - 3); j < Math.Min(all.Count, i + 4); j++) keep[j] = true;
            var visible = new List<DiffRow>();
            for (int i = 0; i < all.Count; i++)
                if (keep[i]) visible.Add(all[i]);
                else if (visible.Count == 0 || !visible.Last().Fold) visible.Add(new DiffRow { Fold = true, Left = "…", Right = "…" });
            return visible;
        }
    }
}
