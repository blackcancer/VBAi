using System;
using System.Collections.Generic;
using System.Linq;
namespace CodexVBE
{
    internal sealed class DiffRow
    {
        public string Left { get; set; }
        public string Right { get; set; }
        public int? Old { get; set; }
        public int? New { get; set; }
        public int Hunk { get; set; } = -1;
        public bool Fold { get; set; }
        public string Unified { get { return Right ?? Left; } }
    }
    internal static class DiffModel
    {
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
