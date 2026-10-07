namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System.Linq;
    using VBAi;

    /// <summary>Vérifie les lignes de diff et le repli du contexte inchangé.</summary>
    public sealed partial class GitReviewTests
    {
        /// <summary>Inclut les grands modules et replie uniquement le contexte inchangé.</summary>
        [TestMethod]
        public void DiffIncludesLargeModulesAndFoldsOnlyUnchangedContext()
        {
            string before = string.Join("\n", Enumerable.Range(0, 5000).Select(x => "line " + x));
            string after = before.Replace("line 4200\n", "changed\n");
            var rows = DiffModel.Build(before, after, false, false);
            Assert.AreEqual(5000, rows.Count);
            Assert.AreEqual("changed", rows[4200].Right);
            var folded = DiffModel.Build(before, after, true, true);
            Assert.IsTrue(folded.Count < 20);
            Assert.IsTrue(folded.Any(x => x.Left == "line 4200"));
            Assert.IsTrue(folded.Any(x => x.Right == "changed"));
            Assert.AreEqual(1, folded.Count(x => x.Right == "changed"));
        }
    }
}

namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System.Linq;
    using VBAi;
    /// <summary>Vérifie l’alignement des insertions, suppressions et groupes de contexte masqué.</summary>
    [TestClass]
    public sealed class DiffModelCoverageTests
    {
        /// <summary>Conserve les numéros de lignes lors d’insertions et suppressions dans les deux layouts.</summary>
        [TestMethod]
        public void InsertionsAndDeletionsRetainLineNumbersInBothLayouts()
        {
            foreach (var unified in new[] { false, true })
            {
                var insert = DiffModel.Build("a\nz", "a\nb\nc\nz", unified, false);
                Assert.AreEqual(4, insert.Count); Assert.IsNull(insert[1].Old); Assert.AreEqual(2, insert[1].New); Assert.AreEqual("b", insert[1].Unified);
                var delete = DiffModel.Build("a\nb\nc\nz", "a\nz", unified, false);
                Assert.AreEqual(4, delete.Count); Assert.AreEqual(2, delete[1].Old); Assert.IsNull(delete[1].New); Assert.AreEqual("b", delete[1].Unified);
                var replace = DiffModel.Build("a\nb\nc", "x", unified, false);
                Assert.IsTrue(replace.All(r => r.Hunk >= 0)); Assert.AreEqual(unified ? 4 : 3, replace.Count);
                Assert.AreEqual("a", new DiffRow { Left = "a" }.Unified);
            }
        }
        /// <summary>Maintient le contexte modifié et regroupe les lignes cachées consécutives.</summary>
        [TestMethod]
        public void CollapseKeepsChangedContextAndGroupsConsecutiveHiddenRows()
        {
            Assert.AreEqual(0, DiffModel.Build("", "", true, true).Count);
            var unchanged = DiffModel.Build("a\nb\nc", "a\nb\nc", false, true);
            Assert.AreEqual(1, unchanged.Count); Assert.IsTrue(unchanged[0].Fold);
            string before = string.Join("\n", Enumerable.Range(0, 30).Select(i => "line" + i));
            string after = before.Replace("line0\n", "first\n").Replace("line15\n", "middle\n").Replace("line29", "last");
            var rows = DiffModel.Build(before, after, false, true);
            Assert.AreEqual(2, rows.Count(r => r.Fold)); Assert.AreEqual(3, rows.Count(r => r.Hunk >= 0));
            Assert.IsFalse(rows.First().Fold); Assert.IsFalse(rows.Last().Fold);
        }
    }
}
