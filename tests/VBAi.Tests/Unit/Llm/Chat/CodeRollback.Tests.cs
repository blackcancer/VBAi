namespace VBAi.Tests.Unit
{
    using System;
    using System.Linq;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using VBAi;

    /// <summary>Vérifie la restauration des changements de code et le refus des contextes ambigus.</summary>
    public sealed partial class HistoryAndCodeTests
    {
        /// <summary>Restaure exactement le texte précédent et refuse si son contexte n’est plus unique.</summary>
        [TestMethod]
        public void RollbackRestoresExactChangeAndRefusesAmbiguousContext()
        {
            var before = "start\nold\nend";
            var after = "start\nnew\nend";
            var change = new CodeChange
            {
                Module = "M1",
                Before = before,
                After = after
            };
            Assert.AreEqual(before.Replace("\n", "\r\n"), CodeRollback.Apply(change, after));
            var conflict = Assert.ThrowsException<InvalidOperationException>(() => CodeRollback.Apply(change, "other\nnew\nother"));
            StringAssert.Contains(conflict.Message, "Conflit");
        }
    }
}

namespace VBAi.Tests.Unit
{
    using System;
    using System.Linq;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed class CodeRollbackBoundaryTests
    {
        [TestMethod]
        public void SmallDiffMatrixRestoresInsertionsDeletionsReordersAndNormalizesNewlines()
        {
            CollectionAssert.AreEqual(new string[0],CodeRollback.Lines(null));
            CollectionAssert.AreEqual(new[] {"a","b","c","d"},CodeRollback.Lines("a\r\nb\rc\nd"));
            string[,] cases = { {"","x"},{"x",""},{"a","b"},{"a\nb","a\nb\nc"},{"a\nb\nc","a\nb"},
                {"a\nb","x\na\nb"},{"x\na\nb","a\nb"},{"a\nb\nc","c\nb\na"},{"a\nb\nc","a\nx\nc"} };
            for(int i=0;i<cases.GetLength(0);i++)
            {
                string before=cases[i,0],after=cases[i,1];
                var hunks=CodeRollback.Hunks(before,after);Assert.IsTrue(hunks.Length>0);
                CollectionAssert.AreEqual(Enumerable.Range(0,hunks.Length).ToArray(),hunks.Select(h=>h.Index).ToArray());
                var change=new CodeChange {Module="M",Before=before,After=after};
                Assert.AreEqual(before.Replace("\n","\r\n"),CodeRollback.Apply(change,after),"case "+i);
            }
            Assert.AreEqual(0,CodeRollback.Hunks("same","same").Length);
            Assert.AreEqual(0,CodeRollback.Hunks("",null).Length);
        }

        [TestMethod]
        public void LargeDiffMatrixUsesBoundedHunkAndKeepsSharedPrefixAndSuffix()
        {
            string common=string.Join("\n",Enumerable.Range(0,1600).Select(i=>"line"+i));
            string[,] cases = { {common,common},{common,common+"\nextra"},{common+"\nextra",common},
                {common,"extra\n"+common},{"extra\n"+common,common},{"old\n"+common,"new\n"+common},
                {common+"\nold",common+"\nnew"} };
            for(int i=0;i<cases.GetLength(0);i++)
            {
                string before=cases[i,0],after=cases[i,1];var hunks=CodeRollback.Hunks(before,after);
                Assert.AreEqual(i==0?0:1,hunks.Length,"case "+i);
                if(i==0)continue;
                Assert.AreEqual(before.Replace("\n","\r\n"),CodeRollback.Apply(new CodeChange {Module="M",Before=before,After=after},after));
            }
        }

        [TestMethod]
        public void PartialUndoAdjustsOffsetsAndMatchesOnlyUniqueUnchangedContext()
        {
            var change=new CodeChange {Module="M",Before="A\nKeep\nB",After="A2\nExtra\nKeep\nB2"};
            string partial=CodeRollback.Apply(change,change.After,0);change.RestoredHunks.Add(0);
            Assert.AreEqual("A\r\nKeep\r\nB2",partial);
            Assert.AreEqual("A\r\nKeep\r\nB",CodeRollback.Apply(change,partial,1));
            Assert.ThrowsException<InvalidOperationException>(()=>CodeRollback.Apply(change,partial,0));
            var lastFirst=new CodeChange {Module="M",Before="A\nKeep\nB",After="A2\nKeep\nB2"};
            partial=CodeRollback.Apply(lastFirst,lastFirst.After,1);lastFirst.RestoredHunks.Add(1);
            Assert.AreEqual("A\r\nKeep\r\nB",CodeRollback.Apply(lastFirst,partial,0));
            var shifted=new CodeChange {Module="M",Before="start\nold\nend",After="start\nnew\nend"};
            Assert.AreEqual("prefix\r\nstart\r\nold\r\nend",CodeRollback.Apply(shifted,"prefix\nstart\nnew\nend"));
            Assert.ThrowsException<InvalidOperationException>(()=>CodeRollback.Apply(shifted,"start\nnew\nend\nstart\nnew\nend"));
            Assert.ThrowsException<InvalidOperationException>(()=>CodeRollback.Apply(new CodeChange {Module="M",Before="old",After=""},"unrelated"));
        }
    }
}
