namespace VBAi.Tests.Unit
{
    using System;
    using System.Linq;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using VBAi;

    /// <summary>Vérifie les plages de modification et les numéros de lignes des diff.</summary>
    public sealed partial class HistoryAndCodeTests
    {
        /// <summary>Refuse une plage invalide avant qu’une modification de l’hôte puisse être faite.</summary>
        [TestMethod]
        public void PreviewRejectsInvalidRangeBeforeAnyHostEdit()
        {
            var request = new Request
            {
                StartLine = 3,
                Count = 1,
                Text = "C"
            };
            Assert.ThrowsException<ArgumentException>(() => CodeChange.Preview("A\nB", request));
            request.StartLine = 2;
            request.Count = -1;
            Assert.ThrowsException<ArgumentException>(() => CodeChange.Preview("A\nB", request));
        }

        /// <summary>Conserve les numéros de lignes séparés des versions ancienne et nouvelle du diff.</summary>
        [TestMethod]
        public void DiffRowsPreserveOldAndNewLineNumbers()
        {
            var rows = CodeChange.BuildRows("first\nold\nlast", "first\nnew\nlast");
            var removed = rows.Single(x => x.Kind == CodeDiffKind.Removed);
            var added = rows.Single(x => x.Kind == CodeDiffKind.Added);
            Assert.AreEqual(2, removed.OldLine);
            Assert.AreEqual(2, added.NewLine);
            Assert.AreEqual("old", removed.Text);
            Assert.AreEqual("new", added.Text);
        }
        [TestMethod]
        public void PreviewBoundaryMatrixRejectsEveryInvalidRangeAndPreservesTrailingLineSemantics()
        {
            using(var culture=new Infrastructure.LocalizationScope())
            {
                foreach(var request in new[] {new Request {StartLine=0,Count=0,Text="x"},new Request {StartLine=1,Count=-1,Text="x"},
                    new Request {StartLine=4,Count=0,Text="x"},new Request {StartLine=2,Count=2,Text="x"},new Request {StartLine=1,Count=0,Text=null}})
                    Assert.ThrowsException<ArgumentException>(()=>CodeChange.Preview("A\nB",request));
                Assert.AreEqual("No code difference.",CodeChange.Preview("A\r\nB",new Request {StartLine=2,Count=1,Text="B\r\n"}));
                var trimmed=CodeChange.PreviewRows("A\r\nB\r\n",new Request {StartLine=2,Count=1,Text="B\r\n"});
                Assert.AreEqual(1,trimmed.Count(row=>row.Kind==CodeDiffKind.Removed && row.Text==""));
                Assert.IsFalse(trimmed.Any(row=>row.Kind==CodeDiffKind.Added));
                Assert.AreEqual("No code difference.",CodeChange.Preview(null,new Request {StartLine=1,Count=0,Text=""}));
                StringAssert.Contains(CodeChange.Preview("A\rB",new Request {StartLine=3,Count=0,Text="C"}),"+C");
                Assert.IsTrue(CodeChange.PreviewRows("A",new Request {StartLine=1,Count=1,Text="B"}).Length>1);
                Assert.AreEqual("+",new CodeDiffLine {Kind=CodeDiffKind.Added}.Sign);
                Assert.AreEqual("−",new CodeDiffLine {Kind=CodeDiffKind.Removed}.Sign);
                Assert.AreEqual("",new CodeDiffLine {Kind=CodeDiffKind.Context}.Sign);
                Assert.AreEqual("",new CodeDiffLine {Kind=CodeDiffKind.Notice}.Sign);
                var unchanged=CodeChange.BuildRows("same","same");Assert.AreEqual(1,unchanged.Length);Assert.AreEqual(CodeDiffKind.Notice,unchanged[0].Kind);
                var changed=CodeChange.BuildRows("old","new");Assert.IsTrue(changed.Length>1);Assert.IsTrue(changed.Any(row=>row.Kind!=CodeDiffKind.Notice));
                var change=new CodeChange("P","M","before","before-sha","after","after-sha",1) {Time=new DateTime(2026,1,2,12,34,56)};
                Assert.AreEqual("12:34:56  P.M",change.Label);Assert.AreEqual(change.Label,change.ToString());
                change.Restored=true;StringAssert.Contains(change.Label,"(restored)");
                StringAssert.Contains(change.Diff,"-before");Assert.IsTrue(change.Rows.Any(row=>row.Kind==CodeDiffKind.Added));
            }
        }

    }
}
