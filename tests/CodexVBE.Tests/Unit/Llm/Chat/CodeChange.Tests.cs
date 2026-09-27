namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Linq;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using CodexVBE;

    public sealed partial class HistoryAndCodeTests
    {
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
    }
}
