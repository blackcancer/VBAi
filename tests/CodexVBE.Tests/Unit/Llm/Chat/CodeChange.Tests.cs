namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Linq;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using CodexVBE;

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
    }
}
