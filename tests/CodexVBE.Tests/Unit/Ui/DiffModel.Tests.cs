namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net;
    using System.Net.Http;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Web.Script.Serialization;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class GitReviewTests
    {
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
