namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Linq;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using CodexVBE;

    public sealed partial class HistoryAndCodeTests
    {
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
