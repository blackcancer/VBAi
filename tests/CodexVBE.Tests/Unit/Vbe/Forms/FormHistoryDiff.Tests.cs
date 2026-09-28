using System.Linq;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class FormHistoryDiffTests
    {
        private static object Property(string name, object value, string error = null)
            => new { Name = name, Value = value, Error = error };
        private static object Tree(params object[] properties) => new { Properties = properties, Controls = new object[0] };
        [TestMethod]
        public void AvailabilityFlagsAndPropertyOrderDoNotProveAnEdit()
        {
            Assert.AreEqual(0, FormHistoryDiff.Compare(Tree(Property("CanRedo", false), Property("Caption", "A")),
                Tree(Property("Caption", "A"), Property("CanRedo", true))).Length);
        }
        [TestMethod]
        public void ReadErrorsAreReportedAndNotPresentedAsEdits()
        {
            var unreadable = Tree(Property("Caption", null, "COM failure"));
            Assert.AreEqual(1, FormHistoryDiff.ReadErrorCount(unreadable));
            Assert.AreEqual(0, FormHistoryDiff.Compare(unreadable, Tree(Property("Caption", "A"))).Length);
            Assert.AreEqual(1, FormHistoryDiff.Compare(Tree(Property("Caption", "A")), Tree(Property("Caption", "B"))).Length);
        }
        [TestMethod]
        public void NestedControlRemovalRetainsCanonicalIdentity()
        {
            var nested = new { Path = "Controls/Frame/Controls/Button", Properties = new[] { Property("Left", 20) }, Children = new object[0] };
            var parent = new { Path = "Controls/Frame", Properties = new object[0], Children = new object[] { nested } };
            var emptyParent = new { Path = "Controls/Frame", Properties = new object[0], Children = new object[0] };
            var changes = FormHistoryDiff.Compare(new { Controls = new[] { parent } }, new { Controls = new[] { emptyParent } });
            Assert.AreEqual(2, changes.Length);
            Assert.IsTrue(changes.All(x => x.Path == "Controls/Frame/Controls/Button"));
            Assert.AreEqual(true, changes.Single(x => x.Property == "$exists").Before);
            Assert.IsNull(changes.Single(x => x.Property == "$exists").After);
        }
    }
}
