using System;
using System.Collections.Generic;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbeCollectionEventsTests
    {
        [TestMethod]
        public void ComponentCallbacksIncludeRenameAndDetachOnScopeSwitch()
        {
            var callbacks = new Dictionary<int, Delegate>(); int changed = 0, added = 0, removed = 0;
            var events = new VbeCollectionEvents(() => changed++, true,
                (s, g, id, h) => { Assert.AreEqual(new Guid("0002E116-0000-0000-C000-000000000046"), g); callbacks.Add(id, h); added++; },
                (s, g, id, h) => { Assert.AreSame(callbacks[id], h); callbacks.Remove(id); removed++; });
            var first = new object(); events.Observe(first); events.Observe(first);
            Assert.AreEqual(4, added);
            callbacks[1].DynamicInvoke(new object()); callbacks[2].DynamicInvoke(new object());
            callbacks[3].DynamicInvoke(new object(), "OldName"); callbacks[6].DynamicInvoke(new object());
            Assert.AreEqual(4, changed);
            events.Observe(new object()); Assert.AreEqual(4, removed);
            var late = callbacks[1]; events.Dispose(); events.Dispose();
            late.DynamicInvoke(new object()); Assert.AreEqual(4, changed); Assert.AreEqual(8, removed);
        }
        [TestMethod]
        public void ProjectEventFailureRemovesAllSuccessfulSubscriptionsAndAllowsRetry()
        {
            var active = new HashSet<int>(); bool fail = true;
            using (var events = new VbeCollectionEvents(() => {}, false,
                (s, g, id, h) => { Assert.AreEqual(new Guid("0002E103-0000-0000-C000-000000000046"), g); if (id == 4 && fail) throw new InvalidOperationException(); active.Add(id); },
                (s, g, id, h) => active.Remove(id)))
            {
                var source = new object();
                Assert.ThrowsException<InvalidOperationException>(() => events.Observe(source));
                Assert.AreEqual(0, active.Count); fail = false; events.Observe(source);
                CollectionAssert.AreEquivalent(new[] { 1, 2, 3, 4 }, new List<int>(active));
                events.Observe(null); Assert.AreEqual(0, active.Count);
            }
        }
    }
}
