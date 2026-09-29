using System;
using System.Collections.Generic;
using VBAi;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbeReferenceEventsTests
    {
        /// <summary>Vérifie le refus après destruction et le désabonnement natif d’une source devenue indisponible.</summary>
        [TestMethod]
        public void DefaultRemovalAndDisposedSubscriptionGuardRemainSafe()
        {
            using (var native = new VbeReferenceEvents(() => { }))
                Assert.ThrowsException<ArgumentException>(() => native.Observe(new object()));
            var events = new VbeReferenceEvents(() => { }, (s, g, id, callback) => { });
            events.Observe(new object()); events.Dispose();
            Assert.ThrowsException<ObjectDisposedException>(() => events.Observe(new object()));
        }
        [TestMethod]
        public void ScopeSwitchAndDisposeRemoveExactHandlers()
        {
            var active = new Dictionary<int, Delegate>(); int count = 0, attaches = 0, detaches = 0;
            var first = new object(); var second = new object(); object expected = first;
            var events = new VbeReferenceEvents(() => count++,
                (s, g, id, h) => { Assert.AreSame(expected, s); active.Add(id, h); attaches++; },
                (s, g, id, h) => { Assert.AreSame(active[id], h); active.Remove(id); detaches++; });
            events.Observe(first); events.Observe(first);
            Assert.AreEqual(2, attaches);
            active[1].DynamicInvoke(new object()); active[2].DynamicInvoke(new object());
            Assert.AreEqual(2, count);
            expected = second; events.Observe(second);
            Assert.AreEqual(2, detaches);
            var callback = active[1]; events.Dispose(); events.Dispose();
            Assert.AreEqual(4, detaches); Assert.AreEqual(0, active.Count);
            callback.DynamicInvoke(new object()); Assert.AreEqual(2, count);
        }
        [TestMethod]
        public void PartialSubscriptionFailureDetachesFirstEventAndCanRetry()
        {
            int detached = 0; bool fail = true;
            using (var events = new VbeReferenceEvents(() => {},
                (s, g, id, h) => { if (id == 2 && fail) throw new InvalidOperationException(); },
                (s, g, id, h) => detached++))
            {
                var source = new object();
                Assert.ThrowsException<InvalidOperationException>(() => events.Observe(source));
                Assert.AreEqual(1, detached); fail = false;
                events.Observe(source); events.Observe(null);
                Assert.AreEqual(3, detached);
            }
        }
    }
}
