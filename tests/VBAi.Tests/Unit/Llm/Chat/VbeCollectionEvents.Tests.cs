using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class VbeCollectionEventsTests
    {
        /// <summary>Double du conteneur COM pour vérifier l’identité de la connexion avant abonnement.</summary>
        [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
        public sealed class ConnectionContainer : IConnectionPointContainer
        {
            /// <summary>Point retourné à la recherche de l’interface demandée.</summary>
            internal IConnectionPoint Point;
            /// <summary>Retourne le point configurable.</summary>
            public void FindConnectionPoint(ref Guid riid, out IConnectionPoint point) { point = Point; }
            /// <summary>Énumération non utilisée par le contrat testé.</summary>
            public void EnumConnectionPoints(out IEnumConnectionPoints points) { points = null; }
        }
        /// <summary>Point de connexion dont l’identité native est contrôlée par le scénario.</summary>
        [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
        public sealed class ConnectionPoint : IConnectionPoint
        {
            /// <summary>Identité de l’interface annoncée.</summary>
            internal Guid Id;
            /// <summary>Compte les abonnements COM réellement marshalisés.</summary>
            internal int Advised, Unadvised;
            /// <summary>Expose l’identité annoncée.</summary>
            public void GetConnectionInterface(out Guid id) { id = Id; }
            /// <summary>Retourne un conteneur absent, non utilisé par ces gardes.</summary>
            public void GetConnectionPointContainer(out IConnectionPointContainer container) { container = null; }
            /// <summary>Abonnement non disponible dans ce double managé.</summary>
            public void Advise(object sink, out int cookie) { cookie = ++Advised; }
            /// <summary>Désabonnement sans effet dans le double.</summary>
            public void Unadvise(int cookie) { Unadvised++; }
            /// <summary>Énumération non utilisée par le contrat testé.</summary>
            public void EnumConnections(out IEnumConnections connections) { connections = null; }
        }

        /// <summary>Refuse les conteneurs absents, interfaces absentes ou incompatibles avant la combinaison COM.</summary>
        [TestMethod]
        public void DefaultNativeConnectionChecksIdentityAndManagedObjectsCannotSubscribe()
        {
            using (var events = new VbeCollectionEvents(() => { }, true))
            {
                Assert.ThrowsException<InvalidOperationException>(() => events.Observe(new object()));
                var source = new ConnectionContainer();
                Assert.ThrowsException<InvalidOperationException>(() => events.Observe(source));
                source.Point = new ConnectionPoint { Id = Guid.Empty };
                Assert.ThrowsException<InvalidOperationException>(() => events.Observe(source));
                source.Point = new ConnectionPoint { Id = new Guid("0002E116-0000-0000-C000-000000000046") };
                Exception failure = null; try { events.Observe(source); } catch (Exception error) { failure = error; }
                Assert.IsNotNull(failure, "A managed connection-point double must not be accepted as a native COM event source.");
                events.Observe(null);
            }
            using (var events = new VbeCollectionEvents(() => { }, false, (s, g, id, h) => { }))
                events.Observe(new object()); // The native default removal safely handles an already-unavailable source.
        }

        /// <summary>Continue les désabonnements après erreur et refuse les nouveaux abonnements après Dispose.</summary>
        [TestMethod]
        public void DetachFailuresAndLateRenameCallbacksDoNotInvalidateDisposedState()
        {
            int removed = 0, changed = 0; Delegate rename = null;
            var events = new VbeCollectionEvents(() => changed++, true,
                (s, g, id, callback) => { if (id == 3) rename = callback; },
                (s, g, id, callback) => { removed++; throw new InvalidOperationException("closed project"); });
            events.Observe(new object()); rename.DynamicInvoke(new object(), "old"); Assert.AreEqual(1, changed);
            events.Dispose(); Assert.AreEqual(4, removed);
            rename.DynamicInvoke(new object(), "old"); Assert.AreEqual(1, changed);
            Assert.ThrowsException<ObjectDisposedException>(() => events.Observe(new object()));
        }

        /// <summary>Vérifie les identités et les quatre abonnements en remplaçant seulement la combinaison native.</summary>
        [TestMethod]
        public void NativeConnectionValidatesEveryInterfaceBeforeCombiningDelegates()
        {
            var point = new ConnectionPoint { Id = new Guid("0002E116-0000-0000-C000-000000000046") };
            var source = new ConnectionContainer { Point = point };
            var original = VbeCollectionEvents.CombineNative; int combined = 0, removed = 0;
            try
            {
                VbeCollectionEvents.CombineNative = (target, id, member, callback) =>
                {
                    Assert.AreSame(source, target); Assert.AreEqual(point.Id, id); Assert.IsNotNull(callback); combined++;
                };
                using (var events = new VbeCollectionEvents(() => { }, true, remove: (s, id, member, callback) => removed++)) events.Observe(source);
                Assert.AreEqual(4, combined); Assert.AreEqual(4, removed);
            }
            finally { VbeCollectionEvents.CombineNative = original; }
        }
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
            using (var events = new VbeCollectionEvents(() => { }, false,
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
