using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace CodexVBE
{
    // All operations and callbacks run on the owning VBE STA. Never inspect a COM
    // object inside an event callback: removed objects may already be invalid.
    internal sealed class VbeCollectionEvents : IDisposable
    {
        /// <summary>Frontière native de combinaison des événements COM, conservée séparément des gardes de connexion.</summary>
        internal static Action<object, Guid, int, Delegate> CombineNative = ComEventsHelper.Combine;
        private readonly Guid iid;
        private readonly Action<object, Guid, int, Delegate> add, remove;
        private readonly List<Tuple<int, Delegate>> handlers = new List<Tuple<int, Delegate>>();
        private readonly List<Tuple<int, Delegate>> attached = new List<Tuple<int, Delegate>>();
        private object source;
        private bool disposed;
        internal VbeCollectionEvents(Action changed, bool components,
            Action<object, Guid, int, Delegate> add = null,
            Action<object, Guid, int, Delegate> remove = null)
        {
            iid = new Guid(components ? "0002E116-0000-0000-C000-000000000046" : "0002E103-0000-0000-C000-000000000046");
            this.add = add ?? Connect;
            this.remove = remove ?? ((target, id, member, callback) => ComEventsHelper.Remove(target, id, member, callback));
            Action<object> item = value => { if (!disposed) changed(); };
            Action<object, string> renamed = (value, oldName) => { if (!disposed) changed(); };
            handlers.Add(Tuple.Create(1, (Delegate)item));
            handlers.Add(Tuple.Create(2, (Delegate)item));
            handlers.Add(Tuple.Create(3, (Delegate)renamed));
            // Reloaded for components, activated for projects. Selection/activation
            // of individual components does not invalidate the symbol inventory.
            handlers.Add(Tuple.Create(components ? 6 : 4, (Delegate)item));
        }
        private static void Connect(object target, Guid iid, int member, Delegate handler)
        {
            var container = target as IConnectionPointContainer;
            if (container == null) throw new InvalidOperationException("No native collection event container.");
            IConnectionPoint point;
            container.FindConnectionPoint(ref iid, out point);
            if (point == null) throw new InvalidOperationException("The native event connection point is absent.");
            Guid actual; point.GetConnectionInterface(out actual);
            if (actual != iid) throw new InvalidOperationException("The native event connection interface does not match.");
            CombineNative(target, iid, member, handler);
        }
        internal void Observe(object next)
        {
            if (disposed) throw new ObjectDisposedException(nameof(VbeCollectionEvents));
            if (ReferenceEquals(source, next)) return;
            Detach();
            source = next;
            if (next == null) return;
            try
            {
                foreach (var entry in handlers)
                {
                    add(source, iid, entry.Item1, entry.Item2);
                    attached.Add(entry);
                }
            }
            catch { Detach(); throw; }
        }
        private void Detach()
        {
            foreach (var entry in attached)
            {
                try { remove(source, iid, entry.Item1, entry.Item2); }
                catch (Exception error) { LoadLog.Write("Collection event detach: " + error.Message); }
            }
            attached.Clear(); source = null;
        }
        public void Dispose() { if (disposed) return; disposed = true; Detach(); }
    }
}
