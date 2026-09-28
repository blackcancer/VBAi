using System;
using System.Runtime.InteropServices;

namespace CodexVBE
{
    // Subscribe and dispose on the VBE STA. Callbacks only invalidate managed state.
    internal sealed class VbeReferenceEvents : IDisposable
    {
        private static readonly Guid EventInterface = new Guid("0002E118-0000-0000-C000-000000000046");
        private readonly Action<object, Guid, int, Delegate> add;
        private readonly Action<object, Guid, int, Delegate> remove;
        private readonly Action<object> handler;
        private object source;
        private bool added, removed, disposed;

        internal VbeReferenceEvents(Action changed,
            Action<object, Guid, int, Delegate> add = null,
            Action<object, Guid, int, Delegate> remove = null)
        {
            this.add = add ?? new Action<object, Guid, int, Delegate>(ComEventsHelper.Combine);
            this.remove = remove ?? ((target, iid, id, callback) => ComEventsHelper.Remove(target, iid, id, callback));
            handler = item => { if (!disposed) changed(); };
        }

        internal void Observe(object next)
        {
            if (disposed) throw new ObjectDisposedException(nameof(VbeReferenceEvents));
            if (ReferenceEquals(source, next)) return;
            Detach();
            if (next == null) return;
            source = next;
            try
            {
                add(source, EventInterface, 1, handler); added = true;
                add(source, EventInterface, 2, handler); removed = true;
            }
            catch { Detach(); throw; }
        }

        private void Detach()
        {
            if (source == null) return;
            // Attempt both removals even when the project has already been closed.
            if (added) TryRemove(1);
            if (removed) TryRemove(2);
            added = removed = false;
            source = null;
        }
        private void TryRemove(int id)
        {
            try { remove(source, EventInterface, id, handler); }
            catch (Exception error) { LoadLog.Write("Reference event detach: " + error.Message); }
        }
        public void Dispose() { if (disposed) return; disposed = true; Detach(); }
    }
}
