using System;
using System.Threading;
using System.Windows.Forms;

namespace VBAi.Tests.Integration
{
    /// <summary>Processes borrowed Office event arguments synchronously on the fixture's owning STA.</summary>
    internal interface IOfficeCoverageEventDispatcher : IDisposable
    {
        void Invoke(Action callback);
    }

    /// <summary>A hidden message target; it never shows a window or takes keyboard focus.</summary>
    internal sealed class OfficeCoverageEventDispatcher : IOfficeCoverageEventDispatcher
    {
        private readonly int owner = Thread.CurrentThread.ManagedThreadId;
        private readonly Control target;
        private volatile bool disposed;

        internal OfficeCoverageEventDispatcher()
        {
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                throw new InvalidOperationException("Office event observation requires the fixture's original STA.");
            target = new Control();
            try
            {
                // Creating the handle on this thread is essential: Invoke must not create it
                // on the RPC callback thread. The control is never parented or made visible.
                IntPtr handle = target.Handle;
                if (handle == IntPtr.Zero) throw new InvalidOperationException("The event dispatch handle was not created.");
            }
            catch { target.Dispose(); throw; }
        }

        public void Invoke(Action callback)
        {
            if (callback == null) throw new ArgumentNullException(nameof(callback));
            if (disposed) throw new ObjectDisposedException(nameof(OfficeCoverageEventDispatcher));
            if (Thread.CurrentThread.ManagedThreadId == owner) callback();
            else target.Invoke(callback);
        }

        public void Dispose()
        {
            if (disposed) return;
            if (Thread.CurrentThread.ManagedThreadId != owner)
                throw new InvalidOperationException("The event dispatch handle must be disposed on its original STA.");
            // Latch before disposing: a failed native disposal must not be repeated.
            disposed = true;
            target.Dispose();
        }
    }
}
