using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VBAi
{
    internal sealed partial class GitModalSession
    {
        private static readonly Dictionary<IntPtr, OwnerLease> owners = new Dictionary<IntPtr, OwnerLease>();
        internal static bool HasActiveOwner { get { lock (owners) return owners.Count != 0; } }
        internal static bool IsActive(IWin32Window owner) { lock (owners) return owner != null && owners.ContainsKey(owner.Handle); }
        internal static OwnerLease TryAcquire(IWin32Window owner)
        {
            if (owner == null) throw new ArgumentNullException(nameof(owner));
            IntPtr handle = owner.Handle;
            lock (owners) if (owners.ContainsKey(handle)) return null;
            Action validate = CaptureOwner(owner); validate();
            lock (owners)
            {
                if (owners.ContainsKey(handle)) return null;
                var lease = new OwnerLease(owner, handle, validate); owners.Add(handle, lease); return lease;
            }
        }

        internal sealed class OwnerLease : IDisposable
        {
            private readonly IWin32Window owner;
            private readonly IntPtr handle;
            private readonly Action validate;
            private bool started, launching, disposed;
            private Task running;
            internal OwnerLease(IWin32Window owner, IntPtr handle, Action validate)
            { this.owner = owner; this.handle = handle; this.validate = validate; }
            internal Task ShowAsync(GitWindow window, Func<Form, IWin32Window, DialogResult> show)
            {
                if (disposed || started) throw new InvalidOperationException("One exact owner session only.");
                started = true;
                launching = true;
                try { return running = ShowOwnedAsync(window, owner, show, validate); }
                finally { launching = false; }
            }
            public void Dispose()
            {
                if (disposed) return;
                if (launching || running != null && !running.IsCompleted) throw new InvalidOperationException("A pending Git session retains its original owner lease.");
                lock (owners)
                {
                    OwnerLease current;
                    if (!owners.TryGetValue(handle, out current) || !ReferenceEquals(current, this))
                        throw new InvalidOperationException("The exact owner lease changed.");
                    owners.Remove(handle); disposed = true;
                }
            }
        }
    }
}
