using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VBAi
{

    /// <summary>Owns the git modal session state and operations.</summary>
    internal sealed partial class GitModalSession
    {

        /// <summary>Maintains the owners state for git modal session.</summary>
        private static readonly Dictionary<IntPtr, OwnerLease> owners = new Dictionary<IntPtr, OwnerLease>();

        /// <summary>Gets the has active owner.</summary>
        /// <value>Current has active owner exposed by git modal session.</value>
        internal static bool HasActiveOwner { get { lock (owners) return owners.Count != 0; } }

        /// <summary>Determines whether active for git modal session.</summary>
        /// <param name="owner">i win32 window that supplies the owner for this operation.</param>
        /// <returns>Boolean indicating the result of the check for is active on git modal session.</returns>
        internal static bool IsActive(IWin32Window owner) { lock (owners) return owner != null && owners.ContainsKey(owner.Handle); }

        /// <summary>Attempts to acquire for git modal session.</summary>
        /// <param name="owner">i win32 window that supplies the owner for this operation.</param>
        /// <returns>owner lease produced by the operation for try acquire on git modal session.</returns>
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

        /// <summary>Owns the owner lease state and operations.</summary>
        internal sealed class OwnerLease : IDisposable
        {

            /// <summary>Maintains the owner state for owner lease.</summary>
            private readonly IWin32Window owner;

            /// <summary>Maintains the handle state for owner lease.</summary>
            private readonly IntPtr handle;

            /// <summary>Maintains the validate state for owner lease.</summary>
            private readonly Action validate;

            /// <summary>Maintains the started and launching and disposed state for owner lease.</summary>
            private bool started, launching, disposed;

            /// <summary>Maintains the running state for owner lease.</summary>
            private Task running;

            /// <summary>Initializes a OwnerLease instance with the supplied state.</summary>
            /// <param name="owner">i win32 window that supplies the owner for this operation.</param>
            /// <param name="handle">Native handle that supplies the handle for this operation.</param>
            /// <param name="validate">action that supplies the validate for this operation.</param>
            internal OwnerLease(IWin32Window owner, IntPtr handle, Action validate)
            { this.owner = owner; this.handle = handle; this.validate = validate; }

            /// <summary>Handles show async for owner lease.</summary>
            /// <param name="window">git window that supplies the window for this operation.</param>
            /// <param name="show">func&lt;form, i win32 window, dialog result&gt; that supplies the show for this operation.</param>
            /// <returns>task produced by the operation for show async on owner lease.</returns>
            internal Task ShowAsync(GitWindow window, Func<Form, IWin32Window, DialogResult> show)
            {
                if (disposed || started) throw new InvalidOperationException("One exact owner session only.");
                started = true;
                launching = true;
                try { return running = ShowOwnedAsync(window, owner, show, validate); }
                finally { launching = false; }
            }

            /// <summary>Disposes  for owner lease.</summary>
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
