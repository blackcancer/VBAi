using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VBAi
{

    /// <summary>Tracks exclusive Git modal leases by their exact native owner HWND.</summary>
    internal sealed partial class GitModalSession
    {

        /// <summary>Process-local map preventing simultaneous Git dialogs for the same owner window.</summary>
        private static readonly Dictionary<IntPtr, OwnerLease> owners = new Dictionary<IntPtr, OwnerLease>();

        /// <summary>Gets the has active owner.</summary>
        /// <value>Current has active owner exposed by git modal session.</value>
        internal static bool HasActiveOwner { get { lock (owners) return owners.Count != 0; } }

        /// <summary>Checks whether the supplied native owner already holds a Git modal lease.</summary>
        /// <param name="owner">Owner window whose HWND is used as the lease key.</param>
        /// <returns><see langword="true"/> while that exact handle has an active lease.</returns>
        internal static bool IsActive(IWin32Window owner) { lock (owners) return owner != null && owners.ContainsKey(owner.Handle); }

        /// <summary>Validates and atomically acquires an exclusive lease for the supplied live owner HWND.</summary>
        /// <param name="owner">Native owner whose process generation and STA thread will be captured.</param>
        /// <returns>A lease for the owner, or <see langword="null"/> when that owner already has a session.</returns>
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

        /// <summary>Retains one owner identity until its modal runner has completely finished.</summary>
        internal sealed class OwnerLease : IDisposable
        {

            /// <summary>Original WinForms/native owner passed to the modal display call.</summary>
            private readonly IWin32Window owner;

            /// <summary>HWND key reserved in the process-wide owner map.</summary>
            private readonly IntPtr handle;

            /// <summary>Closure that rechecks owner HWND, process generation, and owning STA.</summary>
            private readonly Action validate;

            /// <summary>One-shot lifecycle flags preventing concurrent starts or early release during launch.</summary>
            private bool started, launching, disposed;

            /// <summary>Modal and handoff task retaining this lease until the session terminates.</summary>
            private Task running;

            /// <summary>Creates a lease for one already-validated owner HWND.</summary>
            /// <param name="owner">Original modal owner.</param><param name="handle">Exact HWND reserved in the shared map.</param>
            /// <param name="validate">Owner identity check repeated around the modal and import handoff.</param>
            internal OwnerLease(IWin32Window owner, IntPtr handle, Action validate)
            { this.owner = owner; this.handle = handle; this.validate = validate; }

            /// <summary>Starts the one allowed modal session while retaining the owner lease for its full task lifetime.</summary>
            /// <param name="window">Git form to display.</param><param name="show">Synchronous modal display function.</param>
            /// <returns>The task that completes after the modal return and any admitted operation finish.</returns>
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
                    if (!owners.TryGetValue(handle, out OwnerLease current) || !ReferenceEquals(current, this))
                        throw new InvalidOperationException("The exact owner lease changed.");
                    owners.Remove(handle); disposed = true;
                }
            }
        }
    }
}
