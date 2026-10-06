using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VBAi
{

    /// <summary>Pins one native async operation to its originating VBE STA, independently of host context.</summary>
    internal static class VbeUiTask
    {

        /// <summary>Starts asynchronous native work on the caller's STA and installs a WinForms-backed continuation context.</summary>
        /// <typeparam name="T">Result type produced by the asynchronous operation.</typeparam>
        /// <param name="operation">Work that accesses VBE-owned COM objects and must resume on the originating STA.</param>
        /// <returns>A task for the operation; the temporary dispatcher is disposed when it completes.</returns>
        /// <exception cref="InvalidOperationException">The caller is not running on an STA thread.</exception>
        internal static Task<T> Run<T>(Func<Task<T>> operation)
        {
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                throw new InvalidOperationException("Native VBE work must start on its owning STA thread.");
            var previous = SynchronizationContext.Current;
            Control dispatcher = null;
            try
            {
                dispatcher = new Control();
                VbeInspectionTrace.Current?.Record(VbeInspectionTrace.Phase.OwnerSta);
                var handle = dispatcher.Handle;
                var context = new Context(dispatcher);
                SynchronizationContext.SetSynchronizationContext(context);
                return Complete(operation, dispatcher);
            }
            catch
            {
                dispatcher?.Dispose();
                throw;
            }
            finally { SynchronizationContext.SetSynchronizationContext(previous); }
        }

        /// <summary>Awaits the operation and releases its temporary dispatcher on success or failure.</summary>
        /// <typeparam name="T">Result type returned by the operation.</typeparam>
        /// <param name="operation">STA-bound asynchronous operation.</param>
        /// <param name="dispatcher">Hidden control whose handle marshals continuations to the owning STA.</param>
        /// <returns>The operation's result or exception.</returns>
        private static async Task<T> Complete<T>(Func<Task<T>> operation, Control dispatcher)
        {
            try { return await operation(); }
            finally { dispatcher.Dispose(); }
        }

        /// <summary>Marshals asynchronous continuations to the hidden dispatcher created on the VBE owner STA.</summary>
        private sealed class Context : SynchronizationContext
        {

            /// <summary>Control handle that posts callbacks to the originating WinForms thread.</summary>
            private readonly Control dispatcher;

            /// <summary>Optional inspection trace captured when this context is created.</summary>
            private readonly VbeInspectionTrace trace;

            /// <summary>Creates a synchronization context backed by the supplied dispatcher.</summary>
            /// <param name="dispatcher">Control whose handle belongs to the VBE owner STA.</param>
            internal Context(Control dispatcher) { this.dispatcher = dispatcher; trace = VbeInspectionTrace.Current; }

            /// <summary>Returns this context because it has no copy-specific mutable state.</summary>
            /// <returns>The same dispatcher-bound context.</returns>
            public override SynchronizationContext CreateCopy() { return this; }

            /// <summary>Queues an asynchronous continuation on the owning STA and records enqueue/post diagnostics.</summary>
            /// <param name="callback">Continuation to invoke on the dispatcher thread.</param>
            /// <param name="state">State passed unchanged to the continuation.</param>
            public override void Post(SendOrPostCallback callback, object state)
            {
                trace?.Record(VbeInspectionTrace.Phase.ContinuationEnqueued);
                try
                {
                    dispatcher.BeginInvoke(new Action(() => Invoke(callback, state)));
                    trace?.Record(VbeInspectionTrace.Phase.ContinuationPostReturned);
                }
                catch (Exception error)
                {
                    trace?.Record(VbeInspectionTrace.Phase.ContinuationPostFailed, error);
                    throw;
                }
            }

            /// <summary>Invokes a callback synchronously on the owner STA, marshaling only when called from another thread.</summary>
            /// <param name="callback">Callback to invoke.</param>
            /// <param name="state">State passed unchanged to the callback.</param>
            public override void Send(SendOrPostCallback callback, object state)
            {
                if (dispatcher.InvokeRequired) dispatcher.Invoke(new Action(() => Invoke(callback, state)));
                else Invoke(callback, state);
            }

            /// <summary>Installs this context for the callback duration, records entry/return, and restores the prior context.</summary>
            /// <param name="callback">Continuation or synchronous callback to run.</param>
            /// <param name="state">Callback state supplied by the synchronization caller.</param>
            private void Invoke(SendOrPostCallback callback, object state)
            {
                var previous = Current;
                using (trace?.Enter())
                {
                    trace?.Record(VbeInspectionTrace.Phase.ContinuationEntered);
                    try { SetSynchronizationContext(this); callback(state); }
                    finally { SetSynchronizationContext(previous); trace?.Record(VbeInspectionTrace.Phase.ContinuationReturned); }
                }
            }
        }
    }
}
