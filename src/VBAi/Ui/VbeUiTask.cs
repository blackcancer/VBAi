using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VBAi
{

    /// <summary>Pins one native async operation to its originating VBE STA, independently of host context.</summary>
    internal static class VbeUiTask
    {

        /// <summary>Runs an asynchronous native operation with continuations marshaled to the calling STA.</summary>
        /// <typeparam name="T">Result type produced by the operation.</typeparam>
        /// <param name="operation">Asynchronous work that must start and resume on the owning VBE STA.</param>
        /// <returns>A task that completes with the operation result and disposes its message dispatcher.</returns>
        /// <exception cref="InvalidOperationException">The caller is not on an STA thread.</exception>
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

        /// <summary>Awaits the operation and disposes the dispatcher whether it succeeds or faults.</summary>
        /// <typeparam name="T">Result type produced by the operation.</typeparam>
        /// <param name="operation">Asynchronous work started with the VBE synchronization context installed.</param>
        /// <param name="dispatcher">Hidden WinForms control whose handle accepts STA continuations.</param>
        /// <returns>The operation result or its original exception.</returns>
        private static async Task<T> Complete<T>(Func<Task<T>> operation, Control dispatcher)
        {
            try { return await operation(); }
            finally { dispatcher.Dispose(); }
        }

        /// <summary>Synchronization context that posts callbacks to the hidden control on its creating STA.</summary>
        private sealed class Context : SynchronizationContext
        {

            /// <summary>WinForms control whose handle is bound to the owning VBE STA.</summary>
            private readonly Control dispatcher;

            /// <summary>Optional trace captured at context creation for continuation lifecycle events.</summary>
            private readonly VbeInspectionTrace trace;

            /// <summary>Creates a synchronization context backed by the STA dispatcher and current inspection trace.</summary>
            /// <param name="dispatcher">Control whose handle receives posted callbacks.</param>
            internal Context(Control dispatcher) { this.dispatcher = dispatcher; trace = VbeInspectionTrace.Current; }

            /// <summary>Returns this context because its dispatcher and STA affinity are shared.</summary>
            /// <returns>This synchronization context instance.</returns>
            public override SynchronizationContext CreateCopy() { return this; }

            /// <summary>Queues a callback on the dispatcher handle and records whether posting entered or failed.</summary>
            /// <param name="callback">Callback to invoke on the owning STA.</param>
            /// <param name="state">State object passed unchanged to <paramref name="callback"/>.</param>
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

            /// <summary>Invokes synchronously on the owning STA, marshaling through the dispatcher when needed.</summary>
            /// <param name="callback">Callback to invoke on the owning STA.</param>
            /// <param name="state">State object passed unchanged to <paramref name="callback"/>.</param>
            public override void Send(SendOrPostCallback callback, object state)
            {
                if (dispatcher.InvokeRequired) dispatcher.Invoke(new Action(() => Invoke(callback, state)));
                else Invoke(callback, state);
            }

            /// <summary>Installs this context during callback execution and restores the previous context afterward.</summary>
            /// <param name="callback">Callback to execute.</param>
            /// <param name="state">State object passed unchanged to the callback.</param>
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
