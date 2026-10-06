using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VBAi
{

    /// <summary>Pins one native async operation to its originating VBE STA, independently of host context.</summary>
    internal static class VbeUiTask
    {

        /// <summary>Runs  for vbe ui task.</summary>
        /// <typeparam name="T">The type used for t.</typeparam>
        /// <param name="operation">func&lt;task&lt;t&gt;&gt; that supplies the operation for this operation.</param>
        /// <returns>task&lt;t&gt; produced by the operation for run on vbe ui task.</returns>
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

        /// <summary>Handles complete for vbe ui task.</summary>
        /// <typeparam name="T">The type used for t.</typeparam>
        /// <param name="operation">func&lt;task&lt;t&gt;&gt; that supplies the operation for this operation.</param>
        /// <param name="dispatcher">control that supplies the dispatcher for this operation.</param>
        /// <returns>task&lt;t&gt; produced by the operation for complete on vbe ui task.</returns>
        private static async Task<T> Complete<T>(Func<Task<T>> operation, Control dispatcher)
        {
            try { return await operation(); }
            finally { dispatcher.Dispose(); }
        }

        /// <summary>Owns the context state and operations.</summary>
        private sealed class Context : SynchronizationContext
        {

            /// <summary>Maintains the dispatcher state for context.</summary>
            private readonly Control dispatcher;

            /// <summary>Maintains the trace state for context.</summary>
            private readonly VbeInspectionTrace trace;

            /// <summary>Initializes a Context instance with the supplied state.</summary>
            /// <param name="dispatcher">control that supplies the dispatcher for this operation.</param>
            internal Context(Control dispatcher) { this.dispatcher = dispatcher; trace = VbeInspectionTrace.Current; }

            /// <summary>Creates copy for context.</summary>
            /// <returns>synchronization context produced by the operation for create copy on context.</returns>
            public override SynchronizationContext CreateCopy() { return this; }

            /// <summary>Handles post for context.</summary>
            /// <param name="callback">send or post callback that supplies the callback for this operation.</param>
            /// <param name="state">object that supplies the state for this operation.</param>
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

            /// <summary>Handles send for context.</summary>
            /// <param name="callback">send or post callback that supplies the callback for this operation.</param>
            /// <param name="state">object that supplies the state for this operation.</param>
            public override void Send(SendOrPostCallback callback, object state)
            {
                if (dispatcher.InvokeRequired) dispatcher.Invoke(new Action(() => Invoke(callback, state)));
                else Invoke(callback, state);
            }

            /// <summary>Invokes  for context.</summary>
            /// <param name="callback">send or post callback that supplies the callback for this operation.</param>
            /// <param name="state">object that supplies the state for this operation.</param>
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
