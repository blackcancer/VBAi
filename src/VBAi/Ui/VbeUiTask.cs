using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VBAi
{

    /// <summary>Pins one native async operation to its originating VBE STA, independently of host context.</summary>
    internal static partial class VbeUiTask
    {

        /// <summary>Pins an operation without a result to the calling VBE STA.</summary>
        /// <param name="operation">Operation whose awaits and cleanup must remain on that STA.</param>
        /// <returns>A task preserving completion, cancellation, or the original failure.</returns>
        internal static Task RunAsync(Func<Task> operation)
        {
            if (operation == null) throw new ArgumentNullException(nameof(operation));
            return Run(async () => { await operation(); return true; });
        }

        /// <summary>Runs an asynchronous native operation with continuations marshaled to the calling STA.</summary>
        /// <typeparam name="T">Result type produced by the operation.</typeparam>
        /// <param name="operation">Asynchronous work that must start and resume on the owning VBE STA.</param>
        /// <returns>A task that completes with the operation result and disposes its message dispatcher.</returns>
        /// <exception cref="InvalidOperationException">The caller is not on an STA thread.</exception>
        internal static Task<T> Run<T>(Func<Task<T>> operation)
            => RunOwned(operation, null);

        /// <summary>Creates one dispatcher whose lifetime includes work, cleanup and synchronous terminal publication.</summary>
        /// <typeparam name="T">Result produced by the operation.</typeparam>
        /// <param name="operation">Work that remains on the calling STA.</param>
        /// <param name="completed">Optional synchronous terminal callback after dispatcher cleanup.</param>
        /// <returns>The whole operation lifetime, independent of the caller's ambient context.</returns>
        private static Task<T> RunOwned<T>(Func<Task<T>> operation, Action<Exception> completed)
        {
            if (operation == null) throw new ArgumentNullException(nameof(operation));
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
                return Complete(operation, dispatcher, context, completed);
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
        /// <param name="context">Tracks asynchronous event handlers started by the operation.</param>
        /// <param name="completed">Synchronous terminal callback, or null.</param>
        /// <returns>The operation result or its original exception.</returns>
        private static async Task<T> Complete<T>(Func<Task<T>> operation, Control dispatcher, Context context, Action<Exception> completed)
        {
            T result = default(T);
            Exception failure = null;
            try { result = await operation(); }
            catch (Exception error) { failure = error; }
            // Async-void events notify their captured context through OperationStarted/Completed.
            // Their awaits still need this handle even after the initiating task has settled.
            while (context.HasPendingOperations) await context.WaitForOperationsAsync();
            failure = FinishOwnerOperation(failure, dispatcher.Dispose, completed);
            if (failure != null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
            return result;
        }

        /// <summary>Synchronization context that posts callbacks to the hidden control on its creating STA.</summary>
        private sealed class Context : SynchronizationContext
        {

            /// <summary>WinForms control whose handle is bound to the owning VBE STA.</summary>
            private readonly Control dispatcher;

            /// <summary>Optional trace captured at context creation for continuation lifecycle events.</summary>
            private readonly VbeInspectionTrace trace;

            /// <summary>Protects asynchronous event accounting when completion arrives from another thread.</summary>
            private readonly object operationsGate = new object();
            /// <summary>Number of asynchronous event handlers whose builder still owns this context.</summary>
            private int operations;
            /// <summary>Signals that all currently owned asynchronous handlers have completed.</summary>
            private TaskCompletionSource<bool> operationsDrained;

            /// <summary>Reports whether a handler still owns this dispatcher, including admission before a drain continuation resumes.</summary>
            internal bool HasPendingOperations { get { lock (operationsGate) { return operations != 0; } } }

            /// <summary>Retains the dispatcher for an asynchronous event handler admitted on this context.</summary>
            public override void OperationStarted()
            {
                lock (operationsGate) { operations++; }
            }

            /// <summary>Releases a completed event handler and wakes owner cleanup when the last handler finishes.</summary>
            public override void OperationCompleted()
            {
                TaskCompletionSource<bool> drained = null;
                lock (operationsGate)
                {
                    if (operations == 0) throw new InvalidOperationException("No asynchronous owner operation is pending.");
                    if (--operations == 0) { drained = operationsDrained; operationsDrained = null; }
                }
                drained?.TrySetResult(true);
            }

            /// <summary>Waits for asynchronous event handlers before disposing their owner dispatcher.</summary>
            /// <returns>An already completed task when no handler is pending, or the current drain signal.</returns>
            internal Task WaitForOperationsAsync()
            {
                lock (operationsGate)
                {
                    if (operations == 0) return Task.CompletedTask;
                    if (operationsDrained == null)
                        operationsDrained = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
                    return operationsDrained.Task;
                }
            }

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
