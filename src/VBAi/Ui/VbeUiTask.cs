using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VBAi
{
    /// <summary>Pins one native async operation to its originating VBE STA, independently of host context.</summary>
    internal static class VbeUiTask
    {
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

        private static async Task<T> Complete<T>(Func<Task<T>> operation, Control dispatcher)
        {
            try { return await operation(); }
            finally { dispatcher.Dispose(); }
        }

        private sealed class Context : SynchronizationContext
        {
            private readonly Control dispatcher;
            private readonly VbeInspectionTrace trace;
            internal Context(Control dispatcher) { this.dispatcher = dispatcher; trace = VbeInspectionTrace.Current; }
            public override SynchronizationContext CreateCopy() { return this; }
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
            public override void Send(SendOrPostCallback callback, object state)
            {
                if (dispatcher.InvokeRequired) dispatcher.Invoke(new Action(() => Invoke(callback, state)));
                else Invoke(callback, state);
            }
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
