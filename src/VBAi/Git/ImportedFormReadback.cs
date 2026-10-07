using System;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;

namespace VBAi
{
    /// <summary>Verifies one settled form import without replaying native mutations or weakening resource equality.</summary>
    internal static class ImportedFormReadback
    {
        /// <summary>Allows at most one owner continuation and one additional complete capture after a returned native import.</summary>
        /// <param name="target">Complete frozen target, including exact logical FRX bytes.</param>
        /// <param name="initial">First complete post-import capture.</param>
        /// <param name="nativeFormsReturned">True only after ordinary native form imports and their original font deliveries returned.</param>
        /// <param name="revalidate">Read-only thread, owner, component, and recovery checks before another capture.</param>
        /// <param name="capture">Complete readback on the same owner STA; never a mutation callback.</param>
        /// <param name="pulse">Optional owner-scheduling dependency; production uses one real UI timer tick.</param>
        /// <returns>Whether the complete observed state exactly matches the target.</returns>
        internal static async Task<bool> VerifyAsync(VbaGitSnapshot target, VbaGitSnapshot initial,
            bool nativeFormsReturned, Action revalidate, Func<VbaGitSnapshot> capture, Func<Task> pulse = null)
        {
            if (initial.SameAs(target)) return true;
            if (!nativeFormsReturned) return false;
            // BeginInvoke can drain a callback posted by Task.Yield inside the same
            // native dispatch. A single UI timer tick permits that dispatch to return.
            // It never polls until green or repeats an import/setter/activation.
            VbeInspectionTrace.Current?.Record(VbeInspectionTrace.Phase.ImportReadbackYieldBefore);
            await (pulse ?? (Func<Task>)PulseOwnerAsync)();
            VbeInspectionTrace.Current?.Record(VbeInspectionTrace.Phase.ImportReadbackYieldReturned);
            VbeInspectionTrace.Current?.Record(VbeInspectionTrace.Phase.ImportReadbackGuardBefore);
            revalidate();
            VbeInspectionTrace.Current?.Record(VbeInspectionTrace.Phase.ImportReadbackGuardReturned);
            VbeInspectionTrace.Current?.Record(VbeInspectionTrace.Phase.ImportReadbackCaptureBefore);
            var settled = capture();
            VbeInspectionTrace.Current?.Record(VbeInspectionTrace.Phase.ImportReadbackCaptureReturned);
            return settled.SameAs(target);
        }

        /// <summary>Permits one real owning-thread message pulse without polling the native state.</summary>
        /// <returns>Completion after the one-shot timer is stopped and disposed on its owning STA.</returns>
        internal static async Task PulseOwnerAsync()
        {
            int ownerThread = Thread.CurrentThread.ManagedThreadId;
            var context = SynchronizationContext.Current;
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA || context == null)
                throw new InvalidOperationException("Native form persistence requires an owning STA synchronization context.");
            var completion = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer { Interval = 1 };
            EventHandler tick = null;
            Action release = () =>
            {
                var owned = timer;
                timer = null;
                if (owned == null) return;
                owned.Stop();
                owned.Tick -= tick;
                owned.Dispose();
            };
            tick = (sender, arguments) =>
            {
                try
                {
                    // No COM, snapshot or guard runs in Tick. End timer ownership
                    // before posting the continuation to the captured owner context.
                    release();
                    completion.TrySetResult(true);
                }
                catch (Exception error) { completion.TrySetException(error); }
            };
            try
            {
                timer.Tick += tick;
                timer.Start();
                await completion.Task;
                if (Thread.CurrentThread.ManagedThreadId != ownerThread ||
                    Thread.CurrentThread.GetApartmentState() != ApartmentState.STA ||
                    !ReferenceEquals(SynchronizationContext.Current, context))
                    throw new InvalidOperationException("Native form persistence left its owning STA context.");
            }
            finally { release(); }
        }
    }
}
