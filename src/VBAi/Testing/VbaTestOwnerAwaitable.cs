using System;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VBAi
{

    /// <summary>Optional execution ownership used by the runner without imposing a UI on pure test hosts.</summary>
    internal interface IVbaTestContinuationHost
    {

        /// <summary>Wraps completion so the awaiting runner resumes on the host's owning thread.</summary>
        /// <typeparam name="T">Task result type.</typeparam>
        /// <param name="task">Host operation to observe.</param>
        /// <returns>An awaitable that dispatches continuations and verifies owner-thread access.</returns>
        VbaTestOwnerAwaitable<T> AwaitOwner<T>(Task<T> task);
    }

    /// <summary>Resumes success, exceptions and finally blocks through an explicit owning-thread dispatcher.</summary>
    /// <typeparam name="T">Result type produced by the wrapped operation.</typeparam>
    internal struct VbaTestOwnerAwaitable<T>
    {

        /// <summary>Task whose completion or exception is propagated by this awaitable.</summary>
        private readonly Task<T> task;

        /// <summary>Optional dispatcher that posts continuations to the owning UI/host thread.</summary>
        private readonly Action<Action> post;

        /// <summary>Optional assertion called when the awaited result is retrieved.</summary>
        private readonly Action requireOwner;

        /// <summary>Initializes a VbaTestOwnerAwaitable instance with the supplied state.</summary>
        /// <param name="task">Operation whose completion is observed.</param>
        /// <param name="post">Dispatcher used after asynchronous completion.</param>
        /// <param name="requireOwner">Owner-thread check executed by <see cref="Awaiter.GetResult"/>.</param>
        /// <exception cref="ArgumentNullException">Any required argument is null.</exception>
        internal VbaTestOwnerAwaitable(Task<T> task, Action<Action> post, Action requireOwner)
        {
            this.task = task ?? throw new ArgumentNullException(nameof(task));
            this.post = post ?? throw new ArgumentNullException(nameof(post));
            this.requireOwner = requireOwner ?? throw new ArgumentNullException(nameof(requireOwner));
        }

        /// <summary>Initializes a VbaTestOwnerAwaitable instance with the supplied state.</summary>
        /// <param name="task">Task to await without an owning-thread dispatcher.</param>
        private VbaTestOwnerAwaitable(Task<T> task)
        { this.task = task ?? throw new ArgumentNullException(nameof(task)); post = null; requireOwner = null; }

        /// <summary>Creates an awaitable for hosts that do not require a special continuation thread.</summary>
        /// <param name="task">Task to await.</param>
        /// <returns>Awaitable that propagates task completion directly.</returns>
        internal static VbaTestOwnerAwaitable<T> Unowned(Task<T> task) => new VbaTestOwnerAwaitable<T>(task);

        /// <summary>Returns the compiler awaiter carrying the task and optional owner dispatcher.</summary>
        /// <returns>Awaiter used by the C# await pattern.</returns>
        public Awaiter GetAwaiter() => new Awaiter(task, post, requireOwner);

        /// <summary>Carries the awaiter values passed between operations.</summary>
        internal struct Awaiter : ICriticalNotifyCompletion
        {

            /// <summary>Task whose completion this awaiter observes.</summary>
            private readonly Task<T> task;

            /// <summary>Dispatcher for the continuation, or null for an unowned await.</summary>
            private readonly Action<Action> post;

            /// <summary>Owner-thread assertion run before returning the result.</summary>
            private readonly Action requireOwner;

            /// <summary>Initializes a Awaiter instance with the supplied state.</summary>
            /// <param name="task">Observed operation.</param>
            /// <param name="post">Optional continuation dispatcher.</param>
            /// <param name="requireOwner">Optional thread-affinity assertion.</param>
            internal Awaiter(Task<T> task, Action<Action> post, Action requireOwner)
            { this.task = task; this.post = post; this.requireOwner = requireOwner; }

            /// <summary>Gets whether the wrapped task has completed and can be consumed synchronously.</summary>
            /// <value>The task's completion state.</value>
            public bool IsCompleted => task.IsCompleted;

            /// <summary>Verifies owner-thread affinity, then returns or rethrows the task result.</summary>
            /// <returns>The task result.</returns>
            public T GetResult() { requireOwner?.Invoke(); return task.GetAwaiter().GetResult(); }

            /// <summary>Registers a continuation, dispatching it through the owning-thread callback when configured.</summary>
            /// <param name="continuation">Compiler-generated continuation; null is rejected.</param>
            public void OnCompleted(Action continuation)
            {
                if (continuation == null) throw new ArgumentNullException(nameof(continuation));
                if (post == null) { task.GetAwaiter().OnCompleted(continuation); return; }
                var dispatch = post;
                task.ConfigureAwait(false).GetAwaiter().OnCompleted(() => dispatch(continuation));
            }

            /// <summary>Registers an unsafe continuation using the same optional owner-thread dispatch.</summary>
            /// <param name="continuation">Compiler-generated continuation; null is rejected.</param>
            public void UnsafeOnCompleted(Action continuation)
            {
                if (continuation == null) throw new ArgumentNullException(nameof(continuation));
                if (post == null) { task.GetAwaiter().UnsafeOnCompleted(continuation); return; }
                var dispatch = post;
                task.ConfigureAwait(false).GetAwaiter().UnsafeOnCompleted(() => dispatch(continuation));
            }
        }
    }

    /// <summary>Provides a dispatcher-backed awaitable for the VBE thread, independent of SynchronizationContext.</summary>
    internal sealed partial class VbeTestExplorerService : IVbaTestContinuationHost
    {
        // This handle outlives the external UI dispatcher while an active run settles.
        /// <summary>Control handle retained until active native execution has settled, so continuations remain dispatchable.</summary>
        private Control continuationDispatcher;

        /// <summary>Creates the hidden dispatcher control on the service's owning thread.</summary>
        /// <param name="createControl">Factory for the control whose handle receives continuations.</param>
        private void InitializeOwnerContinuations(Func<Control> createControl)
        {
            RequireContinuationOwner();
            var control = createControl();
            try { var handle = control.Handle; continuationDispatcher = control; }
            catch { control.Dispose(); throw; }
        }

        /// <summary>Throws if the caller is not on the thread that owns VBE test execution.</summary>
        private void RequireContinuationOwner()
        {
            if (Thread.CurrentThread.ManagedThreadId != ownerThread)
                throw new InvalidOperationException("The test continuation must resume on its owning VBE thread.");
        }

        /// <summary>Disposes the dispatcher only after shutdown has started and no run remains active.</summary>
        private void ReleaseOwnerContinuationsWhenIdle()
        {
            RequireContinuationOwner();
            if (!disposed || active != null) return;
            var control = continuationDispatcher;
            continuationDispatcher = null;
            control?.Dispose();
        }

        /// <summary>Posts each asynchronous continuation to the retained dispatcher and checks VBE-thread ownership on result access.</summary>
        /// <typeparam name="T">Task result type.</typeparam>
        /// <param name="task">Host operation to await.</param>
        /// <returns>Owner-thread awaitable that remains usable when native VBA clears the synchronization context.</returns>
        public VbaTestOwnerAwaitable<T> AwaitOwner<T>(Task<T> task)
            => new VbaTestOwnerAwaitable<T>(task, action => continuationDispatcher.BeginInvoke(action), RequireContinuationOwner);
    }
}
