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

        /// <summary>Handles await owner for i vba test continuation host.</summary>
        /// <typeparam name="T">The type used for t.</typeparam>
        /// <param name="task">task&lt;t&gt; that supplies the task for this operation.</param>
        /// <returns>vba test owner awaitable&lt;t&gt; produced by the operation for await owner on i vba test continuation host.</returns>
        VbaTestOwnerAwaitable<T> AwaitOwner<T>(Task<T> task);
    }

    /// <summary>Resumes success, exceptions and finally blocks through an explicit owning-thread dispatcher.</summary>
    /// <typeparam name="T">The type used for t.</typeparam>
    internal struct VbaTestOwnerAwaitable<T>
    {

        /// <summary>Maintains the task state for vba test owner awaitable.</summary>
        private readonly Task<T> task;

        /// <summary>Maintains the post state for vba test owner awaitable.</summary>
        private readonly Action<Action> post;

        /// <summary>Maintains the require owner state for vba test owner awaitable.</summary>
        private readonly Action requireOwner;

        /// <summary>Initializes a VbaTestOwnerAwaitable instance with the supplied state.</summary>
        /// <param name="task">task&lt;t&gt; that supplies the task for this operation.</param>
        /// <param name="post">action&lt;action&gt; that supplies the post for this operation.</param>
        /// <param name="requireOwner">action that supplies the require owner for this operation.</param>
        internal VbaTestOwnerAwaitable(Task<T> task, Action<Action> post, Action requireOwner)
        {
            this.task = task ?? throw new ArgumentNullException(nameof(task));
            this.post = post ?? throw new ArgumentNullException(nameof(post));
            this.requireOwner = requireOwner ?? throw new ArgumentNullException(nameof(requireOwner));
        }

        /// <summary>Initializes a VbaTestOwnerAwaitable instance with the supplied state.</summary>
        /// <param name="task">task&lt;t&gt; that supplies the task for this operation.</param>
        private VbaTestOwnerAwaitable(Task<T> task)
        { this.task = task ?? throw new ArgumentNullException(nameof(task)); post = null; requireOwner = null; }

        /// <summary>Handles unowned for vba test owner awaitable.</summary>
        /// <param name="task">task&lt;t&gt; that supplies the task for this operation.</param>
        /// <returns>vba test owner awaitable&lt;t&gt; produced by the operation for unowned on vba test owner awaitable.</returns>
        internal static VbaTestOwnerAwaitable<T> Unowned(Task<T> task) => new VbaTestOwnerAwaitable<T>(task);

        /// <summary>Returns awaiter for vba test owner awaitable.</summary>
        /// <returns>awaiter produced by the operation for get awaiter on vba test owner awaitable.</returns>
        public Awaiter GetAwaiter() => new Awaiter(task, post, requireOwner);

        /// <summary>Carries the awaiter values passed between operations.</summary>
        internal struct Awaiter : ICriticalNotifyCompletion
        {

            /// <summary>Maintains the task state for awaiter.</summary>
            private readonly Task<T> task;

            /// <summary>Maintains the post state for awaiter.</summary>
            private readonly Action<Action> post;

            /// <summary>Maintains the require owner state for awaiter.</summary>
            private readonly Action requireOwner;

            /// <summary>Initializes a Awaiter instance with the supplied state.</summary>
            /// <param name="task">task&lt;t&gt; that supplies the task for this operation.</param>
            /// <param name="post">action&lt;action&gt; that supplies the post for this operation.</param>
            /// <param name="requireOwner">action that supplies the require owner for this operation.</param>
            internal Awaiter(Task<T> task, Action<Action> post, Action requireOwner)
            { this.task = task; this.post = post; this.requireOwner = requireOwner; }

            /// <summary>Gets the is completed.</summary>
            /// <value>Current is completed exposed by awaiter.</value>
            public bool IsCompleted => task.IsCompleted;

            /// <summary>Returns result for awaiter.</summary>
            /// <returns>t produced by the operation for get result on awaiter.</returns>
            public T GetResult() { requireOwner?.Invoke(); return task.GetAwaiter().GetResult(); }

            /// <summary>Handles on completed for awaiter.</summary>
            /// <param name="continuation">action that supplies the continuation for this operation.</param>
            public void OnCompleted(Action continuation)
            {
                if (continuation == null) throw new ArgumentNullException(nameof(continuation));
                if (post == null) { task.GetAwaiter().OnCompleted(continuation); return; }
                var dispatch = post;
                task.ConfigureAwait(false).GetAwaiter().OnCompleted(() => dispatch(continuation));
            }

            /// <summary>Handles unsafe on completed for awaiter.</summary>
            /// <param name="continuation">action that supplies the continuation for this operation.</param>
            public void UnsafeOnCompleted(Action continuation)
            {
                if (continuation == null) throw new ArgumentNullException(nameof(continuation));
                if (post == null) { task.GetAwaiter().UnsafeOnCompleted(continuation); return; }
                var dispatch = post;
                task.ConfigureAwait(false).GetAwaiter().UnsafeOnCompleted(() => dispatch(continuation));
            }
        }
    }

    /// <summary>Owns the vbe test explorer service state and operations.</summary>
    internal sealed partial class VbeTestExplorerService : IVbaTestContinuationHost
    {
        // This handle outlives the external UI dispatcher while an active run settles.
        /// <summary>Maintains the continuation dispatcher state for vbe test explorer service.</summary>
        private Control continuationDispatcher;

        /// <summary>Handles initialize owner continuations for vbe test explorer service.</summary>
        /// <param name="createControl">func&lt;control&gt; that supplies the create control for this operation.</param>
        private void InitializeOwnerContinuations(Func<Control> createControl)
        {
            RequireContinuationOwner();
            var control = createControl();
            try { var handle = control.Handle; continuationDispatcher = control; }
            catch { control.Dispose(); throw; }
        }

        /// <summary>Requires continuation owner for vbe test explorer service.</summary>
        private void RequireContinuationOwner()
        {
            if (Thread.CurrentThread.ManagedThreadId != ownerThread)
                throw new InvalidOperationException("The test continuation must resume on its owning VBE thread.");
        }

        /// <summary>Releases owner continuations when idle for vbe test explorer service.</summary>
        private void ReleaseOwnerContinuationsWhenIdle()
        {
            RequireContinuationOwner();
            if (!disposed || active != null) return;
            var control = continuationDispatcher;
            continuationDispatcher = null;
            control?.Dispose();
        }

        /// <summary>Does not depend on SynchronizationContext, which a native VBA command can clear.</summary>
        /// <typeparam name="T">The type used for t.</typeparam>
        /// <param name="task">task&lt;t&gt; that supplies the task for this operation.</param>
        /// <returns>vba test owner awaitable&lt;t&gt; produced by the operation for await owner on vbe test explorer service.</returns>
        public VbaTestOwnerAwaitable<T> AwaitOwner<T>(Task<T> task)
            => new VbaTestOwnerAwaitable<T>(task, action => continuationDispatcher.BeginInvoke(action), RequireContinuationOwner);
    }
}
