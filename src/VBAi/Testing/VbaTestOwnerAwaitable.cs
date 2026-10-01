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
        VbaTestOwnerAwaitable<T> AwaitOwner<T>(Task<T> task);
    }

    /// <summary>Resumes success, exceptions and finally blocks through an explicit owning-thread dispatcher.</summary>
    internal struct VbaTestOwnerAwaitable<T>
    {
        private readonly Task<T> task;
        private readonly Action<Action> post;
        private readonly Action requireOwner;

        internal VbaTestOwnerAwaitable(Task<T> task, Action<Action> post, Action requireOwner)
        {
            this.task = task ?? throw new ArgumentNullException(nameof(task));
            this.post = post ?? throw new ArgumentNullException(nameof(post));
            this.requireOwner = requireOwner ?? throw new ArgumentNullException(nameof(requireOwner));
        }

        private VbaTestOwnerAwaitable(Task<T> task)
        { this.task = task ?? throw new ArgumentNullException(nameof(task)); post = null; requireOwner = null; }

        internal static VbaTestOwnerAwaitable<T> Unowned(Task<T> task) => new VbaTestOwnerAwaitable<T>(task);
        public Awaiter GetAwaiter() => new Awaiter(task, post, requireOwner);

        internal struct Awaiter : ICriticalNotifyCompletion
        {
            private readonly Task<T> task;
            private readonly Action<Action> post;
            private readonly Action requireOwner;
            internal Awaiter(Task<T> task, Action<Action> post, Action requireOwner)
            { this.task = task; this.post = post; this.requireOwner = requireOwner; }
            public bool IsCompleted => task.IsCompleted;
            public T GetResult() { requireOwner?.Invoke(); return task.GetAwaiter().GetResult(); }
            public void OnCompleted(Action continuation)
            {
                if (continuation == null) throw new ArgumentNullException(nameof(continuation));
                if (post == null) { task.GetAwaiter().OnCompleted(continuation); return; }
                var dispatch = post;
                task.ConfigureAwait(false).GetAwaiter().OnCompleted(() => dispatch(continuation));
            }
            public void UnsafeOnCompleted(Action continuation)
            {
                if (continuation == null) throw new ArgumentNullException(nameof(continuation));
                if (post == null) { task.GetAwaiter().UnsafeOnCompleted(continuation); return; }
                var dispatch = post;
                task.ConfigureAwait(false).GetAwaiter().UnsafeOnCompleted(() => dispatch(continuation));
            }
        }
    }

    internal sealed partial class VbeTestExplorerService : IVbaTestContinuationHost
    {
        // This handle outlives the external UI dispatcher while an active run settles.
        private Control continuationDispatcher;

        private void InitializeOwnerContinuations()
        {
            RequireContinuationOwner();
            var control = new Control();
            try { var handle = control.Handle; continuationDispatcher = control; }
            catch { control.Dispose(); throw; }
        }

        private void RequireContinuationOwner()
        {
            if (Thread.CurrentThread.ManagedThreadId != ownerThread)
                throw new InvalidOperationException("The test continuation must resume on its owning VBE thread.");
        }

        private void ReleaseOwnerContinuationsWhenIdle()
        {
            RequireContinuationOwner();
            if (!disposed || active != null) return;
            var control = continuationDispatcher;
            continuationDispatcher = null;
            control?.Dispose();
        }

        /// <summary>Does not depend on SynchronizationContext, which a native VBA command can clear.</summary>
        public VbaTestOwnerAwaitable<T> AwaitOwner<T>(Task<T> task)
            => new VbaTestOwnerAwaitable<T>(task, action => continuationDispatcher.BeginInvoke(action), RequireContinuationOwner);
    }
}
