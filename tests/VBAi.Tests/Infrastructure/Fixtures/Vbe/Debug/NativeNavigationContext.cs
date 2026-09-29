namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Threading;

    internal sealed class NativeNavigationContext : SynchronizationContext
    {
        private readonly Queue<Action> callbacks = new Queue<Action>();
        internal int Posts;
        internal int RejectPost;
        internal int Pending => callbacks.Count;
        public override void Post(SendOrPostCallback callback, object state)
        {
            if (++Posts == RejectPost) throw new InvalidOperationException("dispatcher rejected post");
            callbacks.Enqueue(() => callback(state));
        }
        internal void RunNext() { callbacks.Dequeue()(); }
        internal void RunAll() { while (callbacks.Count > 0) RunNext(); }
    }
}
