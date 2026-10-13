// SPDX-License-Identifier: MPL-2.0
// Host-independent coalescing of the activity/transcript refresh pattern used by
// VBAi's native transcript (local 213231485c3bba0251357357a09392b750b8c24f).
using System;
using System.Collections.Generic;

namespace Blackcancer.ChatUI.Runtime
{
    /// <summary>
    /// Coalesces replaceable transcript refreshes by key. The host owns the timer and
    /// calls Flush on its UI thread; token deltas must be accumulated before Queue.
    /// </summary>
    public sealed class TranscriptUpdateBatch<TKey>
    {
        private readonly object gate = new object();
        private readonly Dictionary<TKey, Action> callbacks;
        private readonly List<TKey> order = new List<TKey>();

        public TranscriptUpdateBatch(IEqualityComparer<TKey> comparer = null)
        {
            callbacks = new Dictionary<TKey, Action>(comparer ?? EqualityComparer<TKey>.Default);
        }

        public int PendingCount { get { lock (gate) return callbacks.Count; } }

        /// <summary>Queues a refresh, replacing an older callback with the same key.</summary>
        public void Queue(TKey key, Action refresh)
        {
            if (ReferenceEquals(key, null)) throw new ArgumentNullException(nameof(key));
            if (refresh == null) throw new ArgumentNullException(nameof(refresh));
            lock (gate)
            {
                if (!callbacks.ContainsKey(key)) order.Add(key);
                callbacks[key] = refresh;
            }
        }

        /// <summary>
        /// Runs the current snapshot in first-key order. Reentrant Queue calls remain pending
        /// for the next flush. Failed callbacks are reported together and never replayed.
        /// </summary>
        public void Flush()
        {
            List<Action> snapshot;
            lock (gate)
            {
                snapshot = new List<Action>(order.Count);
                foreach (var key in order) snapshot.Add(callbacks[key]);
                callbacks.Clear();
                order.Clear();
            }
            List<Exception> failures = null;
            foreach (var refresh in snapshot)
            {
                try { refresh(); }
                catch (Exception error)
                {
                    if (failures == null) failures = new List<Exception>();
                    failures.Add(error);
                }
            }
            if (failures != null) throw new AggregateException("Transcript refresh failed.", failures);
        }

        /// <summary>Discards pending refreshes when the owning transcript is cleared or disposed.</summary>
        public void Clear()
        {
            lock (gate)
            {
                callbacks.Clear();
                order.Clear();
            }
        }
    }
}
