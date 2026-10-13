using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace VBAi
{
    /// <summary>Balances COM getter acquisitions without releasing borrowed caller objects.</summary>
    internal sealed class VbaTestComReferences : IDisposable
    {
        /// <summary>One entry per getter acquisition, even when COM returns the same RCW.</summary>
        private readonly List<object> acquired = new List<object>();
        /// <summary>Release boundary called on the same owning thread as acquisition.</summary>
        private readonly Action<object> release;
        /// <summary>Selects the one-based VBIDE/Office collections whose items must be acquired without a hidden COM enumerator.</summary>
        private readonly Func<object, bool> indexedCollection;
        /// <summary>Roots scopes whose native outcome or release cannot safely be completed.</summary>
        private static readonly List<VbaTestComReferences> retained = new List<VbaTestComReferences>();
        /// <summary>Records normal completion or deliberate retention without permitting a retry.</summary>
        private bool disposed, keepAlive;

        /// <summary>Creates an owner-thread scope for explicitly acquired references.</summary>
        /// <param name="release">Optional release boundary for isolated ownership tests.</param>
        internal VbaTestComReferences(Action<object> release = null) : this(release, Marshal.IsComObject) { }

        /// <summary>Creates a reference scope with an injectable native-collection classifier.</summary>
        /// <param name="release">Balances one owned acquisition.</param>
        /// <param name="indexedCollection">Selects one-based Count/default-Item access; do not use for zero-based MSForms controls.</param>
        internal VbaTestComReferences(Action<object> release, Func<object, bool> indexedCollection)
        {
            this.release = release ?? Release;
            this.indexedCollection = indexedCollection ?? throw new ArgumentNullException(nameof(indexedCollection));
        }

        /// <summary>Records one getter acquisition, including repeated acquisitions of the same RCW.</summary>
        /// <param name="value">Newly acquired reference; never a borrowed argument.</param>
        /// <returns>The reference for immediate use on its owning thread.</returns>
        internal object Add(object value)
        {
            if (disposed || keepAlive) throw new ObjectDisposedException(nameof(VbaTestComReferences));
            if (value != null) acquired.Add(value);
            return value;
        }

        /// <summary>Enumerates an acquired collection and records its enumerator and each returned item.</summary>
        /// <param name="collection">Result of one COM collection getter.</param>
        /// <returns>Items borrowed until this scope is disposed or retained.</returns>
        internal IEnumerable<object> Items(object collection)
        {
            Add(collection);
            return ItemsFromBorrowed(collection);
        }

        /// <summary>Enumerates a collection already owned by the caller, acquiring only enumerator and items.</summary>
        /// <param name="collection">Borrowed collection; its getter acquisition must have another owner.</param>
        /// <returns>Items borrowed until this scope is disposed or retained.</returns>
        internal IEnumerable<object> ItemsFromBorrowed(object collection)
        {
            if (disposed || keepAlive) throw new ObjectDisposedException(nameof(VbaTestComReferences));
            if (indexedCollection(collection))
            {
                // .NET Framework's EnumeratorViewOfEnumVariant is a managed adapter
                // retaining its native enumerator. ReleaseComObject cannot release
                // that adapter, so acquire the known one-based collection items directly.
                int count = Convert.ToInt32(((dynamic)collection).Count);
                if (count < 0) throw new InvalidOperationException("The native collection returned a negative count.");
                for (int index = 1; index <= count; index++)
                    yield return Add((object)((dynamic)collection)[index]);
                yield break;
            }
            var enumerator = ((IEnumerable)collection).GetEnumerator();
            Add(enumerator);
            while (enumerator.MoveNext()) yield return Add(enumerator.Current);
        }

        /// <summary>Transfers one tracked acquisition to a longer-lived owner.</summary>
        /// <param name="value">Reference whose last matching acquisition is transferred.</param>
        internal void Transfer(object value)
        {
            for (int index = acquired.Count - 1; index >= 0; --index)
                if (ReferenceEquals(acquired[index], value)) { acquired.RemoveAt(index); return; }
            throw new InvalidOperationException("The transferred COM reference was not acquired by this scope.");
        }

        /// <summary>Keeps all remaining acquisitions alive after an uncertain native outcome.</summary>
        internal void Retain()
        {
            if (disposed || keepAlive) return;
            keepAlive = true;
            lock (retained) retained.Add(this);
        }

        /// <summary>Releases once per acquisition, in reverse order; a failed release is never retried.</summary>
        public void Dispose()
        {
            if (disposed || keepAlive) return;
            while (acquired.Count != 0)
            {
                int index = acquired.Count - 1;
                object value = acquired[index];
                try { release(value); }
                catch { Retain(); throw; }
                acquired.RemoveAt(index);
            }
            disposed = true;
        }

        /// <summary>Balances one RCW acquisition without finalizing other owners' references.</summary>
        /// <param name="value">Explicitly acquired COM reference; managed doubles require no release.</param>
        internal static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
            else if (value is IEnumerator && value is IDisposable iterator) iterator.Dispose();
        }
    }
}
