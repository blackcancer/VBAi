using System;
using System.Threading;

namespace VBAi
{

    /// <summary>Temporarily suspends automatic code navigation while an explicit native inspection owns selection.</summary>
    internal sealed class VbeDebugInspection : IDisposable
    {

        /// <summary>Per-thread nesting count for active debugger inspections.</summary>
        [ThreadStatic] private static int depth;

        /// <summary>Managed thread that must dispose this inspection scope.</summary>
        private readonly int ownerThread = Thread.CurrentThread.ManagedThreadId;

        /// <summary>Prevents repeated disposal from decrementing the thread's nesting count twice.</summary>
        private bool disposed;

        /// <summary>Reports whether automatic navigation is suspended on the current thread.</summary>
        /// <value><see langword="true"/> while one or more inspection scopes are active on this thread.</value>
        internal static bool IsActive => depth != 0;

        /// <summary>Begins a thread-affine inspection scope and suspends automatic code navigation.</summary>
        internal VbeDebugInspection() { depth++; }

        /// <summary>Ends this inspection scope once on its creating thread and resumes navigation when nesting reaches zero.</summary>
        /// <exception cref="InvalidOperationException">Disposal is attempted from a thread other than the creating thread.</exception>
        public void Dispose()
        {
            if (Thread.CurrentThread.ManagedThreadId != ownerThread)
                throw new InvalidOperationException("Debugger inspection must finish on its owning thread.");
            if (!disposed) { disposed = true; depth--; }
        }
    }
}
