using System;
using System.Threading;

namespace VBAi
{
    /// <summary>Temporarily suspends automatic code navigation while an explicit native inspection owns selection.</summary>
    internal sealed class VbeDebugInspection : IDisposable
    {
        [ThreadStatic] private static int depth;
        private readonly int ownerThread = Thread.CurrentThread.ManagedThreadId;
        private bool disposed;
        internal static bool IsActive => depth != 0;
        internal VbeDebugInspection() { depth++; }
        public void Dispose()
        {
            if (Thread.CurrentThread.ManagedThreadId != ownerThread)
                throw new InvalidOperationException("Debugger inspection must finish on its owning thread.");
            if (!disposed) { disposed = true; depth--; }
        }
    }
}
