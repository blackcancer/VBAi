using System;
using System.Threading;

namespace VBAi
{

    /// <summary>Temporarily suspends automatic code navigation while an explicit native inspection owns selection.</summary>
    internal sealed class VbeDebugInspection : IDisposable
    {

        /// <summary>Maintains the depth state for vbe debug inspection.</summary>
        [ThreadStatic] private static int depth;

        /// <summary>Maintains the owner thread state for vbe debug inspection.</summary>
        private readonly int ownerThread = Thread.CurrentThread.ManagedThreadId;

        /// <summary>Maintains the disposed state for vbe debug inspection.</summary>
        private bool disposed;

        /// <summary>Gets the is active.</summary>
        /// <value>Current is active exposed by vbe debug inspection.</value>
        internal static bool IsActive => depth != 0;

        /// <summary>Initializes a VbeDebugInspection instance with the supplied state.</summary>
        internal VbeDebugInspection() { depth++; }

        /// <summary>Disposes  for vbe debug inspection.</summary>
        public void Dispose()
        {
            if (Thread.CurrentThread.ManagedThreadId != ownerThread)
                throw new InvalidOperationException("Debugger inspection must finish on its owning thread.");
            if (!disposed) { disposed = true; depth--; }
        }
    }
}
