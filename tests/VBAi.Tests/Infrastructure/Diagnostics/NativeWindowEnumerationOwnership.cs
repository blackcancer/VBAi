using System;
using System.ComponentModel;

namespace VBAi.Tests.Integration
{
    /// <summary>Distinguishes a destroyed enumeration handle from an unreadable live window.</summary>
    internal static class NativeWindowEnumerationOwnership
    {
        internal delegate uint ReadOwner(IntPtr window, out uint processId);
        internal static uint? ReadProcessId(IntPtr window, ReadOwner readOwner, Func<int> lastError, Func<IntPtr, bool> isWindow)
        {
            if (window == IntPtr.Zero) throw new ArgumentException("An enumeration handle cannot be zero.", nameof(window));
            uint thread = readOwner(window, out uint processId);
            if (thread != 0 && processId != 0) return processId;
            int error = lastError(); // Capture before another native read can replace it.
            if (error == 1400 && !isWindow(window)) return null;
            throw new Win32Exception(error, "A live or unverified window owner could not be read during enumeration.");
        }
    }
}
