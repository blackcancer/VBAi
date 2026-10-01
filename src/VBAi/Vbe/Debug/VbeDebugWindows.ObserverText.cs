using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace VBAi
{
    internal static partial class VbeDebugWindows
    {
        private static readonly AsyncLocal<int> observerTextDepth = new AsyncLocal<int>();
        internal delegate IntPtr ObserverTextMessage(IntPtr handle, uint message, IntPtr capacity, StringBuilder text,
            uint flags, uint milliseconds, out UIntPtr result);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "SendMessageTimeoutW", SetLastError = true)]
        private static extern IntPtr NativeObserverTextMessage(IntPtr handle, uint message, IntPtr capacity,
            StringBuilder text, uint flags, uint milliseconds, out UIntPtr result);
        internal static ObserverTextMessage ReadObserverTextMessage = NativeObserverTextMessage;
        internal const uint ObserverTextTimeoutMilliseconds = 500;

        /// <summary>Bounds only the worker's text reads; native COM commands and other read paths are unchanged.</summary>
        private static string ReadObserverText(IntPtr handle)
        {
            var text = new StringBuilder(512);
            UIntPtr length;
            // SMTO_BLOCK | SMTO_ABORTIFHUNG; an empty successful WM_GETTEXT reply remains valid.
            if (ReadObserverTextMessage(handle, 13, new IntPtr(text.Capacity), text, 3,
                ObserverTextTimeoutMilliseconds, out length) == IntPtr.Zero)
                throw new TimeoutException("The native observer text read did not complete within its bounded wait.");
            if (length.ToUInt64() >= (ulong)text.Capacity - 1)
                throw new InvalidOperationException("The native observer text is truncated.");
            return text.ToString();
        }
    }
}
