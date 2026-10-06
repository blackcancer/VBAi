using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace VBAi
{

    /// <summary>Reads native observer text with a short timeout and a bounded destination buffer.</summary>
    internal static partial class VbeDebugWindows
    {

        /// <summary>Async-local recursion depth used to prevent reentrant native observer text reads.</summary>
        private static readonly AsyncLocal<int> observerTextDepth = new AsyncLocal<int>();

        /// <summary>Injectable signature for bounded SendMessageTimeout text reads.</summary>
        /// <param name="handle">Observer HWND receiving the message.</param>
        /// <param name="message">Win32 message ID, WM_GETTEXT in this reader.</param>
        /// <param name="capacity">Buffer capacity passed as WPARAM.</param>
        /// <param name="text">Unicode destination buffer.</param>
        /// <param name="flags">SendMessageTimeout flags, including block and abort-if-hung.</param>
        /// <param name="milliseconds">Maximum native wait duration.</param>
        /// <param name="result">Receives the message-specific result.</param>
        /// <returns>Nonzero when the message completed before timeout; zero on timeout/failure.</returns>
        internal delegate IntPtr ObserverTextMessage(IntPtr handle, uint message, IntPtr capacity, StringBuilder text,
            uint flags, uint milliseconds, out UIntPtr result);

        /// <summary>Sends a Unicode Win32 message with a bounded wait, used here for WM_GETTEXT.</summary>
        /// <param name="handle">Observer HWND.</param>
        /// <param name="message">Win32 message ID.</param>
        /// <param name="capacity">Character capacity passed as WPARAM.</param>
        /// <param name="text">Unicode destination buffer.</param>
        /// <param name="flags">SendMessageTimeout behavior flags.</param>
        /// <param name="milliseconds">Timeout in milliseconds.</param>
        /// <param name="result">Receives the native message result.</param>
        /// <returns>Nonzero on completion; zero on timeout or native failure.</returns>
        [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "SendMessageTimeoutW", SetLastError = true)]
        private static extern IntPtr NativeObserverTextMessage(IntPtr handle, uint message, IntPtr capacity,
            StringBuilder text, uint flags, uint milliseconds, out UIntPtr result);

        /// <summary>Native SendMessageTimeout implementation, replaceable in focused tests.</summary>
        internal static ObserverTextMessage ReadObserverTextMessage = NativeObserverTextMessage;

        /// <summary>Maximum synchronous wait for one observer WM_GETTEXT read, in milliseconds.</summary>
        internal const uint ObserverTextTimeoutMilliseconds = 500;

        /// <summary>Bounds only the worker's text reads; native COM commands and other read paths are unchanged.</summary>
        /// <param name="handle">Observer HWND whose text is read into a 512-character buffer.</param>
        /// <returns>Complete text, including an empty successful reply; timeout or truncation throws.</returns>
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
