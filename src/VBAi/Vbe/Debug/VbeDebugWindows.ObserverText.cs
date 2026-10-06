using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace VBAi
{

    /// <summary>Owns the vbe debug windows state and operations.</summary>
    internal static partial class VbeDebugWindows
    {

        /// <summary>Maintains the observer text depth state for vbe debug windows.</summary>
        private static readonly AsyncLocal<int> observerTextDepth = new AsyncLocal<int>();

        /// <summary>Defines the observer text message callback.</summary>
        /// <param name="handle">Native handle that supplies the handle for this operation.</param>
        /// <param name="message">uint that supplies the message for this operation.</param>
        /// <param name="capacity">Native handle that supplies the capacity for this operation.</param>
        /// <param name="text">string builder that supplies the text for this operation.</param>
        /// <param name="flags">uint that supplies the flags for this operation.</param>
        /// <param name="milliseconds">uint that supplies the milliseconds for this operation.</param>
        /// <param name="result">Native handle that supplies the result for this operation.</param>
        /// <returns>int ptr produced by the operation for operation on vbe debug windows.</returns>
        internal delegate IntPtr ObserverTextMessage(IntPtr handle, uint message, IntPtr capacity, StringBuilder text,
            uint flags, uint milliseconds, out UIntPtr result);

        /// <summary>Handles native observer text message for vbe debug windows.</summary>
        /// <param name="handle">Native handle that supplies the handle for this operation.</param>
        /// <param name="message">uint that supplies the message for this operation.</param>
        /// <param name="capacity">Native handle that supplies the capacity for this operation.</param>
        /// <param name="text">string builder that supplies the text for this operation.</param>
        /// <param name="flags">uint that supplies the flags for this operation.</param>
        /// <param name="milliseconds">uint that supplies the milliseconds for this operation.</param>
        /// <param name="result">Native handle that supplies the result for this operation.</param>
        /// <returns>int ptr produced by the operation for native observer text message on vbe debug windows.</returns>
        [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "SendMessageTimeoutW", SetLastError = true)]
        private static extern IntPtr NativeObserverTextMessage(IntPtr handle, uint message, IntPtr capacity,
            StringBuilder text, uint flags, uint milliseconds, out UIntPtr result);

        /// <summary>Maintains the read observer text message state for vbe debug windows.</summary>
        internal static ObserverTextMessage ReadObserverTextMessage = NativeObserverTextMessage;

        /// <summary>Maintains the observer text timeout milliseconds state for vbe debug windows.</summary>
        internal const uint ObserverTextTimeoutMilliseconds = 500;

        /// <summary>Bounds only the worker's text reads; native COM commands and other read paths are unchanged.</summary>
        /// <param name="handle">Native handle that supplies the handle for this operation.</param>
        /// <returns>Text produced by the operation for read observer text on vbe debug windows.</returns>
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
