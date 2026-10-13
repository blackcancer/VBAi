using System;
using System.Runtime.InteropServices;
using System.Text;

namespace VBAi.Tests.Integration
{
    /// <summary>Reads the displayed selection of an exact native drop-down list without traversing UIA selection peers.</summary>
    internal static class OfficeNativeComboRead
    {
        private const int MaximumLabelLength = 2048;
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr window);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
        [DllImport("user32.dll")] private static extern bool IsChild(IntPtr parent, IntPtr child);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder text, int size);
        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")] private static extern IntPtr GetWindowLongPtr(IntPtr window, int index);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SendMessageTimeout(IntPtr window, uint message, IntPtr wParam, IntPtr lParam,
            uint flags, uint timeout, out IntPtr result);
        [DllImport("user32.dll", EntryPoint = "SendMessageTimeoutW", CharSet = CharSet.Unicode)]
        private static extern IntPtr ReadTextTimeout(IntPtr window, uint message, IntPtr wParam, StringBuilder text,
            uint flags, uint timeout, out IntPtr result);

        internal sealed class TextRead
        {
            internal string Text;
            internal long Characters;
        }

        /// <summary>Rejects editable selectors, foreign or hidden peers before reading their native state.</summary>
        internal static bool IdentityMatches(int expectedPid, uint expectedThread, int actualPid, uint actualThread,
            string nativeClass, long style, bool exists, bool child, bool visible)
            => expectedPid > 0 && expectedThread > 0 && actualPid == expectedPid && actualThread == expectedThread &&
                nativeClass != null && nativeClass.StartsWith("WindowsForms", StringComparison.Ordinal) &&
                nativeClass.Contains("COMBOBOX") && (style & 3) == 3 && exists && child && visible;

        /// <summary>Uses size-bounded WM_GETTEXT, never an unbounded CB_GETLBTEXT buffer or a selection mutation.</summary>
        internal static string Read(IntPtr window, IntPtr owner, int pid, uint thread, Action requireOwner)
        {
            Action guard = () =>
            {
                requireOwner?.Invoke();
                uint actualPid; uint actualThread = GetWindowThreadProcessId(window, out actualPid);
                var nativeClass = new StringBuilder(128);
                GetClassName(window, nativeClass, nativeClass.Capacity);
                if (!IdentityMatches(pid, thread, (int)actualPid, actualThread, nativeClass.ToString(),
                    GetWindowLongPtr(window, -16).ToInt64(), IsWindow(window), IsChild(owner, window), IsWindowVisible(window)))
                    throw new InvalidOperationException("Original native drop-down list ownership changed.");
            };
            Func<uint, long> message = id =>
            {
                IntPtr result;
                if (SendMessageTimeout(window, id, IntPtr.Zero, IntPtr.Zero, 3, 3000, out result) == IntPtr.Zero)
                    throw new TimeoutException("Native combo read did not return; no repeat.");
                return unchecked((int)result.ToInt64());
            };
            return ReadBounded(guard, () => message(0x0147), () => message(0x000E), capacity =>
            {
                var buffer = new StringBuilder(capacity);
                IntPtr count;
                if (ReadTextTimeout(window, 0x000D, new IntPtr(capacity), buffer, 3, 3000, out count) == IntPtr.Zero)
                    throw new TimeoutException("Native combo text read did not return; no repeat.");
                return new TextRead { Text = buffer.ToString(), Characters = count.ToInt64() };
            });
        }

        /// <summary>Requires a stable index and text-length bound around a single bounded label read.</summary>
        internal static string ReadBounded(Action guard, Func<long> index, Func<long> length, Func<int, TextRead> text)
        {
            if (guard == null || index == null || length == null || text == null)
                throw new ArgumentNullException(nameof(guard));
            guard();
            long selected = index();
            guard();
            if (selected == -1) return "";
            if (selected < 0 || selected >= 4096) throw new InvalidOperationException("Native combo index exceeds its bound.");
            long upperBound = length();
            guard();
            if (upperBound < 0 || upperBound > MaximumLabelLength)
                throw new InvalidOperationException("Native combo label exceeds its bound.");
            TextRead result = text(MaximumLabelLength + 1);
            guard();
            if (result == null || result.Text == null || result.Characters != result.Text.Length ||
                result.Characters < 0 || result.Characters > upperBound)
                throw new InvalidOperationException("Native combo label read is incomplete or contradictory.");
            long finalIndex = index();
            guard();
            long finalLength = length();
            guard();
            if (finalIndex != selected || finalLength != upperBound)
                throw new InvalidOperationException("Native combo selection changed during observation; no repeated read.");
            return result.Text;
        }
    }
}
