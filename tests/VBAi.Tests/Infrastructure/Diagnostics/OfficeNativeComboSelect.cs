using System;
using System.Runtime.InteropServices;
using System.Text;

namespace VBAi.Tests.Integration
{
    /// <summary>Selects one actual WinForms choice and delivers its native selection notification without a popup or focus change.</summary>
    internal static class OfficeNativeComboSelect
    {
        [DllImport("user32.dll")] private static extern IntPtr GetParent(IntPtr window);
        [DllImport("user32.dll")] private static extern int GetDlgCtrlID(IntPtr window);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr window);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
        [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr window);
        [DllImport("user32.dll")] private static extern bool IsChild(IntPtr parent, IntPtr child);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder text, int capacity);
        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")] private static extern IntPtr GetWindowLongPtr(IntPtr window, int index);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr SendMessageTimeout(IntPtr window, uint message, IntPtr wParam, IntPtr lParam, uint flags, uint timeout, out IntPtr result);
        [DllImport("user32.dll", EntryPoint = "SendMessageTimeoutW", CharSet = CharSet.Unicode)] private static extern IntPtr FindTextTimeout(IntPtr window, uint message, IntPtr wParam, string text, uint flags, uint timeout, out IntPtr result);

        /// <summary>Requires stable owned selector identity for reads and an enabled selector before each mutation.</summary>
        internal static void Select(IntPtr window, IntPtr owner, int pid, uint thread, string expected, Action requireOwner, Action<string> record)
        {
            if (string.IsNullOrEmpty(expected) || expected.Length > 2048 || requireOwner == null || record == null)
                throw new ArgumentException("An exact bounded choice and ownership/receipt callbacks are required.");
            var parent = GetParent(window);
            Action guard = () =>
            {
                requireOwner();
                uint actualPid, parentPid;
                uint actualThread = GetWindowThreadProcessId(window, out actualPid), parentThread = GetWindowThreadProcessId(parent, out parentPid);
                var cls = new StringBuilder(128); GetClassName(window, cls, cls.Capacity);
                if (!OfficeNativeComboRead.IdentityMatches(pid, thread, (int)actualPid, actualThread, cls.ToString(),
                    GetWindowLongPtr(window, -16).ToInt64(), IsWindow(window), IsChild(owner, window), IsWindowVisible(window))
                    || parent == IntPtr.Zero || GetParent(window) != parent || parentPid != pid || parentThread != thread)
                    throw new InvalidOperationException("Original native choice or notification parent changed.");
            };
            Action requireEnabled = () =>
            {
                guard();
                if (!IsWindowEnabled(window))
                    throw new InvalidOperationException("Original native choice became disabled before a selection action; no retry.");
            };
            long find(long after)
            {
                requireEnabled(); IntPtr result;
                if (FindTextTimeout(window, 0x0158, new IntPtr(after), expected, 3, 3000, out result) == IntPtr.Zero)
                    throw new TimeoutException("Native choice lookup did not return; no mutation was attempted.");
                return result.ToInt64();
            }
            long index = find(-1);
            if (index < 0 || index >= 4096 || find(index) != index)
                throw new InvalidOperationException("The exact native choice is absent, ambiguous or exceeds its bound.");
            long message(IntPtr target, uint id, IntPtr wParam, IntPtr lParam)
            {
                requireEnabled(); IntPtr result;
                if (SendMessageTimeout(target, id, wParam, lParam, 3, 3000, out result) == IntPtr.Zero)
                    throw new TimeoutException("Native selection action did not return; its outcome is uncertain and cannot be replayed.");
                return result.ToInt64();
            }
            SelectOnce(guard, index,
                () => message(window, 0x014E, new IntPtr(index), IntPtr.Zero),
                () => OfficeNativeComboRead.Read(window, owner, pid, thread, guard), expected,
                // CB_SETCURSEL does not send CBN_SELCHANGE; notify the actual parent once.
                // https://learn.microsoft.com/en-us/windows/win32/controls/cbn-selchange
                () => message(parent, 0x0111, new IntPtr((1 << 16) | (GetDlgCtrlID(window) & 0xffff)), window), record);
        }

        /// <summary>Sets once, verifies the exact displayed value, then notifies once; any failure stops the sequence.</summary>
        internal static void SelectOnce(Action guard, long expectedIndex, Func<long> setIndex, Func<string> readText,
            string expectedText, Action notify, Action<string> record)
        {
            if (guard == null || setIndex == null || readText == null || notify == null || record == null)
                throw new ArgumentNullException(nameof(guard));
            if (expectedIndex < 0 || expectedIndex >= 4096 || string.IsNullOrEmpty(expectedText)) throw new ArgumentException("Invalid exact choice.");
            guard(); record("NativeComboSetIntentOnce"); guard();
            if (setIndex() != expectedIndex) throw new InvalidOperationException("The native selection index was not accepted; no notification or retry.");
            guard();
            if (!string.Equals(readText(), expectedText, StringComparison.Ordinal)) throw new InvalidOperationException("The selected native label changed; no notification or retry.");
            guard(); record("NativeComboNotifyIntentOnce"); guard();
            notify(); guard();
            if (!string.Equals(readText(), expectedText, StringComparison.Ordinal)) throw new InvalidOperationException("The selection notification changed the native choice; no replay.");
            record("NativeComboSelectionReadback");
        }
    }
}
