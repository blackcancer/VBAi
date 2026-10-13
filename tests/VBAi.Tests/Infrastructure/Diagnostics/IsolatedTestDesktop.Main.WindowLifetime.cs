using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;

namespace VBAi.Tests.Integration
{
    internal static partial class IsolatedTestDesktop
    {
        /// <summary>Captures the process and thread owning a native window at one instant.</summary>
        internal sealed class MainWindowIdentity
        {
            public uint ProcessId { get; set; }
            public uint ThreadId { get; set; }
        }

        /// <summary>Rejects a destroyed HWND with its native error instead of inventing an owner.</summary>
        private static MainWindowIdentity ReadMainWindowIdentity(IntPtr window)
        {
            SetLastError(0);
            uint pid;
            uint tid = GetWindowThreadProcessId(window, out pid);
            if (tid == 0 && Marshal.GetLastWin32Error() == 1400) throw new Win32Exception(1400);
            return new MainWindowIdentity { ProcessId = pid, ThreadId = tid };
        }

        /// <summary>Reads the class without inheriting an earlier Win32 error.</summary>
        private static string ReadMainWindowClassName(IntPtr window)
        {
            SetLastError(0);
            var name = new StringBuilder(256);
            if (GetClassNameW(window, name, name.Capacity) == 0) throw new Win32Exception(Marshal.GetLastWin32Error());
            return name.ToString();
        }

        /// <summary>Admits an owned row only when its process and thread still match after class and visibility reads.</summary>
        /// <remarks>A destroyed window interrupts the whole inventory. A reused HWND never supplies placement proof.</remarks>
        internal static MainWindow ReadMainWindow(uint ownedPid, IntPtr window,
            Func<IntPtr, MainWindowIdentity> readIdentity, Func<IntPtr, string> readClass,
            Func<IntPtr, bool> readVisible)
        {
            if (ownedPid == 0 || window == IntPtr.Zero) throw new InvalidOperationException("Incomplete main window identity.");
            var before = readIdentity(window);
            if (before == null || before.ProcessId == 0 || before.ThreadId == 0)
                throw new InvalidOperationException("Incomplete main window identity.");
            if (before.ProcessId != ownedPid) return null;
            string name = readClass(window);
            bool visible = readVisible(window);
            var after = readIdentity(window);
            if (after == null || after.ProcessId != before.ProcessId || after.ThreadId != before.ThreadId)
                throw new InvalidOperationException("Main window ownership changed during the inventory.");
            if (string.IsNullOrEmpty(name)) throw new InvalidOperationException("Incomplete main window class.");
            return new MainWindow { Handle = window.ToInt64(), ProcessId = before.ProcessId,
                ThreadId = before.ThreadId, ClassName = name, Visible = visible };
        }
    }
}
