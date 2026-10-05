using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace VBAi.Tests.Integration
{
    /// <summary>Explicit interactive qualification; the inactive-desktop APIs keep their original contract.</summary>
    internal static partial class IsolatedTestDesktop
    {
        [DllImport("user32.dll", SetLastError = true)] private static extern bool EnumWindows(WindowVisitor visitor, IntPtr state);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern int GetClassNameW(IntPtr window, StringBuilder name, int capacity);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);

        internal static void RequireMainObserved(string station, string current, string input)
        {
            if (!string.Equals(station, "WinSta0", StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(current, "Default", StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(input, "Default", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Main qualification requires the current WinSta0\\Default and input Default; no switch or fallback is permitted.");
        }

        internal static void RequireMainCurrent() => RequireMainObserved(ObjectName(GetProcessWindowStation()),
            DesktopName(GetCurrentThreadId()), InputDesktopName());

        /// <summary>Checks placement immediately before the single launch; ownership transfers before any later observation.</summary>
        internal static T LaunchMainOnce<T>(string executable, string[] arguments, string workingDir,
            Action requireCurrent, Func<string, bool> fileExists, Func<string, bool> directoryExists,
            Func<string, string[], string, T> create)
        {
            if (requireCurrent == null || fileExists == null || directoryExists == null || create == null)
                throw new ArgumentNullException("Main launch dependencies");
            if (string.IsNullOrWhiteSpace(executable) || !Path.IsPathRooted(executable) || !fileExists(executable) ||
                string.IsNullOrWhiteSpace(workingDir) || !Path.IsPathRooted(workingDir) || !directoryExists(workingDir))
                throw new ArgumentException("Existing absolute executable and working directory required.");
            requireCurrent();
            return create(executable, arguments, workingDir);
        }

        internal static NativeChild LaunchMain(string executable, string[] arguments, string workingDir) =>
            LaunchMainOnce(executable, arguments, workingDir, RequireMainCurrent, File.Exists, Directory.Exists,
                (image, args, directory) => {
                    var startup = new StartupInfo { Size = (uint)Marshal.SizeOf(typeof(StartupInfo)),
                        Desktop = "WinSta0\\Default", Flags = 0x00000080 };
                    ProcessInformation child;
                    // Null environment inherits the fixture's scoped manifest. No shell/COM activation is involved.
                    if (!CreateProcessW(image, new StringBuilder(CommandLine(image, args)), IntPtr.Zero,
                        IntPtr.Zero, false, 0x08000000, IntPtr.Zero, directory, ref startup, out child))
                        throw new Win32Exception(Marshal.GetLastWin32Error());
                    return new NativeChild(child.Process, child.Thread, (int)child.ProcessId, child.ThreadId);
                });

        internal sealed class MainWindow
        {
            public long Handle { get; set; }
            public uint ProcessId { get; set; }
            public uint ThreadId { get; set; }
            public string ClassName { get; set; }
            public bool Visible { get; set; }
        }

        internal sealed class MainInventory
        {
            public string Desktop { get; set; }
            public bool Complete { get; set; }
            public int EnumerationError { get; set; }
            public int Visited { get; set; }
            public MainWindow[] Windows { get; set; }
        }

        /// <summary>Enumerates the caller's independently checked Default desktop; no foreign thread-desktop query.</summary>
        internal static MainInventory ReadMainWindows(uint ownedPid)
        {
            RequireMainCurrent();
            var rows = new List<MainWindow>(); int visited = 0; Exception failure = null;
            WindowVisitor visitor = (window, state) => {
                try
                {
                    if (++visited > 8192) throw new InvalidOperationException("Main window inventory exceeded its bound.");
                    uint pid; uint tid = GetWindowThreadProcessId(window, out pid);
                    if (window == IntPtr.Zero || pid == 0 || tid == 0) throw new InvalidOperationException("Incomplete main window identity.");
                    if (pid == ownedPid)
                    {
                        var name = new StringBuilder(256);
                        if (GetClassNameW(window, name, name.Capacity) == 0) throw new Win32Exception(Marshal.GetLastWin32Error());
                        rows.Add(new MainWindow { Handle = window.ToInt64(), ProcessId = pid, ThreadId = tid,
                            ClassName = name.ToString(), Visible = IsWindowVisible(window) });
                    }
                    return true;
                }
                catch (Exception error) { failure = error; return false; }
            };
            SetLastError(0);
            bool complete = EnumWindows(visitor, IntPtr.Zero);
            int nativeError = Marshal.GetLastWin32Error();
            if (failure != null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
            RequireMainCurrent();
            var result = new MainInventory { Desktop = "Default", Complete = complete, EnumerationError = nativeError,
                Visited = visited, Windows = rows.ToArray() };
            RequireMainInventory(result, ownedPid, IntPtr.Zero, false, false);
            return result;
        }

        internal static void RequireMainInventory(MainInventory inventory, uint ownedPid, IntPtr requiredRoot,
            bool requireWord, bool requireVbe)
        {
            if (ownedPid == 0 || inventory == null || !inventory.Complete || inventory.Desktop != "Default" ||
                inventory.Visited < 0 || inventory.Visited > 8192 || inventory.Windows == null ||
                inventory.Windows.Length > inventory.Visited || inventory.Windows.Any(row => row == null || row.Handle == 0 ||
                    row.ProcessId != ownedPid || row.ThreadId == 0 || string.IsNullOrEmpty(row.ClassName)) ||
                inventory.Windows.Select(row => row.Handle).Distinct().Count() != inventory.Windows.Length)
                throw new InvalidOperationException("The complete Default owned-window inventory is required.");
            if (requiredRoot != IntPtr.Zero && inventory.Windows.Count(row => row.Handle == requiredRoot.ToInt64()) != 1)
                throw new InvalidOperationException("The exact owned root is absent from Default.");
            if (requireWord && !inventory.Windows.Any(row => row.ClassName == "OpusApp" && row.Visible))
                throw new InvalidOperationException("The owned Word root is not visible on Default.");
            if (requireVbe && inventory.Windows.Count(row => row.ClassName == "wndclass_desked_gsk" && row.Visible) != 1)
                throw new InvalidOperationException("The unique owned VBE is not visible on Default.");
        }
    }
}
