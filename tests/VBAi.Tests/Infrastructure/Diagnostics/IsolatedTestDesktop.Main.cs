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
                (image, args, directory) =>
                {
                    var startup = new StartupInfo
                    {
                        Size = (uint)Marshal.SizeOf(typeof(StartupInfo)),
                        Desktop = "WinSta0\\Default",
                        Flags = 0x00000080
                    };
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
            public bool WindowDestroyedDuringRead { get; set; }
            public int Visited { get; set; }
            public MainWindow[] Windows { get; set; }
            public int ReadAttempts { get; set; }
            public MainEnumerationFailure[] RejectedObservations { get; set; }
        }

        /// <summary>Retains an interrupted read without treating its partial rows as placement proof.</summary>
        internal sealed class MainEnumerationFailure
        {
            public int EnumerationError { get; set; }
            public bool WindowDestroyedDuringRead { get; set; }
            public int Visited { get; set; }
            public int OwnedWindows { get; set; }
        }

        /// <summary>Enumerates the caller's independently checked Default desktop; no foreign thread-desktop query.</summary>
        internal static MainInventory ReadMainWindows(uint ownedPid) =>
            ReadCompleteMainInventory(ownedPid, () => ReadMainWindowsOnce(ownedPid), RequireMainCurrent);

        /// <summary>Allows one fresh read after an interrupted enumeration or a destroyed window; never replays an action.</summary>
        internal static MainInventory ReadCompleteMainInventory(uint ownedPid, Func<MainInventory> read, Action requireMain)
        {
            if (read == null || requireMain == null) throw new ArgumentNullException("Main inventory read dependencies");
            if (ownedPid == 0) throw new ArgumentOutOfRangeException(nameof(ownedPid));
            var rejected = new List<MainEnumerationFailure>();
            int attempt = 1;
            while (true)
            {
                requireMain();
                var inventory = read();
                requireMain();
                RequireMainInventoryShape(inventory, ownedPid);
                if (!inventory.Complete && attempt == 1 &&
                    (inventory.EnumerationError == 0 || inventory.WindowDestroyedDuringRead))
                {
                    rejected.Add(new MainEnumerationFailure { EnumerationError = inventory.EnumerationError,
                        WindowDestroyedDuringRead = inventory.WindowDestroyedDuringRead,
                        Visited = inventory.Visited, OwnedWindows = inventory.Windows.Length });
                    attempt++;
                    continue;
                }
                RequireMainInventory(inventory, ownedPid, IntPtr.Zero, false, false);
                inventory.ReadAttempts = attempt;
                inventory.RejectedObservations = rejected.ToArray();
                return inventory;
            }
        }

        /// <summary>Performs one bounded native enumeration and retains its completion/error status for admission.</summary>
        private static MainInventory ReadMainWindowsOnce(uint ownedPid)
            => ReadMainWindowsOnce(ownedPid, visitor =>
                {
                    SetLastError(0);
                    return EnumWindows((window, state) => visitor(window), IntPtr.Zero);
                }, window => ReadMainWindow(ownedPid, window, ReadMainWindowIdentity,
                    ReadMainWindowClassName, IsWindowVisible), Marshal.GetLastWin32Error, RequireMainCurrent);

        /// <summary>Preserves a destroyed-window read as rejected evidence; other callback failures keep their original exception.</summary>
        internal static MainInventory ReadMainWindowsOnce(uint ownedPid, Func<Func<IntPtr, bool>, bool> enumerate,
            Func<IntPtr, MainWindow> readWindow, Func<int> lastError, Action requireMain)
        {
            requireMain();
            var rows = new List<MainWindow>(); int visited = 0; Exception failure = null;
            bool destroyed = false;
            Func<IntPtr, bool> visitor = window =>
            {
                try
                {
                    if (++visited > 8192) throw new InvalidOperationException("Main window inventory exceeded its bound.");
                    var row = readWindow(window);
                    if (row != null) rows.Add(row);
                    return true;
                }
                catch (Win32Exception error) when (error.NativeErrorCode == 1400)
                { destroyed = true; return false; }
                catch (Exception error) { failure = error; return false; }
            };
            bool complete = enumerate(visitor);
            int nativeError = destroyed ? 1400 : lastError();
            if (failure != null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
            requireMain();
            var result = new MainInventory
            {
                Desktop = "Default",
                Complete = complete && !destroyed,
                EnumerationError = nativeError,
                WindowDestroyedDuringRead = destroyed,
                Visited = visited,
                Windows = rows.ToArray()
            };
            return result;
        }

        internal static void RequireMainInventory(MainInventory inventory, uint ownedPid, IntPtr requiredRoot,
            bool requireWord, bool requireVbe)
        {
            RequireMainInventoryShape(inventory, ownedPid);
            if (!inventory.Complete)
                throw new InvalidOperationException("The complete Default owned-window inventory is required. " +
                    "Complete=false; Win32=" + inventory.EnumerationError + "; Visited=" + inventory.Visited +
                    "; OwnedWindows=" + inventory.Windows.Length + ". No action is authorized by a partial read.");
            if (requiredRoot != IntPtr.Zero && inventory.Windows.Count(row => row.Handle == requiredRoot.ToInt64()) != 1)
                throw new InvalidOperationException("The exact owned root is absent from Default.");
            if (requireWord && !inventory.Windows.Any(row => row.ClassName == "OpusApp" && row.Visible))
                throw new InvalidOperationException("The owned Word root is not visible on Default.");
            if (requireVbe && inventory.Windows.Count(row => row.ClassName == "wndclass_desked_gsk" && row.Visible) != 1)
                throw new InvalidOperationException("The unique owned VBE is not visible on Default.");
        }

        /// <summary>Rejects malformed, foreign, duplicate or oversized rows before any fresh enumeration.</summary>
        private static void RequireMainInventoryShape(MainInventory inventory, uint ownedPid)
        {
            if (ownedPid == 0 || inventory == null || inventory.Desktop != "Default" ||
                (inventory.WindowDestroyedDuringRead && (inventory.Complete || inventory.EnumerationError != 1400)) ||
                inventory.Visited < 0 || inventory.Visited > 8192 || inventory.Windows == null ||
                inventory.Windows.Length > inventory.Visited || inventory.Windows.Any(row => row == null || row.Handle == 0 ||
                    row.ProcessId != ownedPid || row.ThreadId == 0 || string.IsNullOrEmpty(row.ClassName)) ||
                inventory.Windows.Select(row => row.Handle).Distinct().Count() != inventory.Windows.Length)
                throw new InvalidOperationException("The Default owned-window inventory has invalid, foreign, duplicate or oversized rows.");
        }
    }
}
