using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Text;

namespace VBAi.Tests.Integration
{
    /// <summary>Starts owned test processes on an inactive Windows desktop without changing the input desktop.</summary>
    internal static partial class IsolatedTestDesktop
    {
        private const uint DesktopAccess = 0x000000C7; // Read/create window/create menu/enumerate/write; no SwitchDesktop right.
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr CreateDesktopW(string name, IntPtr device, IntPtr mode, uint flags, uint access, IntPtr security);
        [DllImport("user32.dll", SetLastError = true)] private static extern bool CloseDesktop(IntPtr desktop);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr OpenDesktopW(string name, uint flags, bool inherit, uint access);
        private delegate bool WindowVisitor(IntPtr window, IntPtr state);
        [DllImport("user32.dll", SetLastError = true)] private static extern bool EnumDesktopWindows(IntPtr desktop, WindowVisitor visitor, IntPtr state);
        [DllImport("user32.dll", SetLastError = true)] private static extern IntPtr OpenInputDesktop(uint flags, bool inherit, uint access);
        [DllImport("user32.dll", SetLastError = true)] private static extern IntPtr GetThreadDesktop(uint threadId);
        [DllImport("user32.dll", SetLastError = true)] private static extern bool SetThreadDesktop(IntPtr desktop);
        [DllImport("user32.dll", SetLastError = true)] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr window, uint flags);
        [DllImport("user32.dll")] private static extern IntPtr GetProcessWindowStation();
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool GetUserObjectInformationW(IntPtr handle, int index, StringBuilder value, uint size, out uint required);
        [DllImport("kernel32.dll")] internal static extern uint GetCurrentThreadId();
        [DllImport("kernel32.dll")] private static extern void SetLastError(uint error);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool CreateProcessW(string application, StringBuilder command, IntPtr processSecurity,
            IntPtr threadSecurity, bool inherit, uint flags, IntPtr environment, string directory,
            ref StartupInfo startup, out ProcessInformation information);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern uint WaitForSingleObject(IntPtr handle, uint milliseconds);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern bool GetExitCodeProcess(IntPtr handle, out uint exitCode);
        [DllImport("kernel32.dll")] private static extern bool CloseHandle(IntPtr handle);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct StartupInfo
        {
            internal uint Size;
            internal string Reserved, Desktop, Title;
            internal uint X, Y, XSize, YSize, XChars, YChars, Fill, Flags;
            internal ushort Show, ReservedSize;
            internal IntPtr ReservedPointer, StandardInput, StandardOutput, StandardError;
        }
        [StructLayout(LayoutKind.Sequential)]
        private struct ProcessInformation
        {
            internal IntPtr Process, Thread;
            internal uint ProcessId, ThreadId;
        }

        internal static void RequireName(string name)
        {
            Guid value;
            const string prefix = "VBAiTests_";
            if (name == null || !name.StartsWith(prefix, StringComparison.Ordinal) ||
                !Guid.TryParseExact(name.Substring(prefix.Length), "N", out value))
                throw new ArgumentException("Only a generated VBAi test desktop name is permitted.", nameof(name));
        }

        private static string ObjectName(IntPtr handle)
        {
            if (handle == IntPtr.Zero) throw new Win32Exception(Marshal.GetLastWin32Error());
            var name = new StringBuilder(256); uint required;
            if (!GetUserObjectInformationW(handle, 2, name, (uint)(name.Capacity * 2), out required))
                throw new Win32Exception(Marshal.GetLastWin32Error());
            return name.ToString();
        }

        /// <summary>Retains separate native desktop/name failures without treating an unavailable thread as proof of placement.</summary>
        internal sealed class ThreadDesktopObservation
        {
            public uint ThreadId { get; }
            public long Handle { get; }
            public string Name { get; }
            public int DesktopError { get; }
            public int NameError { get; }
            internal ThreadDesktopObservation(uint threadId, long handle, string name, int desktopError, int nameError)
            { ThreadId = threadId; Handle = handle; Name = name; DesktopError = desktopError; NameError = nameError; }
        }

        internal static ThreadDesktopObservation ReadThreadDesktop(uint threadId)
        {
            SetLastError(0);
            IntPtr handle = GetThreadDesktop(threadId);
            int desktopError = Marshal.GetLastWin32Error();
            if (handle == IntPtr.Zero) return new ThreadDesktopObservation(threadId, 0, null, desktopError, 0);
            var name = new StringBuilder(256); uint required;
            SetLastError(0);
            bool named = GetUserObjectInformationW(handle, 2, name, (uint)(name.Capacity * 2), out required);
            int nameError = Marshal.GetLastWin32Error();
            return new ThreadDesktopObservation(threadId, handle.ToInt64(), named ? name.ToString() : null,
                desktopError, named ? 0 : nameError);
        }

        internal static string RequireThreadDesktopName(ThreadDesktopObservation observation)
        {
            if (observation == null) throw new ArgumentNullException(nameof(observation));
            if (observation.Handle == 0)
                throw new Win32Exception(observation.DesktopError, "GetThreadDesktop returned NULL for thread " +
                    observation.ThreadId + "; error=" + observation.DesktopError + ". Desktop identity is unproved.");
            if (observation.NameError != 0 || string.IsNullOrEmpty(observation.Name))
                throw new Win32Exception(observation.NameError, "The desktop name is unavailable for thread " +
                    observation.ThreadId + "; error=" + observation.NameError + ". Desktop identity is unproved.");
            return observation.Name;
        }

        internal static string DesktopName(uint threadId) => RequireThreadDesktopName(ReadThreadDesktop(threadId));

        internal sealed class WindowEnumeration
        {
            internal bool Completed;
            internal int Error;
        }

        /// <summary>Preserves visitor failures alongside a failed identity recheck; partial enumerations never become success.</summary>
        internal static void InventoryNamedWindows(string name, Func<string> readName,
            Func<Func<IntPtr, bool>, WindowEnumeration> enumerate, Func<IntPtr, bool> visit)
        {
            if (readName == null || enumerate == null || visit == null) throw new ArgumentNullException("Named desktop inventory dependencies");
            RequireName(name);
            if (!string.Equals(readName(), name, StringComparison.Ordinal))
                throw new InvalidOperationException("The named desktop inventory handle differs from the required desktop.");
            int count = 0; Exception failure = null;
            WindowEnumeration result = null;
            try
            {
                result = enumerate(window =>
                {
                    if (++count > 8192) return false;
                    try { return visit(window); }
                    catch (Exception error) { failure = error; return false; }
                });
            }
            catch (Exception error)
            {
                failure = failure == null ? error : new AggregateException("Visitor and desktop enumeration both failed.", failure, error);
            }
            try
            {
                if (!string.Equals(readName(), name, StringComparison.Ordinal))
                    throw new InvalidOperationException("The named desktop identity changed during inventory.");
            }
            catch (Exception error)
            {
                failure = failure == null ? error : new AggregateException("Original inventory failure and desktop identity recheck both failed.", failure, error);
            }
            if (failure != null) ExceptionDispatchInfo.Capture(failure).Throw();
            if (result == null) throw new InvalidOperationException("Named desktop enumeration returned no terminal result.");
            RequireWindowInventory(result.Completed, count, result.Error);
        }

        /// <summary>Visits a complete bounded inventory through an exact named inactive desktop handle.</summary>
        internal static void InventoryWindows(string name, Func<IntPtr, bool> visit)
        {
            if (visit == null) throw new ArgumentNullException(nameof(visit));
            RequireCurrent(name);
            IntPtr handle = OpenDesktopW(name, 0, false, 0x41);
            if (handle == IntPtr.Zero) throw new Win32Exception(Marshal.GetLastWin32Error());
            try
            {
                InventoryNamedWindows(name, () => ObjectName(handle), callback =>
                {
                    WindowVisitor visitor = (window, state) => callback(window);
                    SetLastError(0);
                    bool completed = EnumDesktopWindows(handle, visitor, IntPtr.Zero);
                    return new WindowEnumeration { Completed = completed, Error = Marshal.GetLastWin32Error() };
                }, visit);
                RequireCurrent(name);
            }
            finally { CloseDesktop(handle); }
        }

        internal static bool HasWindows(string name)
        {
            RequireName(name);
            IntPtr handle = OpenDesktopW(name, 0, false, 0x41);
            if (handle == IntPtr.Zero) throw new Win32Exception(Marshal.GetLastWin32Error());
            try
            {
                if (!string.Equals(ObjectName(handle), name, StringComparison.Ordinal))
                    throw new InvalidOperationException("The private desktop inventory handle has a different identity.");
                int count = 0;
                WindowVisitor visitor = (window, state) => ++count <= 8192;
                SetLastError(0);
                bool completed = EnumDesktopWindows(handle, visitor, IntPtr.Zero);
                int error = Marshal.GetLastWin32Error();
                if (!string.Equals(ObjectName(handle), name, StringComparison.Ordinal))
                    throw new InvalidOperationException("The private desktop inventory identity changed.");
                return RequireWindowInventory(completed, count, error);
            }
            finally { CloseDesktop(handle); }
        }

        internal static bool RequireWindowInventory(bool completed, int count, int error)
        {
            // On the observed Windows 11 host an empty, valid desktop returns false/error 0,
            // without calling the callback. Any partial inventory/error remains a refusal.
            if (count < 0 || count > 8192 || (!completed && (count != 0 || error != 0)))
                throw new InvalidOperationException("Private desktop window inventory is incomplete.");
            return count != 0;
        }

        /// <summary>Attaches a fresh, window-free sentinel thread only to the exact generated inactive desktop.</summary>
        internal static IDisposable BindPrivateSentinelThread(string name)
        {
            RequireName(name);
            if (string.Equals(InputDesktopName(), name, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("A sentinel may never attach to the input desktop.");
            IntPtr handle = OpenDesktopW(name, 0, false, DesktopAccess);
            if (handle == IntPtr.Zero) throw new Win32Exception(Marshal.GetLastWin32Error());
            if (!SetThreadDesktop(handle)) { int error = Marshal.GetLastWin32Error(); CloseDesktop(handle); throw new Win32Exception(error); }
            // The caller retains this handle until that exact thread's observed exit.
            return new DesktopLease(handle);
        }

        /// <summary>Requires successful inventory containing an exact live sentinel; failure is never interpreted as emptiness.</summary>
        internal static bool HasOtherWindows(string name, IntPtr sentinel, uint sentinelThread, uint sentinelProcess)
        {
            RequireName(name);
            uint actualProcess; uint actualThread = GetWindowThreadProcessId(sentinel, out actualProcess);
            if (sentinel == IntPtr.Zero || actualThread != sentinelThread || actualProcess != sentinelProcess ||
                !string.Equals(DesktopName(actualThread), name, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(InputDesktopName(), name, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The original private sentinel identity is unavailable; inventory is refused.");
            IntPtr handle = OpenDesktopW(name, 0, false, 0x41);
            if (handle == IntPtr.Zero) throw new Win32Exception(Marshal.GetLastWin32Error());
            try
            {
                int count = 0, other = 0; bool sentinelSeen = false;
                WindowVisitor visitor = (window, state) =>
                {
                    if (++count > 8192) return false;
                    if (window == sentinel) sentinelSeen = true;
                    else other++;
                    return true;
                };
                if (!EnumDesktopWindows(handle, visitor, IntPtr.Zero) || !sentinelSeen)
                    throw new InvalidOperationException("Private desktop inventory failed or omitted its live sentinel; ownership is retained.");
                return other != 0;
            }
            finally { CloseDesktop(handle); }
        }

        internal sealed class WindowIdentity
        {
            internal readonly IntPtr Window;
            internal readonly uint ProcessId, ThreadId;
            internal WindowIdentity(IntPtr window, uint process, uint thread) { Window = window; ProcessId = process; ThreadId = thread; }
        }

        internal sealed class WindowInventory
        {
            internal readonly string Desktop;
            internal readonly bool Complete;
            internal readonly WindowIdentity[] Windows;
            internal readonly bool? EnumerationSucceeded, IdentityComplete;
            internal readonly int? EnumerationError, IdentityError;
            internal readonly string CallbackStopReason;
            internal readonly IntPtr FailedWindow;
            internal readonly uint FailedProcessId, FailedThreadId;
            internal WindowInventory(string desktop, bool complete, params WindowIdentity[] windows)
            { Desktop = desktop; Complete = complete; Windows = windows; CallbackStopReason = "NotInstrumented"; }
            internal WindowInventory(string desktop, bool enumerationSucceeded, int enumerationError,
                bool identityComplete, string callbackStopReason, IntPtr failedWindow,
                uint failedProcessId, uint failedThreadId, int identityError, WindowIdentity[] windows)
            {
                Desktop = desktop; EnumerationSucceeded = enumerationSucceeded; EnumerationError = enumerationError;
                IdentityComplete = identityComplete; CallbackStopReason = callbackStopReason;
                FailedWindow = failedWindow; FailedProcessId = failedProcessId; FailedThreadId = failedThreadId;
                IdentityError = identityError; Windows = windows;
                Complete = enumerationSucceeded && identityComplete && callbackStopReason == "None";
            }
            internal string Diagnostic => string.Format(CultureInfo.InvariantCulture,
                "Desktop={0},Complete={1},EnumBOOL={2},EnumError={3},IdentityComplete={4},Count={5},Limit=8192,CallbackStop={6},FailedHWND=0x{7:X},FailedPID={8},FailedTID={9},IdentityError={10}",
                Desktop ?? "<null>", Complete, EnumerationSucceeded?.ToString() ?? "NotInstrumented",
                EnumerationError?.ToString(CultureInfo.InvariantCulture) ?? "NotInstrumented",
                IdentityComplete?.ToString() ?? "NotInstrumented", Windows?.Length ?? -1,
                CallbackStopReason ?? "<null>", FailedWindow.ToInt64(), FailedProcessId, FailedThreadId,
                IdentityError?.ToString(CultureInfo.InvariantCulture) ?? "NotInstrumented");
        }

        /// <summary>Checks foreign HWND membership using explicit desktop inventories, never foreign GetThreadDesktop.</summary>
        internal static void RequireOfficeWindowInventory(string desktop, uint ownedPid, bool requireWindow, IntPtr requiredWindow)
        {
            RequireCurrent(desktop);
            if (requiredWindow != IntPtr.Zero)
            {
                uint windowPid, rootPid;
                uint windowThread = GetWindowThreadProcessId(requiredWindow, out windowPid);
                IntPtr root = GetAncestor(requiredWindow, 2); // GA_ROOT: exact native child ancestry, no focus action.
                uint rootThread = GetWindowThreadProcessId(root, out rootPid);
                if (windowThread == 0 || rootThread == 0 || windowPid != ownedPid || rootPid != ownedPid)
                    throw new InvalidOperationException("The required native window or its top-level root differs from the exact owned Office process.");
                requiredWindow = root;
            }
            long sentinelWindow; uint sentinelPid, sentinelThread;
            if (!long.TryParse(Environment.GetEnvironmentVariable("VBAi_TEST_DESKTOP_SENTINEL_HWND"), NumberStyles.Integer, CultureInfo.InvariantCulture, out sentinelWindow) || sentinelWindow <= 0 ||
                !uint.TryParse(Environment.GetEnvironmentVariable("VBAi_TEST_DESKTOP_SENTINEL_PID"), NumberStyles.Integer, CultureInfo.InvariantCulture, out sentinelPid) || sentinelPid == 0 ||
                !uint.TryParse(Environment.GetEnvironmentVariable("VBAi_TEST_DESKTOP_SENTINEL_TID"), NumberStyles.Integer, CultureInfo.InvariantCulture, out sentinelThread) || sentinelThread == 0)
                throw new InvalidOperationException("The exact live private-desktop sentinel marker is required before Office discovery or mutation.");
            IntPtr privateHandle = OpenDesktopW(desktop, 0, false, 0x41);
            if (privateHandle == IntPtr.Zero) throw new Win32Exception(Marshal.GetLastWin32Error());
            WindowInventory privateInventory;
            try { privateInventory = ReadWindowInventory(privateHandle); }
            finally { CloseDesktop(privateHandle); }
            IntPtr inputHandle = OpenInputDesktop(0, false, 0x41);
            if (inputHandle == IntPtr.Zero) throw new Win32Exception(Marshal.GetLastWin32Error());
            WindowInventory inputInventory;
            try { inputInventory = ReadWindowInventory(inputHandle); }
            finally { CloseDesktop(inputHandle); }
            ValidateOfficeWindowInventories(desktop, ownedPid,
                new WindowIdentity(new IntPtr(sentinelWindow), sentinelPid, sentinelThread),
                privateInventory, inputInventory, requireWindow, requiredWindow);
            string inputAfter = InputDesktopName();
            if (!string.Equals(inputAfter, inputInventory.Desktop, StringComparison.OrdinalIgnoreCase))
                throw InventoryFailure("The user's input desktop changed during the read-only inventory; no native action or desktop restoration is performed. InputAfter=" + inputAfter,
                    desktop, ownedPid, privateInventory, inputInventory);
            RequireCurrent(desktop); // Never restore or override the user's input desktop choice.
        }

        private static WindowInventory ReadWindowInventory(IntPtr handle)
        {
            string name = ObjectName(handle);
            var windows = new List<WindowIdentity>(); bool identityComplete = true;
            string callbackStop = "None"; IntPtr failedWindow = IntPtr.Zero;
            uint failedProcess = 0, failedThread = 0; int identityError = 0;
            WindowVisitor visitor = (window, state) =>
            {
                uint process = 0;
                SetLastError(0);
                uint thread = GetWindowThreadProcessId(window, out process);
                int currentIdentityError = Marshal.GetLastWin32Error(); // Before any other native call.
                windows.Add(new WindowIdentity(window, process, thread));
                if (window == IntPtr.Zero || process == 0 || thread == 0)
                {
                    identityComplete = false; callbackStop = "InvalidWindowIdentity";
                    failedWindow = window; failedProcess = process; failedThread = thread;
                    identityError = currentIdentityError;
                    // EnumDesktopWindows requires a failing callback to set its error.
                    SetLastError((uint)(currentIdentityError != 0 ? currentIdentityError : 1400));
                    return false;
                }
                if (windows.Count > 8192)
                {
                    callbackStop = "WindowLimitExceeded"; failedWindow = window;
                    failedProcess = process; failedThread = thread;
                    SetLastError(234); // ERROR_MORE_DATA: our explicit inventory cap.
                    return false;
                }
                return true;
            };
            SetLastError(0);
            bool complete = EnumDesktopWindows(handle, visitor, IntPtr.Zero);
            int enumerationError = Marshal.GetLastWin32Error(); // Before CloseDesktop or logging.
            return new WindowInventory(name, complete, enumerationError, identityComplete, callbackStop,
                failedWindow, failedProcess, failedThread, identityError, windows.ToArray());
        }

        private static InvalidOperationException InventoryFailure(string reason, string expected, uint ownedPid,
            WindowInventory privateInventory, WindowInventory inputInventory)
        {
            return new InvalidOperationException(reason + " Expected=" + expected + ",OwnedPID=" + ownedPid +
                "; Private={" + (privateInventory?.Diagnostic ?? "<null>") + "}; Input={" +
                (inputInventory?.Diagnostic ?? "<null>") + "}. No native action or automatic inventory retry.");
        }

        /// <summary>Pure acceptance seam: failure, missing sentinel/target or owned input windows always refuse.</summary>
        internal static void ValidateOfficeWindowInventories(string expected, uint ownedPid, WindowIdentity sentinel,
            WindowInventory privateInventory, WindowInventory inputInventory, bool requireWindow, IntPtr requiredWindow)
        {
            RequireName(expected);
            if (ownedPid == 0 || sentinel == null || sentinel.Window == IntPtr.Zero || sentinel.ProcessId == 0 ||
                sentinel.ThreadId == 0 || sentinel.ProcessId == ownedPid || privateInventory == null || inputInventory == null ||
                !privateInventory.Complete || !inputInventory.Complete || privateInventory.Windows == null || inputInventory.Windows == null ||
                !string.Equals(privateInventory.Desktop, expected, StringComparison.OrdinalIgnoreCase) ||
                string.IsNullOrWhiteSpace(inputInventory.Desktop) || string.Equals(inputInventory.Desktop, expected, StringComparison.OrdinalIgnoreCase))
                throw InventoryFailure("Exact private/input desktop inventories are incomplete or unverified; Office actions are refused.", expected, ownedPid, privateInventory, inputInventory);
            bool foundSentinel = false, foundRequired = requiredWindow == IntPtr.Zero; int ownedWindows = 0;
            foreach (var window in privateInventory.Windows)
            {
                if (window == null || window.Window == IntPtr.Zero || window.ProcessId == 0 || window.ThreadId == 0)
                    throw InventoryFailure("Private native window identity is incomplete.", expected, ownedPid, privateInventory, inputInventory);
                if (window.Window == sentinel.Window && window.ProcessId == sentinel.ProcessId && window.ThreadId == sentinel.ThreadId) foundSentinel = true;
                if (window.ProcessId == ownedPid) { ownedWindows++; if (window.Window == requiredWindow) foundRequired = true; }
            }
            foreach (var window in inputInventory.Windows)
                if (window == null || window.Window == IntPtr.Zero || window.ProcessId == 0 || window.ThreadId == 0 || window.ProcessId == ownedPid)
                    throw InventoryFailure("An owned Office window is on the input desktop, or its inventory is incomplete; no native action is allowed.", expected, ownedPid, privateInventory, inputInventory);
            if (!foundSentinel || !foundRequired || (requireWindow && ownedWindows == 0))
                throw InventoryFailure("The exact live sentinel or required owned Office HWND was not observed on the private desktop.", expected, ownedPid, privateInventory, inputInventory);
        }

        internal static string InputDesktopName()
        {
            IntPtr input = OpenInputDesktop(0, false, 1);
            try { return ObjectName(input); }
            finally { if (input != IntPtr.Zero) CloseDesktop(input); }
        }

        internal static void RequireCurrent(string expected)
        {
            RequireObserved(expected, DesktopName(GetCurrentThreadId()), InputDesktopName());
        }

        internal static void RequireObserved(string expected, string actual, string input)
        {
            RequireName(expected);
            if (!string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase) || string.IsNullOrEmpty(input) ||
                string.Equals(input, expected, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Test execution requires its exact inactive desktop; there is no input-desktop fallback.");
        }

        internal static DesktopLease Create(string name)
        {
            RequireName(name);
            if (!string.Equals(ObjectName(GetProcessWindowStation()), "WinSta0", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("An interactive user window station is required without changing its identity.");
            if (string.Equals(InputDesktopName(), name, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The input desktop cannot be used for qualification.");
            IntPtr handle = CreateDesktopW(name, IntPtr.Zero, IntPtr.Zero, 0, DesktopAccess, IntPtr.Zero);
            if (handle == IntPtr.Zero) throw new Win32Exception(Marshal.GetLastWin32Error());
            return new DesktopLease(handle);
        }

        /// <summary>Quotes a single CreateProcess argument; this is not shell text.</summary>
        internal static string Quote(string argument)
        {
            if (argument == null || argument.IndexOf('\0') >= 0) throw new ArgumentException("Invalid process argument.");
            var result = new StringBuilder("\""); int slashes = 0;
            foreach (char value in argument)
            {
                if (value == '\\') { slashes++; continue; }
                result.Append('\\', value == '"' ? slashes * 2 + 1 : slashes);
                result.Append(value); slashes = 0;
            }
            result.Append('\\', slashes * 2); result.Append('"');
            return result.ToString();
        }

        internal static string CommandLine(string executable, string[] arguments)
        {
            if (arguments == null) throw new ArgumentNullException(nameof(arguments));
            var command = new StringBuilder(Quote(executable));
            foreach (string argument in arguments)
                // Office parses switches from the raw command line. Keep only these
                // exact reviewed switches literal; quote every document and other argument.
                command.Append(' ').Append(argument == "/x" || argument == "/automation" || argument == "/a" ? argument : Quote(argument));
            if (command.Length >= 32767) throw new ArgumentException("Process command exceeds the Windows bound.");
            return command.ToString();
        }

        /// <summary>Uses an explicit inactive desktop and retains the original process handle; disposing never terminates a process.</summary>
        internal static NativeChild Launch(string executable, string[] arguments, string workingDir, string desktopName)
        {
            RequireName(desktopName);
            if (!Path.IsPathRooted(executable) || !File.Exists(executable) ||
                !Path.IsPathRooted(workingDir) || !Directory.Exists(workingDir))
                throw new ArgumentException("An existing absolute executable and working directory are required.");
            if (string.Equals(InputDesktopName(), desktopName, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The qualification desktop became interactive; no process is started.");
            var startup = new StartupInfo
            {
                Size = (uint)Marshal.SizeOf(typeof(StartupInfo)),
                Desktop = "WinSta0\\" + desktopName,
                Flags = 0x00000080
            }; // STARTF_FORCEOFFFEEDBACK
            ProcessInformation child;
            if (!CreateProcessW(executable, new StringBuilder(CommandLine(executable, arguments)), IntPtr.Zero,
                IntPtr.Zero, false, 0x08000000, IntPtr.Zero, workingDir, ref startup, out child)) // CREATE_NO_WINDOW
                throw new Win32Exception(Marshal.GetLastWin32Error());
            return new NativeChild(child.Process, child.Thread, (int)child.ProcessId, child.ThreadId);
        }

        internal struct DesktopCloseResult
        {
            internal readonly bool Succeeded;
            internal readonly int Error;
            internal DesktopCloseResult(bool succeeded, int error) { Succeeded = succeeded; Error = error; }
        }

        /// <summary>Owns only a Create/OpenDesktop handle; a failed native close is recorded and never replayed.</summary>
        internal sealed class DesktopLease : IDisposable
        {
            private readonly object gate = new object();
            private readonly Func<IntPtr, DesktopCloseResult> close;
            private IntPtr handle;
            internal IntPtr Handle { get { lock (gate) return handle; } }
            internal bool CloseAttempted { get; private set; }
            internal bool CloseSucceeded { get; private set; }
            internal int CloseError { get; private set; }
            internal Exception CloseFailure { get; private set; }
            internal DesktopLease(IntPtr handle) : this(handle, CloseNative) { }
            internal DesktopLease(IntPtr handle, Func<IntPtr, DesktopCloseResult> close)
            {
                if (handle == IntPtr.Zero) throw new ArgumentException("An owned desktop handle is required.", nameof(handle));
                this.handle = handle;
                this.close = close ?? throw new ArgumentNullException(nameof(close));
            }
            private static DesktopCloseResult CloseNative(IntPtr handle)
            {
                bool succeeded = CloseDesktop(handle);
                int error = succeeded ? 0 : Marshal.GetLastWin32Error();
                return new DesktopCloseResult(succeeded, error);
            }
            public void Dispose()
            {
                lock (gate)
                {
                    if (CloseAttempted)
                    {
                        if (!CloseSucceeded) throw CloseFailure;
                        return;
                    }
                    CloseAttempted = true; // Claim before the single native call, including exceptional delivery.
                    try
                    {
                        DesktopCloseResult result = close(handle);
                        CloseError = result.Error;
                        if (!result.Succeeded) throw new Win32Exception(result.Error, "The single CloseDesktop attempt failed; its owned handle is retained without retry.");
                        CloseSucceeded = true;
                        handle = IntPtr.Zero;
                    }
                    catch (Exception error)
                    {
                        CloseFailure = error;
                        var nativeError = error as Win32Exception;
                        if (nativeError != null) CloseError = nativeError.NativeErrorCode;
                        throw;
                    }
                }
            }
        }

        /// <summary>Terminal publication follows verified closes; the original child handle stays held until then.</summary>
        internal static void CompleteOwnedShutdown(Action closeSentinel, DesktopLease desktop, Action releaseChild, Action publishTerminal)
        {
            if (closeSentinel == null || desktop == null || releaseChild == null || publishTerminal == null)
                throw new ArgumentNullException("Owned shutdown requires its complete close/release/publication contract.");
            closeSentinel();
            desktop.Dispose();
            releaseChild();
            publishTerminal();
        }

        /// <summary>An early refusal may close an unattempted lease once only after no child or sentinel remains.</summary>
        internal static bool PrepareRefusal(DesktopLease desktop, bool retainedOwners, out Exception closeFailure)
        {
            closeFailure = null;
            if (!retainedOwners && desktop != null && !desktop.CloseAttempted)
            {
                try { desktop.Dispose(); }
                catch (Exception error) { closeFailure = error; }
            }
            if (desktop != null && !desktop.CloseSucceeded && closeFailure == null)
                closeFailure = desktop.CloseFailure;
            return retainedOwners || (desktop != null && !desktop.CloseSucceeded);
        }

        internal sealed class NativeChild : IDisposable
        {
            internal IntPtr ProcessHandle { get; private set; }
            internal int ProcessId { get; private set; }
            internal uint ThreadId { get; private set; }
            private IntPtr thread;
            internal NativeChild(IntPtr process, IntPtr thread, int pid, uint tid)
            { ProcessHandle = process; this.thread = thread; ProcessId = pid; ThreadId = tid; }
            internal bool Wait(int milliseconds)
            {
                if (milliseconds < 0 || ProcessHandle == IntPtr.Zero) throw new ArgumentOutOfRangeException(nameof(milliseconds));
                uint state = WaitForSingleObject(ProcessHandle, (uint)milliseconds);
                if (state == 0xFFFFFFFF) throw new Win32Exception(Marshal.GetLastWin32Error());
                if (state != 0 && state != 258) throw new InvalidOperationException("Unexpected original-process wait outcome.");
                return state == 0;
            }
            internal uint ExitCode()
            {
                if (!Wait(0)) throw new InvalidOperationException("A returned launcher is not an observed original exit.");
                uint code;
                if (!GetExitCodeProcess(ProcessHandle, out code)) throw new Win32Exception(Marshal.GetLastWin32Error());
                return code;
            }
            public void Dispose()
            {
                if (thread != IntPtr.Zero) { CloseHandle(thread); thread = IntPtr.Zero; }
                if (ProcessHandle != IntPtr.Zero) { CloseHandle(ProcessHandle); ProcessHandle = IntPtr.Zero; }
            }
        }
    }
}
