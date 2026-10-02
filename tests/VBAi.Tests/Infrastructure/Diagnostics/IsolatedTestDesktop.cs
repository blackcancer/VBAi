using System;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace VBAi.Tests.Integration
{
    /// <summary>Starts owned test processes on an inactive Windows desktop without changing the input desktop.</summary>
    internal static class IsolatedTestDesktop
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
        [DllImport("user32.dll")] private static extern IntPtr GetProcessWindowStation();
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool GetUserObjectInformationW(IntPtr handle, int index, StringBuilder value, uint size, out uint required);
        [DllImport("kernel32.dll")] internal static extern uint GetCurrentThreadId();
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

        internal static string DesktopName(uint threadId) => ObjectName(GetThreadDesktop(threadId));

        internal static bool HasWindows(string name)
        {
            RequireName(name);
            IntPtr handle = OpenDesktopW(name, 0, false, 0x41);
            if (handle == IntPtr.Zero) throw new Win32Exception(Marshal.GetLastWin32Error());
            try
            {
                int count = 0;
                WindowVisitor visitor = (window, state) => ++count <= 8192;
                if (!EnumDesktopWindows(handle, visitor, IntPtr.Zero))
                    throw new InvalidOperationException("Private desktop window inventory is incomplete.");
                return count != 0;
            }
            finally { CloseDesktop(handle); }
        }

        internal static string InputDesktopName()
        {
            IntPtr input = OpenInputDesktop(0, false, 1);
            try { return ObjectName(input); }
            finally { if (input != IntPtr.Zero) CloseDesktop(input); }
        }

        internal static void RequireCurrent(string expected)
        {
            RequireName(expected);
            if (!string.Equals(DesktopName(GetCurrentThreadId()), expected, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(InputDesktopName(), expected, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Test execution requires its exact inactive desktop; there is no input-desktop fallback.");
        }

        internal static IDisposable Create(string name)
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
                // Office parses these literal switches from the raw command line. Quoted
                // "/x" is treated as a workbook name on the qualified Excel build.
                command.Append(' ').Append(argument == "/x" || argument == "/automation" ? argument : Quote(argument));
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
            var startup = new StartupInfo { Size = (uint)Marshal.SizeOf(typeof(StartupInfo)),
                Desktop = "WinSta0\\" + desktopName, Flags = 0x00000080 }; // STARTF_FORCEOFFFEEDBACK
            ProcessInformation child;
            if (!CreateProcessW(executable, new StringBuilder(CommandLine(executable, arguments)), IntPtr.Zero,
                IntPtr.Zero, false, 0x08000000, IntPtr.Zero, workingDir, ref startup, out child)) // CREATE_NO_WINDOW
                throw new Win32Exception(Marshal.GetLastWin32Error());
            return new NativeChild(child.Process, child.Thread, (int)child.ProcessId, child.ThreadId);
        }

        private sealed class DesktopLease : IDisposable
        {
            private IntPtr handle;
            internal DesktopLease(IntPtr handle) { this.handle = handle; }
            public void Dispose() { if (handle != IntPtr.Zero) { CloseDesktop(handle); handle = IntPtr.Zero; } }
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
