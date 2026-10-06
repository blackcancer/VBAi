using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace VBAi
{

    /// <summary>Activates and verifies an exact visible code window owned by the current process and UI thread.</summary>
    internal static class VbaNativeTestWindowFocus
    {

        /// <summary>Snapshot of native window identity and eligibility state used during guarded focus.</summary>
        internal sealed class Window
        {

            /// <summary>Window handle and parent handle captured during one observation.</summary>
            internal IntPtr Handle, Parent;

            /// <summary>Owning process and UI-thread IDs returned by the operating system.</summary>
            internal uint Process, Thread;

            /// <summary>Native window class name used to distinguish the MDI client and VBE code window.</summary>
            internal string Class;

            /// <summary>Observed existence, visibility, and enabled state required for focus operations.</summary>
            internal bool Exists, Visible, Enabled;
        }

        /// <summary>Injectable native-window operations required to locate, activate, and verify focus.</summary>
        internal interface IWindows
        {

            /// <summary>Gets the process ID of the current process.</summary>
            /// <value>Current process ID used to reject foreign windows.</value>
            uint CurrentProcess { get; }

            /// <summary>Gets the current UI thread ID.</summary>
            /// <value>Thread ID used to reject windows owned by another thread.</value>
            uint CurrentThread { get; }

            /// <summary>Captures identity, parent, class, and eligibility for one HWND.</summary>
            /// <param name="handle">Window handle to inspect.</param>
            /// <returns>Window snapshot; implementations represent invalid handles with <see cref="Window.Exists"/> false.</returns>
            Window Read(IntPtr handle);

            /// <summary>Enumerates direct child HWNDs of the supplied parent.</summary>
            /// <param name="parent">Parent whose immediate children are enumerated.</param>
            /// <returns>Direct child handles in native enumeration order.</returns>
            IntPtr[] Children(IntPtr parent);

            /// <summary>Reads the current window caption after ownership and class checks pass.</summary>
            /// <param name="handle">Eligible HWND whose title is required for exact code-window matching.</param>
            /// <returns>Current caption text.</returns>
            string Caption(IntPtr handle);

            /// <summary>Activates one verified MDI child through its verified MDI client.</summary>
            /// <param name="mdi">Owned MDI client handle.</param>
            /// <param name="child">Exact code-window child handle to activate.</param>
            void Activate(IntPtr mdi, IntPtr child);

            /// <summary>Requests keyboard focus for the verified native code-window handle.</summary>
            /// <param name="child">Exact code-window handle already activated.</param>
            void Focus(IntPtr child);

            /// <summary>Gets the current keyboard-focus HWND.</summary>
            /// <value>Focused handle, or zero when no window owns focus.</value>
            IntPtr Focused { get; }
        }

        /// <summary>Exact main, MDI-client, and code-window snapshots selected for the operation.</summary>
        internal sealed class Target
        {

            /// <summary>Verified ancestry chain from the VBE main window to its code child.</summary>
            internal Window Main, Mdi, Code;

            /// <summary>Exact code-window caption used to detect replacement or ambiguity.</summary>
            internal string Caption;
        }

        /// <summary>Finds, activates, and verifies exactly one owned VBE code window with the requested caption.</summary>
        /// <param name="windows">Native window access implementation.</param>
        /// <param name="main">Expected VBE main-window handle.</param>
        /// <param name="caption">Exact code-window caption; blank or over 1,024 characters is refused.</param>
        /// <returns>Verified target snapshots after activation and keyboard-focus checks.</returns>
        /// <exception cref="InvalidOperationException">The owned MDI client or exact code window is absent, ambiguous, or changes during activation.</exception>
        internal static Target Focus(IWindows windows, IntPtr main, string caption)
        {
            if (windows == null) throw new ArgumentNullException(nameof(windows));
            if (string.IsNullOrWhiteSpace(caption) || caption.Length > 1024)
                throw Refusal("The exact code-window caption is unavailable.");
            var root = windows.Read(main);
            RequireOwned(windows, root, "beforeDiscovery/Main");
            var clients = new List<Window>();
            foreach (var handle in windows.Children(main))
            {
                var candidate = windows.Read(handle);
                if (Eligible(windows, candidate) && candidate.Parent == main && candidate.Class == "MDIClient") clients.Add(candidate);
            }
            if (clients.Count != 1) throw Refusal("The owned direct MDI client is unavailable or ambiguous.");
            var mdi = clients[0];
            var matches = new List<Window>();
            foreach (var handle in windows.Children(mdi.Handle))
            {
                var candidate = windows.Read(handle);
                // Never read a title until the native owner, thread, class and parent match.
                if (Eligible(windows, candidate) && candidate.Parent == mdi.Handle && candidate.Class == "VbaWindow"
                    && string.Equals(windows.Caption(handle), caption, StringComparison.Ordinal)) matches.Add(candidate);
            }
            if (matches.Count != 1) throw Refusal("The exact owned native code window is unavailable or ambiguous.");
            var target = new Target { Main = root, Mdi = mdi, Code = matches[0], Caption = caption };
            VerifyTarget(windows, target, "beforeActivate");
            windows.Activate(mdi.Handle, target.Code.Handle);
            // Activation calls native window procedures synchronously and can change state.
            VerifyTarget(windows, target, "afterActivate");
            windows.Focus(target.Code.Handle);
            VerifyFocus(windows, target, "afterFocus");
            return target;
        }

        /// <summary>Revalidates the target chain and requires focus on the code window or one of its owned descendants.</summary>
        /// <param name="windows">Native window access implementation used for fresh observations.</param>
        /// <param name="target">Previously selected target snapshots.</param>
        /// <param name="stage">Diagnostic stage label included in a refusal report.</param>
        internal static void VerifyFocus(IWindows windows, Target target, string stage = "revalidateFocus")
        {
            VerifyTarget(windows, target, stage);
            var focused = windows.Focused;
            for (int depth = 0; depth < 32 && focused != IntPtr.Zero; depth++)
            {
                var current = windows.Read(focused);
                RequireOwned(windows, current, stage + "/FocusedAncestor" + depth);
                if (focused == target.Code.Handle) return;
                if (current.Parent == focused) break;
                focused = current.Parent;
            }
            throw Refusal("The exact native code window or its owned descendant does not have keyboard focus.");
        }

        /// <summary>Rechecks HWND identities, ancestry, caption uniqueness, and ownership before or after native calls.</summary>
        /// <param name="windows">Native window access implementation.</param>
        /// <param name="target">Selected VBE main, MDI client, and code window.</param>
        /// <param name="stage">Diagnostic stage label appended to refusal evidence.</param>
        private static void VerifyTarget(IWindows windows, Target target, string stage)
        {
            VerifyUnchanged(windows, target.Main, stage + "/Main");
            VerifyUnchanged(windows, target.Mdi, stage + "/Mdi");
            VerifyUnchanged(windows, target.Code, stage + "/Code");
            if (target.Mdi.Parent != target.Main.Handle || target.Code.Parent != target.Mdi.Handle
                || target.Mdi.Class != "MDIClient" || target.Code.Class != "VbaWindow"
                || !string.Equals(windows.Caption(target.Code.Handle), target.Caption, StringComparison.Ordinal))
                throw Refusal("The exact owned native code-window mapping changed.");
            int clients = 0, codes = 0;
            foreach (var handle in windows.Children(target.Main.Handle))
            {
                var candidate = windows.Read(handle);
                if (Eligible(windows, candidate) && candidate.Parent == target.Main.Handle && candidate.Class == "MDIClient") clients++;
            }
            foreach (var handle in windows.Children(target.Mdi.Handle))
            {
                var candidate = windows.Read(handle);
                if (Eligible(windows, candidate) && candidate.Parent == target.Mdi.Handle && candidate.Class == "VbaWindow"
                    && string.Equals(windows.Caption(handle), target.Caption, StringComparison.Ordinal)) codes++;
            }
            if (clients != 1 || codes != 1) throw Refusal("The exact owned native code-window mapping became ambiguous.");
            // Same-process title retrieval invokes the native window procedure.
            // Recheck handles after that potentially reentrant inspection.
            VerifyUnchanged(windows, target.Main, stage + "/MainAfterInspection");
            VerifyUnchanged(windows, target.Mdi, stage + "/MdiAfterInspection");
            VerifyUnchanged(windows, target.Code, stage + "/CodeAfterInspection");
        }

        /// <summary>Requires a fresh HWND snapshot to match the expected parent, class, process, and thread identity.</summary>
        /// <param name="windows">Native window access implementation.</param>
        /// <param name="expected">Previously accepted snapshot to compare.</param>
        /// <param name="role">Diagnostic role label such as Main, Mdi, or Code.</param>
        private static void VerifyUnchanged(IWindows windows, Window expected, string role)
        {
            var current = windows.Read(expected.Handle);
            RequireOwned(windows, current, role);
            if (current.Handle != expected.Handle || current.Parent != expected.Parent || current.Class != expected.Class
                || current.Process != expected.Process || current.Thread != expected.Thread)
                throw Refusal("The owned native window identity changed before focus. "
                    + DescribeWindow(windows, current, role) + "; expected={" + DescribeWindow(windows, expected, role) + "}");
        }

        /// <summary>Checks that a snapshot is visible, enabled, valid, and owned by the current process and thread.</summary>
        /// <param name="windows">Source of the current process and thread IDs.</param>
        /// <param name="window">Snapshot to test; null is ineligible.</param>
        /// <returns>True only when all native ownership and visibility checks pass.</returns>
        private static bool Eligible(IWindows windows, Window window) => window != null && window.Exists
            && window.Handle != IntPtr.Zero && window.Visible && window.Enabled
            && window.Process != 0 && window.Process == windows.CurrentProcess
            && window.Thread != 0 && window.Thread == windows.CurrentThread;

        /// <summary>Rejects any window snapshot that fails current-process, current-thread, visibility, or enabled checks.</summary>
        /// <param name="windows">Source of current process and thread IDs.</param>
        /// <param name="window">Observed snapshot to validate.</param>
        /// <param name="role">Diagnostic role label included in refusal details.</param>
        private static void RequireOwned(IWindows windows, Window window, string role)
        {
            if (!Eligible(windows, window)) throw Refusal("The native window is not visible and enabled on the owning process and UI thread. "
                + DescribeWindow(windows, window, role));
        }

        // Diagnostic data comes from the already observed native snapshot. Never
        // read captions, focus another window or enumerate more windows on refusal.
        /// <summary>Formats only previously captured scalar window values; it performs no further native inspection.</summary>
        /// <param name="windows">Source of expected current process and thread IDs.</param>
        /// <param name="window">Observed snapshot, or null when no HWND could be read.</param>
        /// <param name="role">Diagnostic role label.</param>
        /// <returns>Bounded scalar ownership evidence and failed eligibility fields.</returns>
        private static string DescribeWindow(IWindows windows, Window window, string role)
        {
            uint process = windows.CurrentProcess, thread = windows.CurrentThread;
            var failed = new List<string>();
            if (window == null) return "native={role=" + role + ",window=null,expectedPID=" + process + ",expectedThread=" + thread + ",failed=Null}";
            if (!window.Exists) failed.Add("Exists");
            if (window.Handle == IntPtr.Zero) failed.Add("HWnd");
            if (!window.Visible) failed.Add("Visible");
            if (!window.Enabled) failed.Add("Enabled");
            if (window.Process == 0 || window.Process != process) failed.Add("Process");
            if (window.Thread == 0 || window.Thread != thread) failed.Add("Thread");
            var kind = new StringBuilder();
            string name = window.Class ?? "";
            for (int index = 0; index < name.Length && index < 96; index++)
                kind.Append(char.IsControl(name[index]) ? '?' : name[index]);
            return "native={role=" + role + ",HWnd=" + window.Handle.ToInt64() + ",Class=" + kind.ToString()
                + ",Parent=" + window.Parent.ToInt64() + ",PID=" + window.Process + ",Thread=" + window.Thread
                + ",Exists=" + window.Exists + ",Visible=" + window.Visible + ",Enabled=" + window.Enabled
                + ",expectedPID=" + process + ",expectedThread=" + thread + ",failed=" + string.Join("|", failed) + "}";
        }

        /// <summary>Creates the fail-closed exception used when the exact native target cannot be established.</summary>
        /// <param name="message">Refusal reason, optionally including captured scalar diagnostic evidence.</param>
        /// <returns>Invalid-operation exception for the blocked focus operation.</returns>
        private static InvalidOperationException Refusal(string message) => new InvalidOperationException(message);

        /// <summary>Uses user32 to inspect and focus the current thread's native VBE windows.</summary>
        internal sealed class NativeWindows : IWindows
        {

            /// <summary>Defines the enum window callback.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <param name="parameter">Native handle that supplies the parameter for this operation.</param>
            /// <returns>Boolean indicating the result of the check for operation on native windows.</returns>
            private delegate bool EnumWindow(IntPtr window, IntPtr parameter);

            /// <summary>Returns current thread id for native windows.</summary>
            /// <returns>uint produced by the operation for get current thread id on native windows.</returns>
            [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();

            /// <summary>Determines whether window for native windows.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <returns>Boolean indicating the result of the check for is window on native windows.</returns>
            [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr window);

            /// <summary>Determines whether window visible for native windows.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <returns>Boolean indicating the result of the check for is window visible on native windows.</returns>
            [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);

            /// <summary>Determines whether window enabled for native windows.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <returns>Boolean indicating the result of the check for is window enabled on native windows.</returns>
            [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr window);

            /// <summary>Returns parent for native windows.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <returns>int ptr produced by the operation for get parent on native windows.</returns>
            [DllImport("user32.dll")] private static extern IntPtr GetParent(IntPtr window);

            /// <summary>Returns window thread process id for native windows.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <param name="process">uint that supplies the process for this operation.</param>
            /// <returns>uint produced by the operation for get window thread process id on native windows.</returns>
            [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);

            /// <summary>Returns class name for native windows.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <param name="text">string builder that supplies the text for this operation.</param>
            /// <param name="capacity">int that supplies the capacity for this operation.</param>
            /// <returns>int produced by the operation for get class name on native windows.</returns>
            [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder text, int capacity);

            /// <summary>Returns window text for native windows.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <param name="text">string builder that supplies the text for this operation.</param>
            /// <param name="capacity">int that supplies the capacity for this operation.</param>
            /// <returns>int produced by the operation for get window text on native windows.</returns>
            [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr window, StringBuilder text, int capacity);

            /// <summary>Returns window text length for native windows.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <returns>int produced by the operation for get window text length on native windows.</returns>
            [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowTextLength(IntPtr window);

            /// <summary>Handles enum child windows for native windows.</summary>
            /// <param name="parent">Native handle that supplies the parent for this operation.</param>
            /// <param name="callback">enum window that supplies the callback for this operation.</param>
            /// <param name="parameter">Native handle that supplies the parameter for this operation.</param>
            /// <returns>Boolean indicating the result of the check for enum child windows on native windows.</returns>
            [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, EnumWindow callback, IntPtr parameter);

            /// <summary>Handles send message for native windows.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <param name="message">uint that supplies the message for this operation.</param>
            /// <param name="parameter">Native handle that supplies the parameter for this operation.</param>
            /// <param name="unused">Native handle that supplies the unused for this operation.</param>
            /// <returns>int ptr produced by the operation for send message on native windows.</returns>
            [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "SendMessageW")]
            private static extern IntPtr SendMessage(IntPtr window, uint message, IntPtr parameter, IntPtr unused);

            /// <summary>Handles native set focus for native windows.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <returns>int ptr produced by the operation for native set focus on native windows.</returns>
            [DllImport("user32.dll", EntryPoint = "SetFocus")] private static extern IntPtr NativeSetFocus(IntPtr window);

            /// <summary>Returns focus for native windows.</summary>
            /// <returns>int ptr produced by the operation for get focus on native windows.</returns>
            [DllImport("user32.dll")] private static extern IntPtr GetFocus();

            /// <summary>Gets the current process ID from <see cref="Process.GetCurrentProcess"/>.</summary>
            /// <value>Current process ID.</value>
            public uint CurrentProcess { get { using (var process = Process.GetCurrentProcess()) return (uint)process.Id; } }

            /// <summary>Gets the Win32 ID of the current UI thread.</summary>
            /// <value>Current current thread exposed by native windows.</value>
            public uint CurrentThread => GetCurrentThreadId();

            /// <summary>Reads native identity and state for one handle using user32 queries.</summary>
            /// <param name="handle">HWND to inspect.</param>
            /// <returns>Snapshot including existence, visibility, enabled state, ancestry, class, process, and thread.</returns>
            public Window Read(IntPtr handle)
            {
                uint process;
                uint thread = GetWindowThreadProcessId(handle, out process);
                var kind = new StringBuilder(256);
                GetClassName(handle, kind, kind.Capacity);
                return new Window { Handle = handle, Parent = GetParent(handle), Process = process, Thread = thread,
                    Class = kind.ToString(), Exists = IsWindow(handle), Visible = IsWindowVisible(handle), Enabled = IsWindowEnabled(handle) };
            }

            /// <summary>Enumerates descendant HWNDs and refuses trees larger than the inspection cap.</summary>
            /// <param name="parent">Root whose descendant windows are enumerated by user32.</param>
            /// <returns>Enumerated handles; at most 512 are accepted.</returns>
            public IntPtr[] Children(IntPtr parent)
            {
                var children = new List<IntPtr>();
                bool exceeded = false;
                EnumChildWindows(parent, (window, unused) =>
                {
                    if (children.Count >= 512) { exceeded = true; return false; }
                    children.Add(window); return true;
                }, IntPtr.Zero);
                if (exceeded) throw Refusal("The owned native window tree exceeds inspection limits.");
                return children.ToArray();
            }

            /// <summary>Reads a caption only when its reported length is between 1 and 1,024 characters.</summary>
            /// <param name="handle">Eligible HWND whose title is inspected.</param>
            /// <returns>Exact caption when the full reported length is read; otherwise null.</returns>
            public string Caption(IntPtr handle)
            {
                int length = GetWindowTextLength(handle);
                if (length <= 0 || length > 1024) return null;
                var text = new StringBuilder(1025);
                int read = GetWindowText(handle, text, text.Capacity);
                return read == length ? text.ToString() : null;
            }

            /// <summary>Sends the MDI activate message to make the selected child active.</summary>
            /// <param name="mdi">MDI client receiving the activation message.</param>
            /// <param name="child">Child HWND encoded in the activation message.</param>
            public void Activate(IntPtr mdi, IntPtr child) { SendMessage(mdi, 0x0222, child, IntPtr.Zero); }

            /// <summary>Calls user32 SetFocus for the selected code-window HWND.</summary>
            /// <param name="child">Window requested to receive keyboard focus.</param>
            public void Focus(IntPtr child) { NativeSetFocus(child); }

            /// <summary>Reads the focused HWND for the current thread.</summary>
            /// <value>Current focus handle, or zero when none exists.</value>
            public IntPtr Focused => GetFocus();
        }
    }
}
