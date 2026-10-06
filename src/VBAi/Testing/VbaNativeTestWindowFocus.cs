using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace VBAi
{

    /// <summary>Owns the vba native test window focus state and operations.</summary>
    internal static class VbaNativeTestWindowFocus
    {

        /// <summary>Owns the window state and operations.</summary>
        internal sealed class Window
        {

            /// <summary>Maintains the handle and parent state for window.</summary>
            internal IntPtr Handle, Parent;

            /// <summary>Maintains the process and thread state for window.</summary>
            internal uint Process, Thread;

            /// <summary>Maintains the class state for window.</summary>
            internal string Class;

            /// <summary>Maintains the exists and visible and enabled state for window.</summary>
            internal bool Exists, Visible, Enabled;
        }

        /// <summary>Defines the i windows contract.</summary>
        internal interface IWindows
        {

            /// <summary>Gets the current process.</summary>
            /// <value>Current current process exposed by i windows.</value>
            uint CurrentProcess { get; }

            /// <summary>Gets the current thread.</summary>
            /// <value>Current current thread exposed by i windows.</value>
            uint CurrentThread { get; }

            /// <summary>Reads  for i windows.</summary>
            /// <param name="handle">Native handle that supplies the handle for this operation.</param>
            /// <returns>window produced by the operation for read on i windows.</returns>
            Window Read(IntPtr handle);

            /// <summary>Handles children for i windows.</summary>
            /// <param name="parent">Native handle that supplies the parent for this operation.</param>
            /// <returns>int ptr[] produced by the operation for children on i windows.</returns>
            IntPtr[] Children(IntPtr parent);

            /// <summary>Handles caption for i windows.</summary>
            /// <param name="handle">Native handle that supplies the handle for this operation.</param>
            /// <returns>Text produced by the operation for caption on i windows.</returns>
            string Caption(IntPtr handle);

            /// <summary>Handles activate for i windows.</summary>
            /// <param name="mdi">Native handle that supplies the mdi for this operation.</param>
            /// <param name="child">Native handle that supplies the child for this operation.</param>
            void Activate(IntPtr mdi, IntPtr child);

            /// <summary>Handles focus for i windows.</summary>
            /// <param name="child">Native handle that supplies the child for this operation.</param>
            void Focus(IntPtr child);

            /// <summary>Gets the focused.</summary>
            /// <value>Current focused exposed by i windows.</value>
            IntPtr Focused { get; }
        }

        /// <summary>Owns the target state and operations.</summary>
        internal sealed class Target
        {

            /// <summary>Maintains the main and mdi and code state for target.</summary>
            internal Window Main, Mdi, Code;

            /// <summary>Maintains the caption state for target.</summary>
            internal string Caption;
        }

        /// <summary>Handles focus for vba native test window focus.</summary>
        /// <param name="windows">i windows that supplies the windows for this operation.</param>
        /// <param name="main">Native handle that supplies the main for this operation.</param>
        /// <param name="caption">Text that supplies the caption value. Use the format required by the calling operation.</param>
        /// <returns>target produced by the operation for focus on vba native test window focus.</returns>
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

        /// <summary>Handles verify focus for vba native test window focus.</summary>
        /// <param name="windows">i windows that supplies the windows for this operation.</param>
        /// <param name="target">target that supplies the target for this operation.</param>
        /// <param name="stage">Text that supplies the stage value. Use the format required by the calling operation.</param>
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

        /// <summary>Handles verify target for vba native test window focus.</summary>
        /// <param name="windows">i windows that supplies the windows for this operation.</param>
        /// <param name="target">target that supplies the target for this operation.</param>
        /// <param name="stage">Text that supplies the stage value. Use the format required by the calling operation.</param>
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

        /// <summary>Handles verify unchanged for vba native test window focus.</summary>
        /// <param name="windows">i windows that supplies the windows for this operation.</param>
        /// <param name="expected">window that supplies the expected for this operation.</param>
        /// <param name="role">Text that supplies the role value. Use the format required by the calling operation.</param>
        private static void VerifyUnchanged(IWindows windows, Window expected, string role)
        {
            var current = windows.Read(expected.Handle);
            RequireOwned(windows, current, role);
            if (current.Handle != expected.Handle || current.Parent != expected.Parent || current.Class != expected.Class
                || current.Process != expected.Process || current.Thread != expected.Thread)
                throw Refusal("The owned native window identity changed before focus. "
                    + DescribeWindow(windows, current, role) + "; expected={" + DescribeWindow(windows, expected, role) + "}");
        }

        /// <summary>Handles eligible for vba native test window focus.</summary>
        /// <param name="windows">i windows that supplies the windows for this operation.</param>
        /// <param name="window">window that supplies the window for this operation.</param>
        /// <returns>Boolean indicating the result of the check for eligible on vba native test window focus.</returns>
        private static bool Eligible(IWindows windows, Window window) => window != null && window.Exists
            && window.Handle != IntPtr.Zero && window.Visible && window.Enabled
            && window.Process != 0 && window.Process == windows.CurrentProcess
            && window.Thread != 0 && window.Thread == windows.CurrentThread;

        /// <summary>Requires owned for vba native test window focus.</summary>
        /// <param name="windows">i windows that supplies the windows for this operation.</param>
        /// <param name="window">window that supplies the window for this operation.</param>
        /// <param name="role">Text that supplies the role value. Use the format required by the calling operation.</param>
        private static void RequireOwned(IWindows windows, Window window, string role)
        {
            if (!Eligible(windows, window)) throw Refusal("The native window is not visible and enabled on the owning process and UI thread. "
                + DescribeWindow(windows, window, role));
        }

        // Diagnostic data comes from the already observed native snapshot. Never
        // read captions, focus another window or enumerate more windows on refusal.
        /// <summary>Handles describe window for vba native test window focus.</summary>
        /// <param name="windows">i windows that supplies the windows for this operation.</param>
        /// <param name="window">window that supplies the window for this operation.</param>
        /// <param name="role">Text that supplies the role value. Use the format required by the calling operation.</param>
        /// <returns>Text produced by the operation for describe window on vba native test window focus.</returns>
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

        /// <summary>Handles refusal for vba native test window focus.</summary>
        /// <param name="message">Text that supplies the message value. Use the format required by the calling operation.</param>
        /// <returns>invalid operation exception produced by the operation for refusal on vba native test window focus.</returns>
        private static InvalidOperationException Refusal(string message) => new InvalidOperationException(message);

        /// <summary>Owns the native windows state and operations.</summary>
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

            /// <summary>Gets the current process.</summary>
            /// <value>Current current process exposed by native windows.</value>
            public uint CurrentProcess { get { using (var process = Process.GetCurrentProcess()) return (uint)process.Id; } }

            /// <summary>Gets the current thread.</summary>
            /// <value>Current current thread exposed by native windows.</value>
            public uint CurrentThread => GetCurrentThreadId();

            /// <summary>Reads  for native windows.</summary>
            /// <param name="handle">Native handle that supplies the handle for this operation.</param>
            /// <returns>window produced by the operation for read on native windows.</returns>
            public Window Read(IntPtr handle)
            {
                uint process;
                uint thread = GetWindowThreadProcessId(handle, out process);
                var kind = new StringBuilder(256);
                GetClassName(handle, kind, kind.Capacity);
                return new Window { Handle = handle, Parent = GetParent(handle), Process = process, Thread = thread,
                    Class = kind.ToString(), Exists = IsWindow(handle), Visible = IsWindowVisible(handle), Enabled = IsWindowEnabled(handle) };
            }

            /// <summary>Handles children for native windows.</summary>
            /// <param name="parent">Native handle that supplies the parent for this operation.</param>
            /// <returns>int ptr[] produced by the operation for children on native windows.</returns>
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

            /// <summary>Handles caption for native windows.</summary>
            /// <param name="handle">Native handle that supplies the handle for this operation.</param>
            /// <returns>Text produced by the operation for caption on native windows.</returns>
            public string Caption(IntPtr handle)
            {
                int length = GetWindowTextLength(handle);
                if (length <= 0 || length > 1024) return null;
                var text = new StringBuilder(1025);
                int read = GetWindowText(handle, text, text.Capacity);
                return read == length ? text.ToString() : null;
            }

            /// <summary>Handles activate for native windows.</summary>
            /// <param name="mdi">Native handle that supplies the mdi for this operation.</param>
            /// <param name="child">Native handle that supplies the child for this operation.</param>
            public void Activate(IntPtr mdi, IntPtr child) { SendMessage(mdi, 0x0222, child, IntPtr.Zero); }

            /// <summary>Handles focus for native windows.</summary>
            /// <param name="child">Native handle that supplies the child for this operation.</param>
            public void Focus(IntPtr child) { NativeSetFocus(child); }

            /// <summary>Gets the focused.</summary>
            /// <value>Current focused exposed by native windows.</value>
            public IntPtr Focused => GetFocus();
        }
    }
}
