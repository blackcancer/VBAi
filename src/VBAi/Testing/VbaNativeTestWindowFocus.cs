using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace VBAi
{
    internal static class VbaNativeTestWindowFocus
    {
        internal sealed class Window
        {
            internal IntPtr Handle, Parent;
            internal uint Process, Thread;
            internal string Class;
            internal bool Exists, Visible, Enabled;
        }

        internal interface IWindows
        {
            uint CurrentProcess { get; }
            uint CurrentThread { get; }
            Window Read(IntPtr handle);
            IntPtr[] Children(IntPtr parent);
            string Caption(IntPtr handle);
            void Activate(IntPtr mdi, IntPtr child);
            void Focus(IntPtr child);
            IntPtr Focused { get; }
        }

        internal sealed class Target
        {
            internal Window Main, Mdi, Code;
            internal string Caption;
        }

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

        private static void VerifyUnchanged(IWindows windows, Window expected, string role)
        {
            var current = windows.Read(expected.Handle);
            RequireOwned(windows, current, role);
            if (current.Handle != expected.Handle || current.Parent != expected.Parent || current.Class != expected.Class
                || current.Process != expected.Process || current.Thread != expected.Thread)
                throw Refusal("The owned native window identity changed before focus. "
                    + DescribeWindow(windows, current, role) + "; expected={" + DescribeWindow(windows, expected, role) + "}");
        }

        private static bool Eligible(IWindows windows, Window window) => window != null && window.Exists
            && window.Handle != IntPtr.Zero && window.Visible && window.Enabled
            && window.Process != 0 && window.Process == windows.CurrentProcess
            && window.Thread != 0 && window.Thread == windows.CurrentThread;

        private static void RequireOwned(IWindows windows, Window window, string role)
        {
            if (!Eligible(windows, window)) throw Refusal("The native window is not visible and enabled on the owning process and UI thread. "
                + DescribeWindow(windows, window, role));
        }

        // Diagnostic data comes from the already observed native snapshot. Never
        // read captions, focus another window or enumerate more windows on refusal.
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

        private static InvalidOperationException Refusal(string message) => new InvalidOperationException(message);

        internal sealed class NativeWindows : IWindows
        {
            private delegate bool EnumWindow(IntPtr window, IntPtr parameter);
            [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
            [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr window);
            [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
            [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr window);
            [DllImport("user32.dll")] private static extern IntPtr GetParent(IntPtr window);
            [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
            [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder text, int capacity);
            [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr window, StringBuilder text, int capacity);
            [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowTextLength(IntPtr window);
            [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, EnumWindow callback, IntPtr parameter);
            [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "SendMessageW")]
            private static extern IntPtr SendMessage(IntPtr window, uint message, IntPtr parameter, IntPtr unused);
            [DllImport("user32.dll", EntryPoint = "SetFocus")] private static extern IntPtr NativeSetFocus(IntPtr window);
            [DllImport("user32.dll")] private static extern IntPtr GetFocus();
            public uint CurrentProcess { get { using (var process = Process.GetCurrentProcess()) return (uint)process.Id; } }
            public uint CurrentThread => GetCurrentThreadId();
            public Window Read(IntPtr handle)
            {
                uint process;
                uint thread = GetWindowThreadProcessId(handle, out process);
                var kind = new StringBuilder(256);
                GetClassName(handle, kind, kind.Capacity);
                return new Window { Handle = handle, Parent = GetParent(handle), Process = process, Thread = thread,
                    Class = kind.ToString(), Exists = IsWindow(handle), Visible = IsWindowVisible(handle), Enabled = IsWindowEnabled(handle) };
            }
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
            public string Caption(IntPtr handle)
            {
                int length = GetWindowTextLength(handle);
                if (length <= 0 || length > 1024) return null;
                var text = new StringBuilder(1025);
                int read = GetWindowText(handle, text, text.Capacity);
                return read == length ? text.ToString() : null;
            }
            public void Activate(IntPtr mdi, IntPtr child) { SendMessage(mdi, 0x0222, child, IntPtr.Zero); }
            public void Focus(IntPtr child) { NativeSetFocus(child); }
            public IntPtr Focused => GetFocus();
        }
    }
}
