using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Automation;

namespace CodexVBE
{
    // Run these accessibility calls away from the VBE UI thread. UIA and MSAA
    // query native windows, not the VBIDE object model used by VbeDebug.
    internal static class VbeDebugWindows
    {
        private delegate bool EnumWindowCallback(IntPtr handle, IntPtr parameter);

        [DllImport("user32.dll")] private static extern bool EnumWindows(EnumWindowCallback callback, IntPtr parameter);
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, EnumWindowCallback callback, IntPtr parameter);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr handle, out uint processId);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr handle, StringBuilder text, int capacity);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr handle, StringBuilder text, int capacity);
        [DllImport("user32.dll")] private static extern int GetDlgCtrlID(IntPtr handle);
        [DllImport("user32.dll")] private static extern IntPtr GetDlgItem(IntPtr handle, int controlId);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr handle);
        [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr handle, int message, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "SendMessageW")]
        private static extern IntPtr SendMessageText(IntPtr handle, int message, IntPtr wParam, string text);
        [DllImport("user32.dll", EntryPoint = "SendMessageW")]
        private static extern IntPtr SendMessageInt(IntPtr handle, int message, IntPtr wParam, IntPtr lParam);
        [DllImport("oleacc.dll")] private static extern int AccessibleObjectFromWindow(IntPtr handle, uint objectId,
            ref Guid interfaceId, [MarshalAs(UnmanagedType.Interface)] out object accessible);

        private const int BmClick = 0x00F5;
        private const uint ObjidClient = 4294967292;
        private static readonly Guid IidAccessible = new Guid("618736e0-3c3d-11cf-810c-00aa00389b71");

        public static object Capture(bool includeCallStack)
        {
            IntPtr root = FindVbeRoot();
            if (root == IntPtr.Zero) throw new InvalidOperationException("The VBE window is not open in this host process.");
            var panes = ChildWindows(root);
            IntPtr locals = FindPane(panes, "Variables locales", "Locals");
            IntPtr watches = FindPane(panes, "Espions", "Watch", "Watches");
            IntPtr immediate = FindPane(panes, "Exécution", "Immediate");
            object stack = includeCallStack ? ReadCallStack(locals) : null;
            return new {
                HostProcessId = Process.GetCurrentProcess().Id,
                Locals = ReadList(locals),
                Watches = ReadList(watches),
                Immediate = ReadImmediate(immediate),
                CallStack = stack,
                Limits = "Native UI accessibility is observed only for visible panes. A missing pane or unavailable value is not an empty debugger collection. Breakpoints and exception state are not exposed by this snapshot."
            };
        }

        public static object CompleteAddWatch(Request request)
        {
            IntPtr dialog = IntPtr.Zero;
            for (int attempt = 0; attempt < 60 && dialog == IntPtr.Zero; attempt++)
            { Thread.Sleep(50); dialog = FindDialog("Ajouter un espion", "Add Watch"); }
            if (dialog == IntPtr.Zero) throw new InvalidOperationException("The Add Watch dialog did not open.");
            bool completed = false;
            try
            {
                IntPtr edit = GetDlgItem(dialog, 4853);
                IntPtr project = GetDlgItem(dialog, 4858);
                IntPtr module = GetDlgItem(dialog, 4857);
                IntPtr procedure = GetDlgItem(dialog, 4856);
                IntPtr ok = GetDlgItem(dialog, 1);
                if (edit == IntPtr.Zero || project == IntPtr.Zero || module == IntPtr.Zero || ok == IntPtr.Zero)
                    throw new InvalidOperationException("The native Add Watch dialog controls changed.");
                string shownProject = WindowText(project);
                string shownModule = WindowText(module);
                string shownProcedure = procedure == IntPtr.Zero ? null : WindowText(procedure);
                if (!string.Equals(shownProject, request.Project, StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(shownModule, request.Module, StringComparison.OrdinalIgnoreCase) ||
                    (!string.IsNullOrWhiteSpace(request.Procedure) &&
                     !string.Equals(shownProcedure, request.Procedure, StringComparison.OrdinalIgnoreCase)))
                    throw new InvalidOperationException("Add Watch context changed: " + shownProject + "." + shownModule + "." + shownProcedure);
                // EM_REPLACESEL triggers the VBE's edit notifications. WM_SETTEXT
                // alone changes the visible text but is rejected as an empty expression.
                SendMessageInt(edit, 0x00B1, IntPtr.Zero, new IntPtr(-1)); // EM_SETSEL
                SendMessageText(edit, 0x00C2, new IntPtr(1), request.Expression); // EM_REPLACESEL
                if (!string.Equals(WindowText(edit), request.Expression, StringComparison.Ordinal))
                    throw new InvalidOperationException("The watch expression was not reflected by the native edit control.");
                if (!PostMessage(ok, BmClick, IntPtr.Zero, IntPtr.Zero))
                    throw new InvalidOperationException("The Add Watch dialog refused its OK command.");
                IntPtr error = IntPtr.Zero;
                for (int attempt = 0; attempt < 40; attempt++)
                {
                    Thread.Sleep(50);
                    dialog = FindDialog("Ajouter un espion", "Add Watch");
                    if (dialog == IntPtr.Zero) { completed = true; break; }
                    error = FindDialog("Microsoft Visual Basic pour Applications", "Microsoft Visual Basic for Applications");
                    if (error != IntPtr.Zero) break;
                }
                if (error != IntPtr.Zero)
                {
                    string detail = AccessibleDialogMessage(error);
                    CloseDialog(error);
                    throw new InvalidOperationException("VBE rejected the watch expression: " + detail);
                }
                if (!completed) throw new InvalidOperationException("The Add Watch dialog did not close after OK.");
                IntPtr root = FindVbeRoot();
                IntPtr watches = root == IntPtr.Zero ? IntPtr.Zero :
                    FindPane(ChildWindows(root), "Espions", "Watch", "Watches");
                object watchState = ReadList(watches);
                return new { Added = true, request.Expression, Context = new {
                    Project = shownProject, Module = shownModule, Procedure = shownProcedure },
                    Watches = watchState,
                    Verification = watches == IntPtr.Zero ? "Pending" : "ReadbackAvailable",
                    NextRead = watches == IntPtr.Zero ?
                        "Open the Watches pane and call debug_windows in a separate request to verify the expression and value." :
                        "Call debug_windows in a separate request to confirm the watch value after VBE refresh." };
            }
            finally
            {
                if (!completed)
                {
                    IntPtr remaining = FindDialog("Ajouter un espion", "Add Watch");
                    if (remaining != IntPtr.Zero) CloseDialog(remaining);
                }
            }
        }

        public static object SelectWatch(Request request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Expression) ||
                string.IsNullOrWhiteSpace(request.Context))
                throw new ArgumentException("Expression and Context are required.");
            IntPtr root = FindVbeRoot();
            IntPtr pane = root == IntPtr.Zero ? IntPtr.Zero : FindPane(ChildWindows(root), "Espions", "Watch", "Watches");
            if (pane == IntPtr.Zero) throw new InvalidOperationException("The native Watches pane must be visible.");
            var rows = MatchingWatchRows(pane, request.Expression, request.Context);
            if (rows.Count != 1)
                throw new InvalidOperationException("Expected one matching native watch; found " + rows.Count + ".");
            object pattern;
            if (!rows[0].TryGetCurrentPattern(SelectionItemPattern.Pattern, out pattern))
                throw new InvalidOperationException("The native watch row does not support selection.");
            ((SelectionItemPattern)pattern).Select();
            rows[0].SetFocus();
            return new { Selected = true, request.Expression, request.Context };
        }

        public static object VerifyWatchRemoved(Request request)
        {
            IntPtr root = FindVbeRoot();
            IntPtr pane = root == IntPtr.Zero ? IntPtr.Zero : FindPane(ChildWindows(root), "Espions", "Watch", "Watches");
            if (pane == IntPtr.Zero)
                return new { Removed = false, VerificationPending = true,
                    Error = "The Watches pane is no longer visible; absence cannot be verified." };
            for (int attempt = 0; attempt < 20; attempt++)
            {
                if (MatchingWatchRows(pane, request.Expression, request.Context).Count == 0)
                    return new { Removed = true, VerificationPending = false, Error = (string)null };
                Thread.Sleep(50);
            }
            return new { Removed = false, VerificationPending = true,
                Error = "The selected watch is still visible after the native command." };
        }

        private static List<AutomationElement> MatchingWatchRows(IntPtr pane, string expression, string context)
        {
            AutomationElement root = AutomationElement.FromHandle(pane);
            var condition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem);
            AutomationElementCollection elements = root.FindAll(TreeScope.Descendants, condition);
            var matches = new List<AutomationElement>();
            for (int index = 0; index < elements.Count; index++)
            {
                string raw = elements[index].Current.Name ?? "";
                Match parsed = Regex.Match(raw, @"^\s*(.*?)\s+(?:Valeur|Value)\s+.*?\s+Type\s+.*?\s+(?:Contexte|Context)\s+(.*?)\s*$",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline);
                if (parsed.Success && string.Equals(parsed.Groups[1].Value, expression, StringComparison.Ordinal) &&
                    string.Equals(parsed.Groups[2].Value, context, StringComparison.OrdinalIgnoreCase))
                    matches.Add(elements[index]);
            }
            return matches;
        }

        private static string AccessibleDialogMessage(IntPtr dialog)
        {
            AutomationElement root = AutomationElement.FromHandle(dialog);
            AutomationElementCollection elements = root.FindAll(TreeScope.Descendants, Condition.TrueCondition);
            for (int index = 0; index < elements.Count; index++)
            {
                string name = elements[index].Current.Name;
                if (!string.IsNullOrWhiteSpace(name) && name != "OK" && name != "Aide" && name != "Help") return name;
            }
            return "Native validation error.";
        }

        private static void CloseDialog(IntPtr dialog)
        {
            IntPtr cancel = GetDlgItem(dialog, 2);
            if (cancel != IntPtr.Zero) PostMessage(cancel, BmClick, IntPtr.Zero, IntPtr.Zero);
        }

        private static IntPtr FindDialog(params string[] titles)
        {
            IntPtr result = IntPtr.Zero;
            uint currentPid = (uint)Process.GetCurrentProcess().Id;
            EnumWindows((handle, parameter) => {
                uint pid;
                GetWindowThreadProcessId(handle, out pid);
                if (pid != currentPid || ClassName(handle) != "#32770" || !IsWindowVisible(handle)) return true;
                string title = WindowText(handle);
                foreach (string candidate in titles)
                    if (string.Equals(title, candidate, StringComparison.OrdinalIgnoreCase))
                    { result = handle; return false; }
                return true;
            }, IntPtr.Zero);
            return result;
        }

        private static IntPtr FindVbeRoot()
        {
            IntPtr result = IntPtr.Zero;
            uint currentPid = (uint)Process.GetCurrentProcess().Id;
            EnumWindows((handle, parameter) => {
                uint pid;
                GetWindowThreadProcessId(handle, out pid);
                if (pid == currentPid && ClassName(handle) == "wndclass_desked_gsk" && IsWindowVisible(handle))
                { result = handle; return false; }
                return true;
            }, IntPtr.Zero);
            return result;
        }

        private static List<IntPtr> ChildWindows(IntPtr root)
        {
            var result = new List<IntPtr>();
            EnumChildWindows(root, (handle, parameter) => {
                if (ClassName(handle) == "VbaWindow" && IsWindowVisible(handle)) result.Add(handle);
                return true;
            }, IntPtr.Zero);
            return result;
        }

        private static IntPtr FindPane(IEnumerable<IntPtr> panes, params string[] names)
        {
            foreach (IntPtr pane in panes)
            {
                string title = WindowText(pane);
                foreach (string name in names)
                    if (string.Equals(title, name, StringComparison.OrdinalIgnoreCase)) return pane;
            }
            return IntPtr.Zero;
        }

        private static object ReadList(IntPtr handle)
        {
            if (handle == IntPtr.Zero) return new { Available = false, Items = new object[0], Error = "Window is not visible." };
            try
            {
                AutomationElement root = AutomationElement.FromHandle(handle);
                var listCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem);
                AutomationElementCollection elements = root.FindAll(TreeScope.Descendants, listCondition);
                var items = new List<object>();
                for (int index = 0; index < elements.Count; index++)
                {
                    AutomationElement element = elements[index];
                    string raw = element.Current.Name;
                    object pattern;
                    if (element.TryGetCurrentPattern(ValuePattern.Pattern, out pattern))
                        raw = ((ValuePattern)pattern).Current.Value ?? raw;
                    // MSForms' localized list rows flatten all columns into one accessible string.
                    Match match = Regex.Match(raw ?? "", @"^Expression\s+(.*?)\s+Valeur\s+(.*?)\s+Type\s+(.*?)\s*$",
                        RegexOptions.IgnoreCase | RegexOptions.Singleline);
                    if (!match.Success)
                        match = Regex.Match(raw ?? "", @"^Expression\s+(.*?)\s+Value\s+(.*?)\s+Type\s+(.*?)\s*$",
                            RegexOptions.IgnoreCase | RegexOptions.Singleline);
                    Match watch = Regex.Match(raw ?? "", @"^\s*(.*?)\s+Valeur\s+(.*?)\s+Type\s+(.*?)\s+Contexte\s+(.*?)\s*$",
                        RegexOptions.IgnoreCase | RegexOptions.Singleline);
                    if (!watch.Success)
                        watch = Regex.Match(raw ?? "", @"^\s*(.*?)\s+Value\s+(.*?)\s+Type\s+(.*?)\s+Context\s+(.*?)\s*$",
                            RegexOptions.IgnoreCase | RegexOptions.Singleline);
                    if (match.Success && string.IsNullOrWhiteSpace(match.Groups[1].Value) &&
                        (match.Groups[2].Value.IndexOf("Aucune variable", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         match.Groups[2].Value.IndexOf("No variables", StringComparison.OrdinalIgnoreCase) >= 0))
                        continue;
                    items.Add(new { Raw = raw, Expression = match.Success ? match.Groups[1].Value :
                            watch.Success ? watch.Groups[1].Value : null,
                        Value = match.Success ? match.Groups[2].Value.Trim() :
                            watch.Success ? watch.Groups[2].Value.Trim() : null,
                        Type = match.Success ? match.Groups[3].Value :
                            watch.Success ? watch.Groups[3].Value : null,
                        Context = watch.Success ? watch.Groups[4].Value : null,
                        Parsed = match.Success || watch.Success });
                }
                return new { Available = true, Items = items.ToArray(), Error = (string)null };
            }
            catch (Exception ex) { return new { Available = true, Items = new object[0], Error = ex.Message }; }
        }

        private static object ReadImmediate(IntPtr handle)
        {
            if (handle == IntPtr.Zero) return new { Available = false, Text = (string)null, Error = "Window is not visible." };
            try
            {
                AutomationElement root = AutomationElement.FromHandle(handle);
                var condition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Document);
                AutomationElementCollection documents = root.FindAll(TreeScope.Descendants, condition);
                if (documents.Count != 1) throw new InvalidOperationException("Expected one Immediate document; found " + documents.Count + ".");
                object pattern;
                if (!documents[0].TryGetCurrentPattern(TextPattern.Pattern, out pattern))
                    throw new InvalidOperationException("The Immediate document does not expose TextPattern.");
                return new { Available = true, Text = ((TextPattern)pattern).DocumentRange.GetText(-1), Error = (string)null };
            }
            catch (Exception ex) { return new { Available = true, Text = (string)null, Error = ex.Message }; }
        }

        private static object ReadCallStack(IntPtr locals)
        {
            IntPtr dialog = FindCallStackDialog();
            bool openedHere = dialog == IntPtr.Zero;
            if (openedHere)
            {
                if (locals == IntPtr.Zero)
                    return new { Available = false, Frames = new string[0], Error = "Locals window must be visible to open Call Stack without a shortcut." };
                IntPtr button = IntPtr.Zero;
                EnumChildWindows(locals, (handle, parameter) => {
                    if (ClassName(handle) == "Button" && GetDlgCtrlID(handle) == 4601 && IsWindowVisible(handle))
                    { button = handle; return false; }
                    return true;
                }, IntPtr.Zero);
                if (button == IntPtr.Zero || !PostMessage(button, BmClick, IntPtr.Zero, IntPtr.Zero))
                    return new { Available = false, Frames = new string[0], Error = "The native Call Stack button was not available." };
                for (int attempt = 0; attempt < 30 && dialog == IntPtr.Zero; attempt++)
                { Thread.Sleep(50); dialog = FindCallStackDialog(); }
            }
            if (dialog == IntPtr.Zero)
                return new { Available = false, Frames = new string[0], Error = "Call Stack dialog did not open." };
            try
            {
                IntPtr list = IntPtr.Zero;
                EnumChildWindows(dialog, (handle, parameter) => {
                    if (ClassName(handle) == "ListBox") { list = handle; return false; }
                    return true;
                }, IntPtr.Zero);
                if (list == IntPtr.Zero) throw new InvalidOperationException("Call Stack list was not found.");
                object accessible;
                Guid iid = IidAccessible;
                int hr = AccessibleObjectFromWindow(list, ObjidClient, ref iid, out accessible);
                if (hr != 0 || !(accessible is Accessibility.IAccessible))
                    throw new COMException("Call Stack MSAA object was unavailable.", hr);
                var listAccess = (Accessibility.IAccessible)accessible;
                var frames = new List<string>();
                for (int index = 1; index <= Math.Min(listAccess.accChildCount, 256); index++)
                    frames.Add(listAccess.get_accName(index));
                return new { Available = true, Frames = frames.ToArray(), Error = (string)null };
            }
            catch (Exception ex) { return new { Available = true, Frames = new string[0], Error = ex.Message }; }
            finally
            {
                if (openedHere)
                {
                    EnumChildWindows(dialog, (handle, parameter) => {
                        if (ClassName(handle) == "Button" && GetDlgCtrlID(handle) == 2)
                        { PostMessage(handle, BmClick, IntPtr.Zero, IntPtr.Zero); return false; }
                        return true;
                    }, IntPtr.Zero);
                }
            }
        }

        private static IntPtr FindCallStackDialog()
        {
            IntPtr result = IntPtr.Zero;
            uint currentPid = (uint)Process.GetCurrentProcess().Id;
            EnumWindows((handle, parameter) => {
                uint pid;
                GetWindowThreadProcessId(handle, out pid);
                string title = WindowText(handle);
                if (pid == currentPid && ClassName(handle) == "#32770" && IsWindowVisible(handle) &&
                    (title == "Pile des appels" || title == "Call Stack"))
                { result = handle; return false; }
                return true;
            }, IntPtr.Zero);
            return result;
        }

        private static string ClassName(IntPtr handle)
        { var text = new StringBuilder(128); GetClassName(handle, text, text.Capacity); return text.ToString(); }
        private static string WindowText(IntPtr handle)
        { var text = new StringBuilder(512); GetWindowText(handle, text, text.Capacity); return text.ToString(); }
    }
}
