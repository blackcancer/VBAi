using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Forms;
using Accessibility;

namespace VBAi.Desktop.Helper
{
    /// <summary>
    /// Synthetic, owned WinForms action discriminator for an inactive test desktop.
    /// The caller chooses when to run it. This library has no Office, VBE, Git or provider entry point.
    /// Every action is delivered at most once; a failed or uncertain delivery is reported as a gap.
    /// </summary>
    public static class PrivateUiActionCanary
    {
        private const uint BmClick = 0x00F5;
        private const uint SmtoAbortIfHung = 0x0002;
        private const uint ObjidClient = 0xFFFFFFFC;
        private const uint GwOwner = 4;
        private static readonly Guid IAccessibleId = new Guid("618736E0-3C3D-11CF-810C-00AA00389B71");
        private const string GitLabel = "GitHub · synchronize VBA…";

        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        [DllImport("user32.dll", SetLastError = true)] private static extern IntPtr GetThreadDesktop(uint threadId);
        [DllImport("user32.dll", SetLastError = true)] private static extern IntPtr OpenInputDesktop(uint flags, bool inherit, uint access);
        [DllImport("user32.dll", SetLastError = true)] private static extern bool CloseDesktop(IntPtr desktop);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool GetUserObjectInformationW(IntPtr handle, int index, StringBuilder value, uint length, out uint needed);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
        [DllImport("user32.dll")] private static extern bool IsChild(IntPtr parent, IntPtr child);
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr window);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
        [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr window, uint command);
        [DllImport("user32.dll")] private static extern IntPtr GetParent(IntPtr window);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr window, StringBuilder value, int length);
        [DllImport("user32.dll")] private static extern int GetWindowLong(IntPtr window, int index);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SendMessageTimeout(IntPtr window, uint message, IntPtr wParam, IntPtr lParam,
            uint flags, uint timeout, out IntPtr result);
        [DllImport("oleacc.dll", EntryPoint = "AccessibleObjectFromWindow")]
        private static extern int AccessibleObjectFromWindow(IntPtr window, uint objectId, ref Guid iid,
            [MarshalAs(UnmanagedType.Interface)] out IAccessible accessible);

        public static readonly string[] ActionMap = {
            "TextBox: ValuePattern.SetValue -> exact Value/Text readback",
            "TabControl: SelectionItemPattern.Select -> selected tab readback",
            "ComboBox: ExpandCollapsePattern.Expand -> visible list readback",
            "ComboBox: SelectionItemPattern.Select -> exact selected item readback",
            "ComboBox: ExpandCollapsePattern.Collapse -> collapsed state readback",
            "Native Button: one bounded BM_CLICK -> one managed Click event",
            "ChatActionButton Options: one bounded BM_CLICK -> one owned ToolStrip popup",
            "Virtual ToolStrip Git item: one MSAA accDoDefaultAction -> one managed Click event",
            "Owned Form: WindowPattern.Close -> exact HWND destruction and FormClosed event"
        };

        private static string ObjectName(IntPtr handle)
        {
            if (handle == IntPtr.Zero) throw new InvalidOperationException("Desktop handle unavailable.");
            var name = new StringBuilder(256); uint needed;
            if (!GetUserObjectInformationW(handle, 2, name, (uint)(name.Capacity * 2), out needed))
                throw new InvalidOperationException("Desktop name unavailable.");
            return name.ToString();
        }

        private static void RequireInactiveDesktop(string expected)
        {
            Guid identifier;
            if (expected == null || !expected.StartsWith("VBAiTests_", StringComparison.Ordinal) ||
                !Guid.TryParseExact(expected.Substring("VBAiTests_".Length), "N", out identifier))
                throw new InvalidOperationException("Only a generated VBAi test desktop is permitted.");
            if (!string.Equals(ObjectName(GetThreadDesktop(GetCurrentThreadId())), expected, StringComparison.Ordinal))
                throw new InvalidOperationException("The canary thread is outside its exact private desktop.");
            IntPtr input = OpenInputDesktop(0, false, 1);
            try
            {
                if (string.Equals(ObjectName(input), expected, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("The canary must never run on the input desktop.");
            }
            finally { if (input != IntPtr.Zero) CloseDesktop(input); }
        }

        private static void Guard(IntPtr root, IntPtr handle, uint threadId, int processId)
        {
            uint pid; uint tid = GetWindowThreadProcessId(handle, out pid);
            if (handle == IntPtr.Zero || !IsWindow(handle) || tid != threadId || pid != processId ||
                (handle != root && !IsChild(root, handle)))
                throw new InvalidOperationException("Exact owned native control identity changed.");
        }

        private static AutomationElement One(AutomationElement root, string id, int processId)
        {
            var found = root.FindAll(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.AutomationIdProperty, id)).Cast<AutomationElement>().ToArray();
            if (found.Length != 1 || found[0].Current.ProcessId != processId || !found[0].Current.IsEnabled)
                throw new InvalidOperationException("Exact owned UIA control is absent or ambiguous: " + id);
            return found[0];
        }

        private static T Pattern<T>(AutomationElement item, AutomationPattern pattern) where T : class
        {
            object value;
            if (!item.TryGetCurrentPattern(pattern, out value) || !(value is T))
                throw new InvalidOperationException("Required canary pattern unavailable: " + pattern.ProgrammaticName);
            return (T)value;
        }

        private static IntPtr NativeAncestor(AutomationElement item)
        {
            for (int depth = 0; item != null && depth < 32; depth++, item = TreeWalker.RawViewWalker.GetParent(item))
                if (item.Current.NativeWindowHandle != 0) return new IntPtr(item.Current.NativeWindowHandle);
            throw new InvalidOperationException("Virtual item has no bounded native ancestor.");
        }

        private sealed class Observation
        {
            internal readonly bool Proven;
            internal readonly object State;
            internal Observation(bool proven, object state) { Proven = proven; State = state; }
        }

        private static bool Once(Action<object> receipt, string name, string method, Action deliver, Func<Observation> readback)
        {
            receipt(new { Phase = "ActionIntent", Name = name, Method = method, Attempts = 1 });
            try
            {
                deliver();
                receipt(new { Phase = "ActionReturned", Name = name, Method = method, Attempts = 1 });
            }
            catch (Exception error)
            {
                Observation immediate = null;
                string readbackErrorType = null;
                try { immediate = readback(); }
                catch (Exception readbackError) { readbackErrorType = readbackError.GetType().FullName; }
                receipt(new { Phase = "ActionFailedOrUncertain", Name = name, Method = method,
                    ErrorType = error.GetType().FullName, HResult = "0x" + unchecked((uint)error.HResult).ToString("X8"),
                    ImmediateState = immediate == null ? null : immediate.State,
                    ImmediateReadbackProven = immediate != null && immediate.Proven,
                    ReadbackErrorType = readbackErrorType, Attempts = 1, Replay = false });
                return false;
            }
            try
            {
                Observation observation = readback();
                var bound = Stopwatch.StartNew();
                while (!observation.Proven && bound.ElapsedMilliseconds < 1500)
                {
                    Thread.Sleep(25); observation = readback(); // Read-only settlement; delivery is never repeated.
                }
                receipt(new { Phase = "ActionReadback", Name = name, observation.Proven, observation.State,
                    Replay = false });
                return observation.Proven;
            }
            catch (Exception error)
            {
                receipt(new { Phase = "ReadbackUnavailable", Name = name, ErrorType = error.GetType().FullName,
                    HResult = "0x" + unchecked((uint)error.HResult).ToString("X8"), Replay = false });
                return false;
            }
        }

        private static void ClickOnce(IntPtr root, IntPtr button, uint threadId, int processId)
        {
            Guard(root, button, threadId, processId);
            IntPtr ignored;
            if (SendMessageTimeout(button, BmClick, IntPtr.Zero, IntPtr.Zero, SmtoAbortIfHung, 5000, out ignored) == IntPtr.Zero)
                throw new InvalidOperationException("Bounded BM_CLICK returned no terminal send result.");
        }

        private static void DefaultActionOnce(IntPtr root, IntPtr popup, AutomationElement virtualItem,
            uint threadId, int processId)
        {
            uint pid; uint tid = GetWindowThreadProcessId(popup, out pid);
            if (popup == IntPtr.Zero || !IsWindow(popup) || !IsWindowVisible(popup) || pid != processId || tid != threadId ||
                NativeAncestor(virtualItem) != popup || virtualItem.Current.NativeWindowHandle != 0 ||
                virtualItem.Current.ProcessId != processId || !virtualItem.Current.IsEnabled ||
                virtualItem.Current.IsOffscreen || !string.Equals(virtualItem.Current.Name, GitLabel, StringComparison.Ordinal) ||
                virtualItem.Current.ControlType != ControlType.MenuItem)
                throw new InvalidOperationException("Exact owned virtual Git item changed before default action.");
            IntPtr owner = GetWindow(popup, GwOwner);
            uint ownerPid; uint ownerTid = GetWindowThreadProcessId(owner, out ownerPid);
            var popupClass = new StringBuilder(128); var ownerClass = new StringBuilder(128);
            GetClassName(popup, popupClass, popupClass.Capacity);
            GetClassName(owner, ownerClass, ownerClass.Capacity);
            int domain = popupClass.ToString().IndexOf(".app.", StringComparison.Ordinal);
            bool exactHiddenWinFormsOwner = owner != root && !IsWindowVisible(owner) &&
                GetParent(owner) == IntPtr.Zero && GetWindow(owner, GwOwner) == IntPtr.Zero &&
                (GetWindowLong(owner, -16) & 0x40000000) == 0 &&
                ownerClass.ToString().StartsWith("WindowsForms10.Window.0", StringComparison.Ordinal) &&
                domain > 0 && ownerClass.ToString().EndsWith(popupClass.ToString().Substring(domain), StringComparison.Ordinal);
            if (owner == IntPtr.Zero || !IsWindow(owner) || ownerPid != processId || ownerTid != threadId ||
                (owner != root && !exactHiddenWinFormsOwner))
                throw new InvalidOperationException("The synthetic popup owner is not an exact owned top window.");
            Guid iid = IAccessibleId; IAccessible accessibility = null;
            try
            {
                int hr = AccessibleObjectFromWindow(popup, ObjidClient, ref iid, out accessibility);
                if (hr != 0 || accessibility == null)
                    throw new InvalidOperationException("Exact popup MSAA client object unavailable: 0x" +
                        unchecked((uint)hr).ToString("X8"));
                int exactChild = 0, count = accessibility.accChildCount;
                if (count < 1 || count > 64) throw new InvalidOperationException("Bounded MSAA popup children unavailable.");
                for (int child = 1; child <= count; child++)
                    if (string.Equals(accessibility.get_accName(child), GitLabel, StringComparison.Ordinal))
                    {
                        if (exactChild != 0) throw new InvalidOperationException("Ambiguous exact MSAA Git item.");
                        exactChild = child;
                    }
                if (exactChild == 0) throw new InvalidOperationException("Exact MSAA Git item absent.");
                accessibility.accDoDefaultAction(exactChild);
            }
            finally { if (accessibility != null && Marshal.IsComObject(accessibility)) Marshal.ReleaseComObject(accessibility); }
        }

        /// <summary>Runs only owned synthetic UI actions after both threads prove the inactive desktop.</summary>
        public static void Run(string expectedDesktop, Action<object> receipt)
        {
            if (receipt == null) throw new ArgumentNullException(nameof(receipt));
            RequireInactiveDesktop(expectedDesktop);
            CanaryForm form = null; Exception uiError = null;
            using (var ready = new ManualResetEventSlim())
            {
                var ui = new Thread(() => {
                    try
                    {
                        RequireInactiveDesktop(expectedDesktop);
                        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);
                        using (var owned = new CanaryForm())
                        {
                            form = owned;
                            owned.Shown += (sender, args) => ready.Set();
                            Application.Run(owned);
                        }
                    }
                    catch (Exception error) { uiError = error; ready.Set(); }
                }) { IsBackground = true };
                ui.SetApartmentState(ApartmentState.STA); ui.Start();
                try
                {
                    if (!ready.Wait(TimeSpan.FromSeconds(10))) throw new TimeoutException("Synthetic private UI was not shown.");
                    if (uiError != null) throw new InvalidOperationException("Synthetic private UI failed before action.", uiError);
                    IntPtr rootHandle = form.Handle; uint process; uint uiThread = GetWindowThreadProcessId(rootHandle, out process);
                    Guard(rootHandle, rootHandle, uiThread, Process.GetCurrentProcess().Id);
                    if (process != Process.GetCurrentProcess().Id || !IsWindowVisible(rootHandle))
                        throw new InvalidOperationException("Synthetic root visibility or PID differs.");
                    var root = AutomationElement.FromHandle(rootHandle);
                    receipt(new { Phase = "CanaryReady", ProcessId = process, UiThreadId = uiThread,
                        RootHandle = rootHandle.ToInt64(), Desktop = expectedDesktop,
                        Actions = ActionMap });
                    RunActions(form, root, rootHandle, uiThread, (int)process, receipt);
                }
                finally
                {
                    try
                    {
                        if (form != null && !form.IsDisposed && form.IsHandleCreated)
                            form.BeginInvoke((Action)(() => { if (!form.IsDisposed) form.Close(); }));
                    }
                    catch (Exception error)
                    {
                        receipt(new { Phase = "CanaryCleanupFailed", ErrorType = error.GetType().FullName,
                            HResult = "0x" + unchecked((uint)error.HResult).ToString("X8"), ForceTermination = false });
                    }
                    if (!ui.Join(TimeSpan.FromSeconds(10)))
                        receipt(new { Phase = "CanaryUiThreadPending", Retain = true, ForceTermination = false });
                }
            }
        }

        private static void RunActions(CanaryForm form, AutomationElement root, IntPtr rootHandle,
            uint uiThread, int processId, Action<object> receipt)
        {
            Probe(receipt, "TextBox", () => {
            var text = One(root, "CanaryText", processId);
            var textPattern = Pattern<ValuePattern>(text, ValuePattern.Pattern);
            Guard(rootHandle, new IntPtr(text.Current.NativeWindowHandle), uiThread, processId);
            Once(receipt, "TextBox", "ValuePattern.SetValue", () => textPattern.SetValue("after"),
                () => new Observation(textPattern.Current.Value == "after", new { Value = textPattern.Current.Value }));
            });

            Probe(receipt, "TabControl", () => {
            var tabs = One(root, "CanaryTabs", processId);
            IntPtr tabsHandle = new IntPtr(tabs.Current.NativeWindowHandle);
            Guard(rootHandle, tabsHandle, uiThread, processId);
            var matches = tabs.FindAll(TreeScope.Children,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem))
                .Cast<AutomationElement>().Where(item => item.Current.Name == "Second" && item.Current.ProcessId == processId).ToArray();
            if (matches.Length != 1 || NativeAncestor(matches[0]) != tabsHandle)
                throw new InvalidOperationException("The exact virtual second tab is absent or ambiguous.");
            var tab = matches[0];
            Once(receipt, "TabControl", "SelectionItemPattern.Select",
                () => Pattern<SelectionItemPattern>(tab, SelectionItemPattern.Pattern).Select(),
                () => { bool selected = (bool)form.Invoke((Func<bool>)(() => form.Tabs.SelectedIndex == 1));
                    return new Observation(selected, new { Selected = selected }); });
            });

            Probe(receipt, "ComboBox", () => {
            var combo = One(root, "CanaryCombo", processId);
            Guard(rootHandle, new IntPtr(combo.Current.NativeWindowHandle), uiThread, processId);
            var expand = Pattern<ExpandCollapsePattern>(combo, ExpandCollapsePattern.Pattern);
            if (Once(receipt, "ComboBoxExpand", "ExpandCollapsePattern.Expand", expand.Expand,
                () => new Observation(expand.Current.ExpandCollapseState == ExpandCollapseState.Expanded,
                    new { State = expand.Current.ExpandCollapseState.ToString() })))
            {
                var entries = combo.FindAll(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.NameProperty, "Scope B")).Cast<AutomationElement>()
                    .Where(item => item.Current.ControlType == ControlType.ListItem && item.Current.ProcessId == processId).ToArray();
                if (entries.Length != 1) receipt(new { Phase = "ActionGap", Name = "ComboBoxSelect", Reason = "Exact list item absent or ambiguous." });
                else Once(receipt, "ComboBoxSelect", "SelectionItemPattern.Select",
                    () => Pattern<SelectionItemPattern>(entries[0], SelectionItemPattern.Pattern).Select(),
                    () => { int selected = (int)form.Invoke((Func<int>)(() => form.Combo.SelectedIndex));
                        return new Observation(selected == 1, new { SelectedIndex = selected }); });
                Once(receipt, "ComboBoxCollapse", "ExpandCollapsePattern.Collapse", expand.Collapse,
                    () => new Observation(expand.Current.ExpandCollapseState == ExpandCollapseState.Collapsed,
                        new { State = expand.Current.ExpandCollapseState.ToString() }));
            }
            else receipt(new { Phase = "ActionGap", Name = "ComboBoxSelectAndCollapse", Reason = "Expand not proved; no dependent action." });
            });

            Probe(receipt, "NativeButton", () => {
            var native = One(root, "CanaryNativeButton", processId);
            IntPtr nativeHandle = new IntPtr(native.Current.NativeWindowHandle);
            Once(receipt, "NativeButton", "SendMessageTimeout(BM_CLICK)",
                () => ClickOnce(rootHandle, nativeHandle, uiThread, processId),
                () => { int count = (int)form.Invoke((Func<int>)(() => form.NativeClicks));
                    return new Observation(count == 1, new { Clicks = count }); });
            });

            Probe(receipt, "ChatOptionsAndVirtualGit", () => {
            var options = One(root, "CanaryOptions", processId);
            IntPtr optionsHandle = new IntPtr(options.Current.NativeWindowHandle);
            bool opened = Once(receipt, "ChatActionButtonOptions", "SendMessageTimeout(BM_CLICK)",
                () => ClickOnce(rootHandle, optionsHandle, uiThread, processId),
                () => { int count = (int)form.Invoke((Func<int>)(() => form.OptionsClicks));
                    bool visible = (bool)form.Invoke((Func<bool>)(() => form.OptionsMenu.Visible));
                    return new Observation(count == 1 && visible, new { Clicks = count, PopupVisible = visible }); });
            if (opened && (bool)form.Invoke((Func<bool>)(() => form.OptionsMenu.Visible)))
            {
                IntPtr popup = (IntPtr)form.Invoke((Func<IntPtr>)(() => form.OptionsMenu.Handle));
                var popupElement = AutomationElement.FromHandle(popup);
                var items = popupElement.FindAll(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem))
                    .Cast<AutomationElement>().Where(item => item.Current.Name == GitLabel &&
                        item.Current.ProcessId == processId).ToArray();
                if (items.Length != 1 || popupElement.Current.ControlType != ControlType.Menu)
                    receipt(new { Phase = "ActionGap", Name = "VirtualGitItem", Reason = "Exact popup/menu item absent or ambiguous." });
                else Once(receipt, "VirtualGitItem", "IAccessible.accDoDefaultAction",
                    () => DefaultActionOnce(rootHandle, popup, items[0], uiThread, processId),
                    () => { int count = (int)form.Invoke((Func<int>)(() => form.GitClicks));
                        return new Observation(count == 1, new { Clicks = count }); });
            }
            else receipt(new { Phase = "ActionGap", Name = "VirtualGitItem", Reason = "Options popup not proved." });
            });

            Probe(receipt, "OwnedForm", () => {
            var window = Pattern<WindowPattern>(root, WindowPattern.Pattern);
            Once(receipt, "OwnedForm", "WindowPattern.Close", window.Close,
                () => new Observation(!IsWindow(rootHandle) && form.ClosedEvents == 1,
                    new { Destroyed = !IsWindow(rootHandle), ClosedEvents = form.ClosedEvents }));
            });
        }

        private static void Probe(Action<object> receipt, string name, Action action)
        {
            try { action(); }
            catch (Exception error)
            {
                receipt(new { Phase = "ActionPreflightGap", Name = name,
                    ErrorType = error.GetType().FullName,
                    HResult = "0x" + unchecked((uint)error.HResult).ToString("X8"),
                    ActionDelivered = false, Replay = false });
            }
        }

        private sealed class CanaryForm : Form
        {
            internal readonly TabControl Tabs = new TabControl();
            internal readonly ComboBox Combo = new ComboBox();
            internal readonly ContextMenuStrip OptionsMenu = new ContextMenuStrip();
            internal int NativeClicks, OptionsClicks, GitClicks, ClosedEvents;
            protected override bool ShowWithoutActivation => true;
            protected override CreateParams CreateParams
            {
                get { var parameters = base.CreateParams; parameters.ExStyle |= 0x08000000; return parameters; }
            }
            internal CanaryForm()
            {
                Text = "VBAi Private UI Action Canary"; Name = "CanaryRoot";
                Width = 460; Height = 340; StartPosition = FormStartPosition.Manual;
                var text = new TextBox { Name = "CanaryText", Text = "before", Left = 12, Top = 12, Width = 200 };
                Tabs.Name = "CanaryTabs"; Tabs.Left = 12; Tabs.Top = 45; Tabs.Width = 300; Tabs.Height = 80;
                Tabs.TabPages.Add(new TabPage("First") { Name = "CanaryFirstTab" });
                Tabs.TabPages.Add(new TabPage("Second") { Name = "CanarySecondTab" });
                Combo.Name = "CanaryCombo"; Combo.Left = 12; Combo.Top = 130; Combo.Width = 200;
                Combo.DropDownStyle = ComboBoxStyle.DropDownList;
                Combo.Items.AddRange(new object[] { "Scope A", "Scope B" }); Combo.SelectedIndex = 0;
                var native = new Button { Name = "CanaryNativeButton", Text = "Native button", Left = 12, Top = 170, Width = 130 };
                native.Click += (sender, args) => NativeClicks++;
                var options = new VBAi.ChatActionButton { Name = "CanaryOptions", Text = "Options", Left = 150, Top = 170, Width = 130 };
                options.Click += (sender, args) => { OptionsClicks++; OptionsMenu.Show(options, 0, options.Height); };
                var git = new ToolStripMenuItem(GitLabel) { Name = "github" };
                git.Click += (sender, args) => GitClicks++;
                OptionsMenu.Items.Add(git);
                Controls.AddRange(new Control[] { text, Tabs, Combo, native, options });
                FormClosed += (sender, args) => ClosedEvents++;
            }
        }
    }
}
