using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Automation.Text;

namespace CodexVBE
{
    // Run these accessibility calls away from the VBE UI thread. UIA and MSAA
    // query native windows, not the VBIDE object model used by VbeDebug.
    internal static class VbeDebugWindows
    {
        internal delegate bool EnumWindowCallback(IntPtr handle, IntPtr parameter);

        [DllImport("user32.dll", EntryPoint = "EnumWindows")] private static extern bool NativeEnumWindows(EnumWindowCallback callback, IntPtr parameter);
        [DllImport("user32.dll", EntryPoint = "EnumChildWindows")] private static extern bool NativeEnumChildWindows(IntPtr parent, EnumWindowCallback callback, IntPtr parameter);
        [DllImport("user32.dll", EntryPoint = "GetWindowThreadProcessId")] private static extern uint NativeGetWindowThreadProcessId(IntPtr handle, out uint processId);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "GetClassName")] private static extern int NativeGetClassName(IntPtr handle, StringBuilder text, int capacity);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "GetWindowText")] private static extern int NativeGetWindowText(IntPtr handle, StringBuilder text, int capacity);
        [DllImport("user32.dll", EntryPoint = "GetDlgCtrlID")] private static extern int NativeGetDlgCtrlID(IntPtr handle);
        [DllImport("user32.dll", EntryPoint = "GetDlgItem")] private static extern IntPtr NativeGetDlgItem(IntPtr handle, int controlId);
        [DllImport("user32.dll", EntryPoint = "IsWindowVisible")] private static extern bool NativeIsWindowVisible(IntPtr handle);
        [DllImport("user32.dll", EntryPoint = "PostMessage")] private static extern bool NativePostMessage(IntPtr handle, int message, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "SendMessageW")]
        private static extern IntPtr NativeSendMessageText(IntPtr handle, int message, IntPtr wParam, string text);
        [DllImport("user32.dll", EntryPoint = "SendMessageW")]
        private static extern IntPtr NativeSendMessageInt(IntPtr handle, int message, IntPtr wParam, IntPtr lParam);
        [DllImport("oleacc.dll", EntryPoint = "AccessibleObjectFromWindow")] private static extern int NativeAccessibleObjectFromWindow(IntPtr handle, uint objectId,
            ref Guid interfaceId, [MarshalAs(UnmanagedType.Interface)] out object accessible);

        internal delegate uint WindowProcessReader(IntPtr handle, out uint processId);
        internal delegate int AccessibleClientReader(IntPtr handle, uint objectId, ref Guid interfaceId, out object accessible);

        // These delegates isolate the operating-system boundary; dialog decisions stay in this class.
        internal static Func<EnumWindowCallback, IntPtr, bool> EnumWindows = NativeEnumWindows;
        internal static Func<IntPtr, EnumWindowCallback, IntPtr, bool> EnumChildWindows = NativeEnumChildWindows;
        internal static WindowProcessReader GetWindowThreadProcessId = NativeGetWindowThreadProcessId;
        internal static Func<IntPtr, StringBuilder, int, int> GetClassName = NativeGetClassName;
        internal static Func<IntPtr, StringBuilder, int, int> GetWindowText = NativeGetWindowText;
        internal static Func<IntPtr, int> GetDlgCtrlID = NativeGetDlgCtrlID;
        internal static Func<IntPtr, int, IntPtr> GetDlgItem = NativeGetDlgItem;
        internal static Func<IntPtr, bool> IsWindowVisible = NativeIsWindowVisible;
        internal static Func<IntPtr, int, IntPtr, IntPtr, bool> PostMessage = NativePostMessage;
        internal static Func<IntPtr, int, IntPtr, string, IntPtr> SendMessageText = NativeSendMessageText;
        internal static Func<IntPtr, int, IntPtr, IntPtr, IntPtr> SendMessageInt = NativeSendMessageInt;
        internal static AccessibleClientReader AccessibleObjectFromWindow = NativeAccessibleObjectFromWindow;
        internal static Action<int> PauseNative = Thread.Sleep;

        private const int BmClick = 0x00F5;
        private const int WmChar = 0x0102;
        private const int WmKeyDown = 0x0100;
        private const int WmKeyUp = 0x0101;
        private const int VkReturn = 13;
        private const uint ObjidClient = 4294967292;
        private static readonly Guid IidAccessible = new Guid("618736e0-3c3d-11cf-810c-00aa00389b71");

        internal sealed class NativeControl
        {
            public IntPtr Handle;
            public string Kind;
            public string Text;
            public bool Visible;
        }

        internal interface INativeProbe
        {
            IntPtr VbeRoot();
            List<IntPtr> Children(IntPtr root);
            IntPtr Pane(IEnumerable<IntPtr> panes, params string[] names);
            object List(IntPtr handle);
            object Immediate(IntPtr handle);
            object CallStack(IntPtr locals);
            IntPtr Dialog(params string[] titles);
            List<NativeControl> DialogControls(IntPtr dialog);
            string DialogMessage(IntPtr dialog);
            bool Click(IntPtr handle);
            bool Visible(IntPtr handle);
            void Pause(int milliseconds);
        }

        // Keep the Win32 boundary injectable for dialog decisions. The production
        // implementation below still calls the same native functions.
        internal interface IWatchProbe
        {
            IntPtr Dialog(params string[] titles);
            IntPtr Item(IntPtr dialog, int id);
            string Text(IntPtr handle);
            bool Click(IntPtr handle);
            bool Checked(IntPtr handle);
            void Replace(IntPtr handle, string text);
            void Pause(int milliseconds);
            string Message(IntPtr dialog);
            void Close(IntPtr dialog);
            IntPtr VbeRoot();
            List<IntPtr> Children(IntPtr root);
            IntPtr Pane(IEnumerable<IntPtr> panes, params string[] names);
            object List(IntPtr handle);
            int WatchMatches(IntPtr pane, string expression, string context);
            bool SelectWatchRow(IntPtr pane, string expression, string context);
        }

        private sealed class NativeWatchProbe : IWatchProbe
        {
            private List<AutomationElement> lastRows;
            public IntPtr Dialog(params string[] titles) { return FindDialog(titles); }
            public IntPtr Item(IntPtr dialog, int id) { return GetDlgItem(dialog, id); }
            public string Text(IntPtr handle) { return WindowText(handle); }
            public bool Click(IntPtr handle) { return PostMessage(handle, BmClick, IntPtr.Zero, IntPtr.Zero); }
            public bool Checked(IntPtr handle) { return SendMessageInt(handle, 0x00F0, IntPtr.Zero, IntPtr.Zero).ToInt32() == 1; }
            public void Replace(IntPtr handle, string value)
            {
                SendMessageInt(handle, 0x00B1, IntPtr.Zero, new IntPtr(-1));
                SendMessageText(handle, 0x00C2, new IntPtr(1), value);
            }
            public void Pause(int milliseconds) { PauseNative(milliseconds); }
            public string Message(IntPtr dialog) { return AccessibleDialogMessage(dialog); }
            public void Close(IntPtr dialog) { CloseDialog(dialog); }
            public IntPtr VbeRoot() { return FindVbeRoot(); }
            public List<IntPtr> Children(IntPtr root) { return ChildWindows(root); }
            public IntPtr Pane(IEnumerable<IntPtr> panes, params string[] names) { return FindPane(panes, names); }
            public object List(IntPtr handle) { return ReadList(handle); }
            public int WatchMatches(IntPtr pane, string expression, string context)
            {
                lastRows = MatchingWatchRows(pane, expression, context);
                return lastRows.Count;
            }
            public bool SelectWatchRow(IntPtr pane, string expression, string context)
            {
                AutomationElement row = lastRows.Single();
                if (!row.TryGetCurrentPattern(SelectionItemPattern.Pattern, out object pattern)) return false;
                ((SelectionItemPattern)pattern).Select();
                row.SetFocus();
                return true;
            }
        }

        internal sealed class SignatureChild
        {
            public int Index;
            public int Role;
            public string Name;
        }

        internal interface ISignatureProbe
        {
            IntPtr Dialog();
            IList<SignatureChild> Children(IntPtr dialog);
            void Cancel(IntPtr dialog, int index);
            void Close(IntPtr dialog);
            void Pause(int milliseconds);
        }

        private sealed class NativeSignatureProbe : ISignatureProbe
        {
            private Accessibility.IAccessible accessible;
            public IntPtr Dialog() { return FindSignatureDialog(); }
            public IList<SignatureChild> Children(IntPtr dialog)
            {
                accessible = SignatureAccessible(dialog);
                var children = new List<SignatureChild>();
                for (int index = 1; index <= Math.Min(accessible.accChildCount, 64); index++)
                {
                    try
                    {
                        children.Add(new SignatureChild { Index = index, Name = accessible.get_accName(index),
                            Role = Convert.ToInt32(accessible.get_accRole(index)) });
                    }
                    catch { /* A single inaccessible MSAA child is skipped. */ }
                }
                return children;
            }
            public void Cancel(IntPtr dialog, int index) { accessible.accDoDefaultAction(index); }
            public void Close(IntPtr dialog) { PostMessage(dialog, 0x0010, IntPtr.Zero, IntPtr.Zero); }
            public void Pause(int milliseconds) { PauseNative(milliseconds); }
        }

        internal sealed class OptionsControl
        {
            public string Name;
            public string Type;
            public object Value;
            public string Error;
            public bool Visible = true;
            public bool Enabled = true;
        }

        internal sealed class OptionsChoice
        {
            public string Name;
            public bool Selected;
            public bool Readable = true;
        }

        internal interface IOptionsProbe
        {
            IntPtr Dialog();
            IList<string> Tabs(IntPtr dialog);
            IList<OptionsControl> Controls(IntPtr dialog, int tabIndex);
            IList<OptionsChoice> ErrorChoices(IntPtr dialog);
            void Close(IntPtr dialog);
            void Pause(int milliseconds);
        }

        private sealed class NativeOptionsProbe : IOptionsProbe
        {
            private AutomationElement root;
            private AutomationElementCollection tabItems;
            public IntPtr Dialog() { return FindDialog("Options"); }
            public IList<string> Tabs(IntPtr dialog)
            {
                root = AutomationElement.FromHandle(dialog);
                tabItems = root.FindAll(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem));
                return tabItems.Cast<AutomationElement>().Select(tab => tab.Current.Name).ToArray();
            }
            public IList<OptionsControl> Controls(IntPtr dialog, int tabIndex)
            {
                object tabPattern;
                if (!tabItems[tabIndex].TryGetCurrentPattern(SelectionItemPattern.Pattern, out tabPattern))
                    throw new InvalidOperationException("A native VBE Options tab is unreadable.");
                ((SelectionItemPattern)tabPattern).Select();
                PauseNative(75);
                var descendants = root.FindAll(TreeScope.Descendants,
                    new OrCondition(
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.CheckBox),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.RadioButton),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ComboBox),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.List),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Slider),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text)));
                if (descendants.Count > 2000)
                    throw new InvalidOperationException("The Options dialog has too many controls to inspect safely: " +
                        descendants.Count + ".");
                var controls = new List<OptionsControl>();
                for (int index = 0; index < descendants.Count; index++)
                {
                    AutomationElement element = descendants[index];
                    try
                    {
                        ControlType kind = element.Current.ControlType;
                        var control = new OptionsControl { Name = element.Current.Name,
                            Type = kind.ProgrammaticName, Visible = !element.Current.IsOffscreen,
                            Enabled = element.Current.IsEnabled };
                        if (!control.Visible || !control.Enabled) { controls.Add(control); continue; }
                        try
                        {
                            if (kind == ControlType.CheckBox &&
                                element.TryGetCurrentPattern(TogglePattern.Pattern, out object toggle))
                                control.Value = ((TogglePattern)toggle).Current.ToggleState.ToString();
                            else if ((kind == ControlType.RadioButton || kind == ControlType.ListItem) &&
                                element.TryGetCurrentPattern(SelectionItemPattern.Pattern, out object selection))
                                control.Value = ((SelectionItemPattern)selection).Current.IsSelected;
                            else if ((kind == ControlType.Edit || kind == ControlType.ComboBox) &&
                                !element.Current.IsPassword &&
                                element.TryGetCurrentPattern(ValuePattern.Pattern, out object input))
                                control.Value = ((ValuePattern)input).Current.Value;
                            else if (kind == ControlType.Slider &&
                                element.TryGetCurrentPattern(RangeValuePattern.Pattern, out object slider))
                                control.Value = ((RangeValuePattern)slider).Current.Value;
                        }
                        catch (Exception ex) { control.Error = ex.Message; }
                        controls.Add(control);
                    }
                    catch (ElementNotAvailableException) { }
                }
                return controls;
            }
            public IList<OptionsChoice> ErrorChoices(IntPtr dialog)
            {
                AutomationElement options = AutomationElement.FromHandle(dialog);
                var generalCondition = new AndCondition(
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem),
                    new OrCondition(
                        new PropertyCondition(AutomationElement.NameProperty, "Général"),
                        new PropertyCondition(AutomationElement.NameProperty, "General")));
                AutomationElementCollection tabs = options.FindAll(TreeScope.Descendants, generalCondition);
                if (tabs.Count != 1 || !tabs[0].TryGetCurrentPattern(SelectionItemPattern.Pattern, out object tabPattern))
                    throw new InvalidOperationException("The native General options tab is unavailable.");
                ((SelectionItemPattern)tabPattern).Select();
                var radios = options.FindAll(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.RadioButton));
                var result = new List<OptionsChoice>();
                for (int index = 0; index < radios.Count; index++)
                {
                    AutomationElement radio = radios[index];
                    var choice = new OptionsChoice { Name = radio.Current.Name };
                    choice.Readable = radio.TryGetCurrentPattern(SelectionItemPattern.Pattern, out object pattern);
                    if (choice.Readable) choice.Selected = ((SelectionItemPattern)pattern).Current.IsSelected;
                    result.Add(choice);
                }
                return result;
            }
            public void Close(IntPtr dialog) { CloseDialog(dialog); }
            public void Pause(int milliseconds) { PauseNative(milliseconds); }
        }

        internal interface IImmediateProbe
        {
            IntPtr VbeRoot();
            List<IntPtr> Children(IntPtr root);
            IntPtr Pane(IEnumerable<IntPtr> panes, params string[] names);
            string Prepare(IntPtr pane);
            string Text(IntPtr pane);
            bool PostChar(IntPtr pane, char character);
            bool PostEnter(IntPtr pane);
            void Pause(int milliseconds);
        }

        private sealed class NativeImmediateProbe : IImmediateProbe
        {
            public IntPtr VbeRoot() { return FindVbeRoot(); }
            public List<IntPtr> Children(IntPtr root) { return ChildWindows(root); }
            public IntPtr Pane(IEnumerable<IntPtr> panes, params string[] names) { return FindPane(panes, names); }
            public string Prepare(IntPtr pane)
            {
                AutomationElement document = ImmediateDocument(pane);
                var pattern = (TextPattern)document.GetCurrentPattern(TextPattern.Pattern);
                string before = pattern.DocumentRange.GetText(-1);
                document.SetFocus();
                TextPatternRange atEnd = pattern.DocumentRange.Clone();
                atEnd.MoveEndpointByRange(TextPatternRangeEndpoint.Start, pattern.DocumentRange,
                    TextPatternRangeEndpoint.End);
                atEnd.Select();
                return before;
            }
            public string Text(IntPtr pane) { return ImmediateText(pane); }
            public bool PostChar(IntPtr pane, char character)
            { return PostMessage(pane, WmChar, new IntPtr(character), IntPtr.Zero); }
            public bool PostEnter(IntPtr pane)
            {
                return PostMessage(pane, WmKeyDown, new IntPtr(VkReturn), IntPtr.Zero) &&
                    PostMessage(pane, WmKeyUp, new IntPtr(VkReturn), IntPtr.Zero);
            }
            public void Pause(int milliseconds) { PauseNative(milliseconds); }
        }

        private sealed class NativeProbe : INativeProbe
        {
            public IntPtr VbeRoot() { return FindVbeRoot(); }
            public List<IntPtr> Children(IntPtr root) { return ChildWindows(root); }
            public IntPtr Pane(IEnumerable<IntPtr> panes, params string[] names) { return FindPane(panes, names); }
            public object List(IntPtr handle) { return ReadList(handle); }
            public object Immediate(IntPtr handle) { return ReadImmediate(handle); }
            public object CallStack(IntPtr locals) { return ReadCallStack(locals); }
            public IntPtr Dialog(params string[] titles) { return FindDialog(titles); }
            public List<NativeControl> DialogControls(IntPtr dialog)
            {
                var controls = new List<NativeControl>();
                EnumChildWindows(dialog, (handle, parameter) => {
                    controls.Add(new NativeControl { Handle = handle, Kind = ClassName(handle),
                        Text = WindowText(handle), Visible = IsWindowVisible(handle) });
                    return true;
                }, IntPtr.Zero);
                return controls;
            }
            public string DialogMessage(IntPtr dialog) { return AccessibleDialogMessage(dialog); }
            public bool Click(IntPtr handle) { return PostMessage(handle, BmClick, IntPtr.Zero, IntPtr.Zero); }
            public bool Visible(IntPtr handle) { return IsWindowVisible(handle); }
            public void Pause(int milliseconds) { PauseNative(milliseconds); }
        }

        public static object Capture(bool includeCallStack)
        {
            return Capture(includeCallStack, new NativeProbe());
        }

        internal static object Capture(bool includeCallStack, INativeProbe native)
        {
            if (native == null) throw new ArgumentNullException(nameof(native));
            IntPtr root = native.VbeRoot();
            if (root == IntPtr.Zero) throw new InvalidOperationException("The VBE window is not open in this host process.");
            var panes = native.Children(root);
            IntPtr locals = native.Pane(panes, "Variables locales", "Locals");
            IntPtr watches = native.Pane(panes, "Espions", "Watch", "Watches");
            IntPtr immediate = native.Pane(panes, "Exécution", "Immediate");
            object stack = includeCallStack ? native.CallStack(locals) : null;
            return new {
                HostProcessId = Process.GetCurrentProcess().Id,
                Locals = native.List(locals),
                Watches = native.List(watches),
                Immediate = native.Immediate(immediate),
                CallStack = stack,
                Limits = "Native UI accessibility is observed only for visible panes. A missing pane or unavailable value is not an empty debugger collection. Breakpoints and exception state are not exposed by this snapshot."
            };
        }

        // The Compile command can open a modal native diagnostic. Its UI-thread
        // Execute call cannot be awaited with Control.Invoke in that case.
        public static void EnsureNoCompileDialog()
        {
            EnsureNoCompileDialog(new NativeProbe());
        }

        internal static void EnsureNoCompileDialog(INativeProbe native)
        {
            if (native.Dialog("Microsoft Visual Basic pour Applications",
                "Microsoft Visual Basic for Applications") != IntPtr.Zero)
                throw new InvalidOperationException("A native VBE dialog is already open; compilation was not started.");
        }

        public static object ReadDebugDialog()
        {
            return ReadDebugDialog(new NativeProbe());
        }

        internal static object ReadDebugDialog(INativeProbe native)
        {
            IntPtr dialog = native.Dialog("Microsoft Visual Basic pour Applications",
                "Microsoft Visual Basic for Applications");
            if (dialog == IntPtr.Zero)
                return new { Available = false, Diagnostic = (string)null,
                    Buttons = new string[0], Error = (string)null };
            var messages = new List<string>();
            var buttons = new List<string>();
            foreach (var control in native.DialogControls(dialog))
            {
                if (!control.Visible || string.IsNullOrWhiteSpace(control.Text)) continue;
                string title = control.Text;
                string kind = control.Kind;
                if (kind == "Static") messages.Add(title);
                else if (kind == "Button") buttons.Add(title);
            }
            if (messages.Count != 1)
                return new { Available = true, Diagnostic = (string)null,
                    Buttons = buttons.ToArray(), Error = "Expected one native diagnostic message; found " + messages.Count + "." };
            return new { Available = true, Diagnostic = messages[0],
                Buttons = buttons.ToArray(), Error = (string)null };
        }

        public static void EnsureNoDebugOptionsDialog()
        {
            EnsureNoDebugOptionsDialog(new NativeProbe());
        }

        internal static void EnsureNoDebugOptionsDialog(INativeProbe native)
        {
            if (native.Dialog("Options") != IntPtr.Zero)
                throw new InvalidOperationException("A VBE Options dialog is already open; the add-in will not close a user-owned dialog.");
        }

        public static void EnsureNoSignatureDialog()
        {
            if (FindSignatureDialog() != IntPtr.Zero)
                throw new InvalidOperationException("A VBE Digital Signature dialog is already open; the add-in will not close a user-owned dialog.");
        }

        public static object ReadSignatureDialog(string project)
        {
            return ReadSignatureDialog(project, new NativeSignatureProbe());
        }

        internal static object ReadSignatureDialog(string project, ISignatureProbe native)
        {
            IntPtr dialog = IntPtr.Zero;
            for (int attempt = 0; attempt < 60 && dialog == IntPtr.Zero; attempt++)
            { native.Pause(50); dialog = native.Dialog(); }
            if (dialog == IntPtr.Zero)
                throw new InvalidOperationException("The native VBE Digital Signature dialog did not open.");
            var labels = new List<string>();
            var buttons = new List<string>();
            bool cancelled = false;
            try
            {
                int cancelIndex = 0;
                foreach (SignatureChild child in native.Children(dialog))
                {
                    string name = child.Name;
                    int role = child.Role;
                    if (string.IsNullOrWhiteSpace(name)) continue;
                    if (role == 43)
                    {
                        buttons.Add(name);
                        if (string.Equals(name, "Annuler", StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(name, "Cancel", StringComparison.OrdinalIgnoreCase))
                            cancelIndex = child.Index;
                    }
                    else if (role == 41) labels.Add(name);
                }
                if (cancelIndex == 0)
                    throw new InvalidOperationException("The native Digital Signature dialog has no accessible Cancel button.");
                native.Cancel(dialog, cancelIndex);
                cancelled = true;
            }
            finally
            {
                if (!cancelled && native.Dialog() == dialog)
                    native.Close(dialog); // WM_CLOSE on our exact dialog
            }
            bool closed = false;
            for (int attempt = 0; attempt < 20; attempt++)
            {
                if (native.Dialog() == IntPtr.Zero) { closed = true; break; }
                native.Pause(50);
            }
            if (!closed) throw new InvalidOperationException("The signature dialog was read but did not close.");
            string currentCertificate = CertificateBeforeHeading(labels,
                "Signature actuelle du projet VBA", "The VBA project is currently signed as");
            string signAsCertificate = CertificateBeforeHeading(labels, "Signer en tant que", "Sign as");
            return new { Project = project, CurrentCertificate = currentCertificate,
                SignAsCertificate = signAsCertificate, Labels = labels.ToArray(), Buttons = buttons.ToArray(),
                DialogClosed = true, Verification = "NativeSignatureDialogReadback",
                Limit = "Labels reflect the native dialog; no certificate was selected, assigned, removed or cryptographically validated." };
        }

        // Windows' protected certificate picker does not expose its buttons to
        // UIA/MSAA. The user confirms the named certificate there; the add-in
        // checks VBE's readback and completes only its own native VBE dialog.
        public static object CompleteProjectSignature(string project, string thumbprint, string certificateName,
            bool unsignedVerified)
        {
            IntPtr dialog = IntPtr.Zero;
            for (int attempt = 0; attempt < 60 && dialog == IntPtr.Zero; attempt++)
            { PauseNative(50); dialog = FindSignatureDialog(); }
            if (dialog == IntPtr.Zero)
                throw new InvalidOperationException("The native VBE Digital Signature dialog did not open.");
            bool completed = false;
            try
            {
                Accessibility.IAccessible root = SignatureAccessible(dialog);
                var initial = SignatureLabels(root);
                string current = CertificateBeforeHeading(initial,
                    "Signature actuelle du projet VBA", "The VBA project is currently signed as");
                string signAs = CertificateBeforeHeading(initial, "Signer en tant que", "Sign as");
                if ((!unsignedVerified && !IsNoCertificate(current)) ||
                    (!IsNoCertificate(signAs) && !string.Equals(signAs, certificateName, StringComparison.OrdinalIgnoreCase)))
                    throw new InvalidOperationException("This command adds the first signature only; the project already has a certificate selected.");
                bool chosen = string.Equals(signAs, certificateName, StringComparison.OrdinalIgnoreCase) ||
                    (unsignedVerified && IsNoCertificate(signAs) &&
                     string.Equals(current, certificateName, StringComparison.OrdinalIgnoreCase));
                bool pickerOpened = !chosen;
                if (!chosen) InvokeSignatureButton(root, "Choisir...", "Choose...");
                for (int attempt = 0; !chosen && attempt < 2400; attempt++)
                {
                    PauseNative(50);
                    if (FindSignatureDialog() != dialog)
                        throw new InvalidOperationException("The native VBE signature dialog closed during certificate selection.");
                    if (attempt % 4 != 0) continue;
                    string selected = CertificateBeforeHeading(SignatureLabels(root), "Signer en tant que", "Sign as");
                    if (string.Equals(selected, certificateName, StringComparison.OrdinalIgnoreCase))
                    { chosen = true; break; }
                    if (!IsNoCertificate(selected))
                        throw new InvalidOperationException("A different certificate was selected; the signature was cancelled.");
                }
                if (!chosen)
                    throw new TimeoutException("The certificate was not confirmed in Windows Security within two minutes.");
                InvokeSignatureButton(root, "OK");
                for (int attempt = 0; attempt < 60; attempt++)
                {
                    if (FindSignatureDialog() == IntPtr.Zero) { completed = true; break; }
                    PauseNative(50);
                }
                if (!completed) throw new InvalidOperationException("The native VBE signature dialog did not close after OK.");
                return new { Project = project, CertificateThumbprint = thumbprint,
                    CertificateName = certificateName, SignatureAssigned = true,
                    Verification = "NativeVbeCertificateReadbackAndDialogClose",
                    SelectionSource = pickerOpened ? "WindowsCertificatePicker" : "NativeVbeExistingCertificate",
                    CertificatePicker = pickerOpened
                        ? "The Windows certificate picker returned the named certificate; the add-in did not interact with its protected controls."
                        : "The named certificate was already associated with the project in the native VBE dialog." };
            }
            finally
            {
                if (!completed && FindSignatureDialog() == dialog)
                    PostMessage(dialog, 0x0010, IntPtr.Zero, IntPtr.Zero);
            }
        }

        private static Accessibility.IAccessible SignatureAccessible(IntPtr dialog)
        {
            object accessible;
            Guid iid = IidAccessible;
            int hr = AccessibleObjectFromWindow(dialog, ObjidClient, ref iid, out accessible);
            if (hr != 0 || !(accessible is Accessibility.IAccessible))
                throw new COMException("The native signature dialog is not accessible through MSAA.", hr);
            return (Accessibility.IAccessible)accessible;
        }

        private static List<string> SignatureLabels(Accessibility.IAccessible root)
        {
            var labels = new List<string>();
            for (int index = 1; index <= Math.Min(root.accChildCount, 64); index++)
            {
                try
                {
                    if (Convert.ToInt32(root.get_accRole(index)) == 41)
                    {
                        string name = root.get_accName(index);
                        if (!string.IsNullOrWhiteSpace(name)) labels.Add(name);
                    }
                }
                catch (COMException) { }
            }
            return labels;
        }

        private static bool IsNoCertificate(string name)
        {
            return name != null && (name.Equals("[Aucun certificat]", StringComparison.OrdinalIgnoreCase) ||
                name.Equals("[No certificate]", StringComparison.OrdinalIgnoreCase) ||
                name.Equals("[None]", StringComparison.OrdinalIgnoreCase));
        }

        private static void InvokeSignatureButton(Accessibility.IAccessible root, params string[] names)
        {
            for (int index = 1; index <= Math.Min(root.accChildCount, 64); index++)
            {
                try
                {
                    if (Convert.ToInt32(root.get_accRole(index)) == 43 &&
                        names.Any(name => string.Equals(root.get_accName(index), name, StringComparison.OrdinalIgnoreCase)))
                    { root.accDoDefaultAction(index); return; }
                }
                catch (COMException) { }
            }
            throw new InvalidOperationException("The expected accessible signature button is unavailable: " + names[0]);
        }

        private static string CertificateBeforeHeading(IList<string> labels, params string[] headings)
        {
            for (int index = 2; index < labels.Count; index++)
                if (headings.Any(heading => string.Equals(labels[index], heading, StringComparison.OrdinalIgnoreCase)) &&
                    (labels[index - 1].StartsWith("Nom du certificat", StringComparison.OrdinalIgnoreCase) ||
                     labels[index - 1].StartsWith("Certificate name", StringComparison.OrdinalIgnoreCase)))
                    return labels[index - 2];
            return null;
        }

        private static IntPtr FindSignatureDialog()
        {
            IntPtr result = IntPtr.Zero;
            uint currentPid = (uint)Process.GetCurrentProcess().Id;
            EnumWindows((handle, parameter) => {
                uint pid;
                GetWindowThreadProcessId(handle, out pid);
                if (pid != currentPid || !IsWindowVisible(handle)) return true;
                string title = WindowText(handle);
                if (!string.Equals(title, "Signature numérique", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(title, "Digital Signature", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(title, "Signature électronique", StringComparison.OrdinalIgnoreCase)) return true;
                string kind = ClassName(handle);
                if (kind != "#32770" && !kind.StartsWith("bosa_sdm_", StringComparison.OrdinalIgnoreCase)) return true;
                result = handle;
                return false;
            }, IntPtr.Zero);
            return result;
        }

        public static object ReadDebugOptions()
        {
            return ReadDebugOptions(new NativeOptionsProbe());
        }

        internal static object ReadDebugOptions(IOptionsProbe native)
        {
            IntPtr dialog = IntPtr.Zero;
            for (int attempt = 0; attempt < 60 && dialog == IntPtr.Zero; attempt++)
            { native.Pause(50); dialog = native.Dialog(); }
            if (dialog == IntPtr.Zero) throw new InvalidOperationException("The native VBE Options dialog did not open.");
            string selected = null;
            string[] choices = null;
            try
            {
                IList<OptionsChoice> radios = native.ErrorChoices(dialog);
                if (radios.Count != 3)
                    throw new InvalidOperationException("Expected three native error trapping choices; found " + radios.Count + ".");
                var names = new List<string>();
                foreach (OptionsChoice radio in radios)
                {
                    string name = radio.Name;
                    if (string.IsNullOrWhiteSpace(name) ||
                        !(name.StartsWith("Arrêt ", StringComparison.OrdinalIgnoreCase) ||
                          name.StartsWith("Break ", StringComparison.OrdinalIgnoreCase)) ||
                        !radio.Readable)
                        throw new InvalidOperationException("A native error trapping radio is unreadable.");
                    names.Add(name);
                    if (radio.Selected)
                    {
                        if (selected != null) throw new InvalidOperationException("Multiple error trapping choices appear selected.");
                        selected = name;
                    }
                }
                if (selected == null) throw new InvalidOperationException("No error trapping choice appears selected.");
                choices = names.ToArray();
            }
            finally { native.Close(dialog); }
            bool closed = false;
            for (int attempt = 0; attempt < 20; attempt++)
            {
                if (native.Dialog() == IntPtr.Zero) { closed = true; break; }
                native.Pause(50);
            }
            if (!closed) throw new InvalidOperationException("The add-in read VBE Options but could not close its dialog.");
            return new { Scope = "VBE", ErrorTrapping = selected, Choices = choices,
                Verification = "NativeOptionsReadback", DialogClosed = true,
                Limit = "This is the currently displayed VBE-wide preference, not a diagnosis of an active runtime error." };
        }
        public static object ReadVbeOptions()
        {
            return ReadVbeOptions(new NativeOptionsProbe());
        }

        internal static object ReadVbeOptions(IOptionsProbe native)
        {
            IntPtr dialog = IntPtr.Zero;
            for (int attempt = 0; attempt < 60 && dialog == IntPtr.Zero; attempt++)
            { native.Pause(50); dialog = native.Dialog(); }
            if (dialog == IntPtr.Zero)
                throw new InvalidOperationException("The native VBE Options dialog did not open.");
            var tabs = new List<object>();
            try
            {
                IList<string> tabNames = native.Tabs(dialog);
                if (tabNames.Count < 1 || tabNames.Count > 8)
                    throw new InvalidOperationException("Unexpected native VBE Options tab count: " + tabNames.Count + ".");
                for (int tabIndex = 0; tabIndex < tabNames.Count; tabIndex++)
                {
                    string tabName = tabNames[tabIndex];
                    if (string.IsNullOrWhiteSpace(tabName))
                        throw new InvalidOperationException("A native VBE Options tab is unreadable.");
                    IList<OptionsControl> observed = native.Controls(dialog, tabIndex);
                    if (observed.Count > 2000)
                        throw new InvalidOperationException("The Options dialog has too many controls to inspect safely: " +
                            observed.Count + ".");
                    var controls = new List<object>();
                    foreach (OptionsControl control in observed)
                    {
                        if (!control.Visible || !control.Enabled) continue;
                        if (string.IsNullOrWhiteSpace(control.Name) && control.Type == "ControlType.Text") continue;
                        controls.Add(new { control.Name, control.Type, control.Value, control.Error });
                    }
                    tabs.Add(new { Tab = tabName, Controls = controls, Count = controls.Count });
                }
            }
            finally { native.Close(dialog); }
            bool closed = false;
            for (int attempt = 0; attempt < 20; attempt++)
            {
                if (native.Dialog() == IntPtr.Zero) { closed = true; break; }
                native.Pause(50);
            }
            if (!closed) throw new InvalidOperationException("The add-in read VBE Options but could not close its dialog.");
            return new { Scope = "VBE", Tabs = tabs, Count = tabs.Count,
                DialogClosed = true, Verification = "NativeOptionsReadback",
                Limit = "Only visible native controls were observed; no settings were changed." };
        }
        public static object ExecuteImmediate(string command)
        {
            return ExecuteImmediate(command, new NativeImmediateProbe());
        }

        internal static object ExecuteImmediate(string command, IImmediateProbe native)
        {
            if (string.IsNullOrWhiteSpace(command) || command.Length > 2048 ||
                command.IndexOfAny(new[] { '\r', '\n', '\0' }) >= 0 ||
                command.Any(character => char.IsControl(character)))
                throw new ArgumentException("Immediate command must be one nonempty printable line of at most 2048 characters.");
            IntPtr root = native.VbeRoot();
            if (root == IntPtr.Zero) throw new InvalidOperationException("The VBE window is not open.");
            IntPtr pane = native.Pane(native.Children(root), "Exécution", "Immediate");
            if (pane == IntPtr.Zero) throw new InvalidOperationException("The Immediate window must be visible.");
            string before = native.Prepare(pane);
            foreach (char character in command)
                if (!native.PostChar(pane, character))
                    throw new InvalidOperationException("The native Immediate pane rejected a character message.");
            string echoed = null;
            for (int attempt = 0; attempt < 40; attempt++)
            {
                native.Pause(25);
                echoed = native.Text(pane);
                if (echoed == before + command || echoed == before + command + "\r\n") break;
            }
            if (echoed != before + command && echoed != before + command + "\r\n")
                throw new InvalidOperationException("The Immediate pane did not echo the exact command; Enter was not sent.");
            if (!native.PostEnter(pane))
                throw new InvalidOperationException("The native Immediate pane rejected Enter.");
            string after = echoed;
            for (int attempt = 0; attempt < 40; attempt++)
            {
                native.Pause(25);
                after = native.Text(pane);
                if (after != echoed) break;
            }
            return new { Command = command, TextBefore = before, TextAfter = after,
                OutputDelta = after.StartsWith(before, StringComparison.Ordinal)
                    ? after.Substring(before.Length) : (string)null,
                CommandEchoObserved = true,
                Verification = after != echoed ? "ImmediateTextChangedAfterEnter" : "Pending",
                VerificationPending = after == echoed,
                Limit = "Text changed after Enter, but this does not prove an arbitrary VBA statement had the intended side effect. Read debug_state and relevant values separately." };
        }

        public static object ChangeDebugItem(Request request)
        {
            if (request == null || (request.Pane != "locals" && request.Pane != "watches") ||
                (request.Action != "expand" && request.Action != "collapse") ||
                request.PathSegments == null || request.PathSegments.Length == 0 ||
                request.PathSegments.Length > 16 || request.PathSegments.Any(string.IsNullOrWhiteSpace))
                throw new ArgumentException("Pane, Action and nonempty PathSegments are required.");
            IntPtr root = FindVbeRoot();
            if (root == IntPtr.Zero) throw new InvalidOperationException("The VBE window is not open.");
            IntPtr pane = request.Pane == "locals" ?
                FindPane(ChildWindows(root), "Variables locales", "Locals") :
                FindPane(ChildWindows(root), "Espions", "Watch", "Watches");
            if (pane == IntPtr.Zero) throw new InvalidOperationException("The requested debug pane is not visible.");
            AutomationElement scope = AutomationElement.FromHandle(pane);
            var condition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem);
            AutomationElementCollection rows = scope.FindAll(TreeScope.Descendants, condition);
            AutomationElement target = null;
            int matches = 0;
            for (int index = 0; index < rows.Count; index++)
            {
                if (!ItemPath(rows[index]).SequenceEqual(request.PathSegments, StringComparer.Ordinal)) continue;
                if (request.Pane == "watches" && !string.IsNullOrWhiteSpace(request.Context) &&
                    !string.Equals(WatchRootContext(rows[index]), request.Context, StringComparison.OrdinalIgnoreCase))
                    continue;
                target = rows[index]; matches++;
            }
            if (matches != 1)
                throw new InvalidOperationException("Expected one matching debug item; found " + matches + ".");
            object rawPattern;
            if (!target.TryGetCurrentPattern(ExpandCollapsePattern.Pattern, out rawPattern))
                throw new InvalidOperationException("The selected debug item cannot be expanded or collapsed.");
            var pattern = (ExpandCollapsePattern)rawPattern;
            if (request.Action == "expand") pattern.Expand();
            else pattern.Collapse();
            PauseNative(50);
            int childCount = target.FindAll(TreeScope.Children, condition).Count;
            bool verified = request.Action == "expand" ? childCount > 0 : childCount == 0;
            return new { request.Pane, request.Action, request.PathSegments,
                ChildCount = childCount, Verification = verified ? "Observed" : "Pending",
                VerificationPending = !verified,
                CountLimit = "ChildCount counts UIA-exposed children at this moment; long native trees may expose only a subset.",
                NextRead = "Call debug_windows to read the current hierarchical rows." };
        }

        public static object RespondDebugDialog(Request request)
        {
            return RespondDebugDialog(request, new NativeProbe());
        }

        internal static object RespondDebugDialog(Request request, INativeProbe native)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Diagnostic) ||
                string.IsNullOrWhiteSpace(request.Button))
                throw new ArgumentException("Diagnostic and Button are required.");
            IntPtr dialog = native.Dialog("Microsoft Visual Basic pour Applications",
                "Microsoft Visual Basic for Applications");
            if (dialog == IntPtr.Zero) throw new InvalidOperationException("No native VBE dialog is visible.");
            IntPtr target = IntPtr.Zero;
            string message = null;
            int matches = 0;
            foreach (var control in native.DialogControls(dialog))
            {
                if (!control.Visible) continue;
                string kind = control.Kind;
                string title = control.Text;
                if (kind == "Static" && !string.IsNullOrWhiteSpace(title)) message = title;
                if (kind == "Button" && string.Equals(title, request.Button, StringComparison.Ordinal))
                { target = control.Handle; matches++; }
            }
            if (!string.Equals(message, request.Diagnostic, StringComparison.Ordinal))
                throw new InvalidOperationException("The native diagnostic changed before the button could be activated.");
            if (!IsRecognizedDiagnostic(message))
                throw new InvalidOperationException("The visible VBE dialog is not a recognized VBA diagnostic.");
            if (matches != 1 || target == IntPtr.Zero)
                throw new InvalidOperationException("Expected exactly one matching native dialog button.");
            if (!native.Click(target))
                throw new InvalidOperationException("The native dialog button could not be activated.");
            for (int attempt = 0; attempt < 40; attempt++)
            {
                native.Pause(50);
                if (!native.Visible(dialog))
                    return new { Activated = true, request.Button, Diagnostic = message,
                        Verification = "DialogClosed", VerificationPending = false };
            }
            return new { Activated = true, request.Button, Diagnostic = message,
                Verification = "Pending", VerificationPending = true };
        }

        public static string AwaitCompileDialog(ManualResetEventSlim completed)
        {
            return AwaitCompileDialog(completed, new NativeProbe());
        }

        internal static string AwaitCompileDialog(ManualResetEventSlim completed, INativeProbe native)
        {
            for (int attempt = 0; attempt < 200; attempt++)
            {
                IntPtr dialog = native.Dialog("Microsoft Visual Basic pour Applications",
                    "Microsoft Visual Basic for Applications");
                if (dialog != IntPtr.Zero)
                {
                    string diagnostic = native.DialogMessage(dialog);
                    IntPtr ok = native.DialogControls(dialog)
                        .Where(control => control.Kind == "Button" &&
                            string.Equals(control.Text, "OK", StringComparison.OrdinalIgnoreCase))
                        .Select(control => control.Handle).FirstOrDefault();
                    if (ok == IntPtr.Zero || !native.Click(ok))
                        throw new InvalidOperationException("The native compile diagnostic could not be dismissed.");
                    if (!completed.Wait(3000))
                        throw new TimeoutException("The Compile command did not return after its diagnostic closed.");
                    return diagnostic;
                }
                if (completed.IsSet)
                {
                    // Allow the VBE to surface a delayed diagnostic after Execute.
                    native.Pause(250);
                    dialog = native.Dialog("Microsoft Visual Basic pour Applications",
                        "Microsoft Visual Basic for Applications");
                    if (dialog == IntPtr.Zero) return null;
                }
                native.Pause(50);
            }
            throw new TimeoutException("The native Compile command did not finish within ten seconds.");
        }

        internal static bool IsRecognizedDiagnostic(string message)
        {
            return message != null && Regex.IsMatch(message,
                @"^(Erreur d'exécution|Run-time error|Erreur de compilation|Compile error)",
                RegexOptions.IgnoreCase);
        }

        public static object CompleteAddWatch(Request request)
        {
            return CompleteAddWatch(request, new NativeWatchProbe());
        }

        internal static object CompleteAddWatch(Request request, IWatchProbe native)
        {
            IntPtr dialog = IntPtr.Zero;
            for (int attempt = 0; attempt < 60 && dialog == IntPtr.Zero; attempt++)
            { native.Pause(50); dialog = native.Dialog("Ajouter un espion", "Add Watch"); }
            if (dialog == IntPtr.Zero) throw new InvalidOperationException("The Add Watch dialog did not open.");
            bool completed = false;
            try
            {
                IntPtr edit = native.Item(dialog, 4853);
                IntPtr project = native.Item(dialog, 4858);
                IntPtr module = native.Item(dialog, 4857);
                IntPtr procedure = native.Item(dialog, 4856);
                IntPtr ok = native.Item(dialog, 1);
                if (edit == IntPtr.Zero || project == IntPtr.Zero || module == IntPtr.Zero || ok == IntPtr.Zero)
                    throw new InvalidOperationException("The native Add Watch dialog controls changed.");
                string shownProject = native.Text(project);
                string shownModule = native.Text(module);
                string shownProcedure = procedure == IntPtr.Zero ? null : native.Text(procedure);
                if (!string.Equals(shownProject, request.Project, StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(shownModule, request.Module, StringComparison.OrdinalIgnoreCase) ||
                    (!string.IsNullOrWhiteSpace(request.Procedure) &&
                     !string.Equals(shownProcedure, request.Procedure, StringComparison.OrdinalIgnoreCase)))
                    throw new InvalidOperationException("Add Watch context changed: " + shownProject + "." + shownModule + "." + shownProcedure);
                int typeId = request.WatchType == "break_when_true" ? 4851 :
                    request.WatchType == "break_when_changed" ? 4852 : 4850;
                IntPtr typeButton = native.Item(dialog, typeId);
                if (typeButton == IntPtr.Zero || !native.Click(typeButton))
                    throw new InvalidOperationException("The native watch type option was unavailable.");
                native.Pause(30);
                if (!native.Checked(typeButton))
                    throw new InvalidOperationException("The native watch type option was not selected.");
                // EM_REPLACESEL triggers the VBE's edit notifications. WM_SETTEXT
                // alone changes the visible text but is rejected as an empty expression.
                native.Replace(edit, request.Expression);
                if (!string.Equals(native.Text(edit), request.Expression, StringComparison.Ordinal))
                    throw new InvalidOperationException("The watch expression was not reflected by the native edit control.");
                if (!native.Click(ok))
                    throw new InvalidOperationException("The Add Watch dialog refused its OK command.");
                IntPtr error = IntPtr.Zero;
                for (int attempt = 0; attempt < 40; attempt++)
                {
                    native.Pause(50);
                    dialog = native.Dialog("Ajouter un espion", "Add Watch");
                    if (dialog == IntPtr.Zero) { completed = true; break; }
                    error = native.Dialog("Microsoft Visual Basic pour Applications", "Microsoft Visual Basic for Applications");
                    if (error != IntPtr.Zero) break;
                }
                if (error != IntPtr.Zero)
                {
                    string detail = native.Message(error);
                    native.Close(error);
                    throw new InvalidOperationException("VBE rejected the watch expression: " + detail);
                }
                if (!completed) throw new InvalidOperationException("The Add Watch dialog did not close after OK.");
                IntPtr root = native.VbeRoot();
                IntPtr watches = root == IntPtr.Zero ? IntPtr.Zero :
                    native.Pane(native.Children(root), "Espions", "Watch", "Watches");
                object watchState = native.List(watches);
                return new { Added = true, request.Expression,
                    WatchType = string.IsNullOrWhiteSpace(request.WatchType) ? "expression" : request.WatchType,
                    Context = new {
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
                    IntPtr remaining = native.Dialog("Ajouter un espion", "Add Watch");
                    if (remaining != IntPtr.Zero) native.Close(remaining);
                }
            }
        }

        public static object CompleteEditWatch(Request request)
        {
            return CompleteEditWatch(request, new NativeWatchProbe());
        }

        internal static object CompleteEditWatch(Request request, IWatchProbe native)
        {
            IntPtr dialog = IntPtr.Zero;
            for (int attempt = 0; attempt < 60 && dialog == IntPtr.Zero; attempt++)
            { native.Pause(50); dialog = native.Dialog("Modifier un espion", "Edit Watch"); }
            if (dialog == IntPtr.Zero) throw new InvalidOperationException("The Edit Watch dialog did not open.");
            bool completed = false;
            try
            {
                IntPtr edit = native.Item(dialog, 4853);
                IntPtr project = native.Item(dialog, 4858);
                IntPtr module = native.Item(dialog, 4857);
                IntPtr procedure = native.Item(dialog, 4856);
                IntPtr ok = native.Item(dialog, 1);
                if (edit == IntPtr.Zero || project == IntPtr.Zero || module == IntPtr.Zero || ok == IntPtr.Zero)
                    throw new InvalidOperationException("The native Edit Watch controls changed.");
                string original = native.Text(edit);
                string shownProject = native.Text(project);
                string shownModule = native.Text(module);
                string shownProcedure = procedure == IntPtr.Zero ? null : native.Text(procedure);
                string shownContext = shownModule + (string.IsNullOrWhiteSpace(shownProcedure) ? "" : "." + shownProcedure);
                if (!string.Equals(original, request.Expression, StringComparison.Ordinal) ||
                    !string.Equals(shownProject, request.Project, StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(shownContext, request.Context, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("The native Edit Watch selection or context changed.");
                if (!string.IsNullOrWhiteSpace(request.WatchType))
                {
                    int typeId = request.WatchType == "break_when_true" ? 4851 :
                        request.WatchType == "break_when_changed" ? 4852 : 4850;
                    IntPtr typeButton = native.Item(dialog, typeId);
                    if (typeButton == IntPtr.Zero || !native.Click(typeButton))
                        throw new InvalidOperationException("The requested native watch type is unavailable.");
                    native.Pause(30);
                    if (!native.Checked(typeButton))
                        throw new InvalidOperationException("The requested native watch type was not selected.");
                }
                native.Replace(edit, request.NewExpression);
                if (!string.Equals(native.Text(edit), request.NewExpression, StringComparison.Ordinal))
                    throw new InvalidOperationException("The new watch expression was not reflected by the native edit control.");
                if (!native.Click(ok))
                    throw new InvalidOperationException("The Edit Watch dialog refused its OK command.");
                IntPtr error = IntPtr.Zero;
                for (int attempt = 0; attempt < 40; attempt++)
                {
                    native.Pause(50);
                    if (native.Dialog("Modifier un espion", "Edit Watch") == IntPtr.Zero)
                    { completed = true; break; }
                    error = native.Dialog("Microsoft Visual Basic pour Applications", "Microsoft Visual Basic for Applications");
                    if (error != IntPtr.Zero) break;
                }
                if (error != IntPtr.Zero)
                {
                    string detail = native.Message(error);
                    native.Close(error);
                    throw new InvalidOperationException("VBE rejected the edited watch: " + detail);
                }
                if (!completed) throw new InvalidOperationException("The Edit Watch dialog did not close after OK.");
                IntPtr root = native.VbeRoot();
                IntPtr pane = root == IntPtr.Zero ? IntPtr.Zero :
                    native.Pane(native.Children(root), "Espions", "Watch", "Watches");
                if (pane == IntPtr.Zero)
                    return new { Edited = true, OldExpression = request.Expression, request.NewExpression,
                        request.Context, Verification = "Pending", VerificationPending = true,
                        Limit = "The Watches pane is not visible; edit cannot be read back." };
                bool verified = false;
                for (int attempt = 0; attempt < 20; attempt++)
                {
                    bool newPresent = native.WatchMatches(pane, request.NewExpression, request.Context) == 1;
                    bool oldAbsent = request.NewExpression == request.Expression ||
                        native.WatchMatches(pane, request.Expression, request.Context) == 0;
                    if (newPresent && oldAbsent) { verified = true; break; }
                    native.Pause(50);
                }
                return new { Edited = true, OldExpression = request.Expression, request.NewExpression,
                    request.Context, Verification = verified ? "ReadbackVerified" : "Pending",
                    VerificationPending = !verified,
                    Limit = "Expression readback does not independently prove a changed watch break condition." };
            }
            finally
            {
                if (!completed)
                {
                    IntPtr remaining = native.Dialog("Modifier un espion", "Edit Watch");
                    if (remaining != IntPtr.Zero) native.Close(remaining);
                }
            }
        }

        public static object CompleteQuickWatch(Request request)
        {
            return CompleteQuickWatch(request, new NativeWatchProbe());
        }

        internal static object CompleteQuickWatch(Request request, IWatchProbe native)
        {
            IntPtr dialog = IntPtr.Zero;
            for (int attempt = 0; attempt < 60 && dialog == IntPtr.Zero; attempt++)
            { native.Pause(50); dialog = native.Dialog("Espion express", "Quick Watch"); }
            if (dialog == IntPtr.Zero) throw new InvalidOperationException("The Quick Watch dialog did not open.");
            try
            {
                IntPtr expressionControl = native.Item(dialog, 4751);
                IntPtr valueControl = native.Item(dialog, 4752);
                IntPtr contextControl = native.Item(dialog, 4753);
                if (expressionControl == IntPtr.Zero || valueControl == IntPtr.Zero ||
                    contextControl == IntPtr.Zero || native.Item(dialog, 2) == IntPtr.Zero)
                    throw new InvalidOperationException("The native Quick Watch controls changed.");
                string expression = native.Text(expressionControl);
                string value = native.Text(valueControl);
                string context = native.Text(contextControl);
                if (!string.Equals(expression, request.Expression, StringComparison.Ordinal) ||
                    !context.StartsWith(request.Project + "." + request.Module + ".", StringComparison.OrdinalIgnoreCase) ||
                    (!string.IsNullOrWhiteSpace(request.Procedure) &&
                     !string.Equals(context, request.Project + "." + request.Module + "." + request.Procedure,
                         StringComparison.OrdinalIgnoreCase)))
                    throw new InvalidOperationException("Quick Watch expression or context differs from the selected code.");
                return new { Expression = expression, Value = value, Context = context,
                    Verification = "NativeDialogReadback",
                    Limit = "Evaluating a VBA expression may call user code; a displayed value is valid for this paused context only." };
            }
            finally { native.Close(dialog); }
        }

        public static object SelectWatch(Request request)
        {
            return SelectWatch(request, new NativeWatchProbe());
        }

        internal static object SelectWatch(Request request, IWatchProbe native)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Expression) ||
                string.IsNullOrWhiteSpace(request.Context))
                throw new ArgumentException("Expression and Context are required.");
            IntPtr root = native.VbeRoot();
            IntPtr pane = root == IntPtr.Zero ? IntPtr.Zero :
                native.Pane(native.Children(root), "Espions", "Watch", "Watches");
            if (pane == IntPtr.Zero) throw new InvalidOperationException("The native Watches pane must be visible.");
            int matches = native.WatchMatches(pane, request.Expression, request.Context);
            if (matches != 1)
                throw new InvalidOperationException("Expected one matching native watch; found " + matches + ".");
            if (!native.SelectWatchRow(pane, request.Expression, request.Context))
                throw new InvalidOperationException("The native watch row does not support selection.");
            return new { Selected = true, request.Expression, request.Context };
        }

        public static object VerifyWatchRemoved(Request request)
        {
            return VerifyWatchRemoved(request, new NativeWatchProbe());
        }

        internal static object VerifyWatchRemoved(Request request, IWatchProbe native)
        {
            IntPtr root = native.VbeRoot();
            IntPtr pane = root == IntPtr.Zero ? IntPtr.Zero :
                native.Pane(native.Children(root), "Espions", "Watch", "Watches");
            if (pane == IntPtr.Zero)
                return new { Removed = false, VerificationPending = true,
                    Error = "The Watches pane is no longer visible; absence cannot be verified." };
            for (int attempt = 0; attempt < 20; attempt++)
            {
                if (native.WatchMatches(pane, request.Expression, request.Context) == 0)
                    return new { Removed = true, VerificationPending = false, Error = (string)null };
                native.Pause(50);
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
                string raw = elements[index].Current.Name;
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
                if (elements[index].Current.ControlType != ControlType.Text) continue;
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
            if (handle == IntPtr.Zero) return new { Available = false, Items = new object[0], Error = "Window is not visible.",
                Coverage = "UIAExposedRowsOnly" };
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
                        raw = ((ValuePattern)pattern).Current.Value;
                    // Skip the native empty-list placeholder before querying its UIA ancestry.
                    if (ParseDebugRow(raw, null) == null) continue;
                    object parsed = ParseDebugRow(raw, ItemPath(element));
                    if (parsed != null) items.Add(parsed);
                }
                return new { Available = true, Items = items.ToArray(), Error = (string)null,
                    Coverage = "UIAExposedRowsOnly" };
            }
            catch (Exception ex) { return new { Available = true, Items = new object[0], Error = ex.Message,
                Coverage = "UIAExposedRowsOnly" }; }
        }

        internal static object ParseDebugRow(string raw, string[] itemPath)
        {
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
                return null;
            itemPath = itemPath ?? new string[0];
            return new { Raw = raw, Expression = match.Success ? match.Groups[1].Value :
                    watch.Success ? watch.Groups[1].Value : null,
                Value = match.Success ? match.Groups[2].Value.Trim() :
                    watch.Success ? watch.Groups[2].Value.Trim() : null,
                Type = match.Success ? match.Groups[3].Value :
                    watch.Success ? watch.Groups[3].Value : null,
                Context = watch.Success ? watch.Groups[4].Value : null,
                PathSegments = itemPath,
                Depth = Math.Max(0, itemPath.Length - 1),
                Parsed = match.Success || watch.Success };
        }

        private static string[] ItemPath(AutomationElement element)
        {
            var segments = new List<string>();
            AutomationElement current = element;
            for (int depth = 0; depth < 16 && current != null &&
                current.Current.ControlType == ControlType.ListItem; depth++)
            {
                string segment = ParseItemPathSegment(current.Current.Name);
                if (segment == null) break;
                segments.Add(segment);
                current = TreeWalker.RawViewWalker.GetParent(current);
            }
            segments.Reverse();
            return segments.ToArray();
        }

        internal static string ParseItemPathSegment(string name)
        {
            Match match = Regex.Match(name ?? "",
                @"^Expression\s+(.*?)\s+(?:Valeur|Value)\s+", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            if (!match.Success)
                match = Regex.Match(name ?? "",
                    @"^\s*(.*?)\s+(?:Valeur|Value)\s+.*?\s+Type\s+.*?\s+(?:Contexte|Context)\s+",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline);
            return match.Success ? match.Groups[1].Value : null;
        }

        private static string WatchRootContext(AutomationElement element)
        {
            AutomationElement root = element;
            AutomationElement parent;
            while ((parent = TreeWalker.RawViewWalker.GetParent(root)) != null &&
                parent.Current.ControlType == ControlType.ListItem)
                root = parent;
            return ParseWatchContext(root.Current.Name);
        }

        internal static string ParseWatchContext(string name)
        {
            Match match = Regex.Match(name ?? "",
                @"\s+(?:Contexte|Context)\s+(.*?)\s*$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            return match.Success ? match.Groups[1].Value : null;
        }

        private static object ReadImmediate(IntPtr handle)
        {
            if (handle == IntPtr.Zero) return new { Available = false, Text = (string)null, Error = "Window is not visible." };
            try
            {
                return new { Available = true, Text = ImmediateText(handle), Error = (string)null };
            }
            catch (Exception ex) { return new { Available = true, Text = (string)null, Error = ex.Message }; }
        }

        private static AutomationElement ImmediateDocument(IntPtr handle)
        {
            AutomationElement root = AutomationElement.FromHandle(handle);
            var condition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Document);
            AutomationElementCollection documents = root.FindAll(TreeScope.Descendants, condition);
            if (documents.Count != 1)
                throw new InvalidOperationException("Expected one Immediate document; found " + documents.Count + ".");
            if (!documents[0].TryGetCurrentPattern(TextPattern.Pattern, out _))
                throw new InvalidOperationException("The Immediate document does not expose TextPattern.");
            return documents[0];
        }

        private static string ImmediateText(IntPtr handle)
        {
            var pattern = (TextPattern)ImmediateDocument(handle).GetCurrentPattern(TextPattern.Pattern);
            return pattern.DocumentRange.GetText(-1);
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
                { PauseNative(50); dialog = FindCallStackDialog(); }
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
