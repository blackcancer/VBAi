using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Automation;

namespace VBAi.Tests.Integration
{
    /// <summary>One-shot actions on the actual installed assistant under the owned, inactive VBE.</summary>
    internal sealed class OllamaOfficeUi
    {
        private delegate bool Visitor(IntPtr hwnd, IntPtr state);
        [DllImport("user32.dll")] private static extern bool EnumWindows(Visitor visitor, IntPtr state);
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, Visitor visitor, IntPtr state);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr hwnd);
        [DllImport("user32.dll")] private static extern bool IsChild(IntPtr parent, IntPtr child);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr hwnd, StringBuilder text, int size);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int size);
        private readonly int pid;
        private readonly Action<object> record;
        private IntPtr vbe, chat;
        private uint ownerThread;
        private AutomationElement root;
        internal bool SentUnsettled { get; private set; }
        internal bool StopEmitted { get; private set; }

        internal OllamaOfficeUi(int processId, Action<object> record)
        { pid = processId; this.record = record; }

        internal void Discover()
        {
            string desktop = Environment.GetEnvironmentVariable("VBAi_TEST_DESKTOP_NAME");
            IsolatedTestDesktop.RequireCurrent(desktop);
            Wait(() => {
                var roots = new List<IntPtr>(); int count = 0;
                EnumWindows((window, state) => {
                    if (++count >= 4096) return false;
                    uint process; GetWindowThreadProcessId(window, out process);
                    if (process != pid || !IsWindowVisible(window)) return true;
                    var name = new StringBuilder(128); GetClassName(window, name, name.Capacity);
                    if (name.ToString() == "wndclass_desked_gsk") roots.Add(window);
                    return true;
                }, IntPtr.Zero);
                if (count >= 4096 || roots.Count > 1) throw new InvalidOperationException("Ambiguous/bounded owned VBE inventory.");
                if (roots.Count == 0) return false;
                vbe = roots[0]; uint nativePid; ownerThread = GetWindowThreadProcessId(vbe, out nativePid);
                IsolatedTestDesktop.RequireOfficeWindowInventory(desktop, (uint)pid, true, vbe);
                var found = new List<IntPtr>(); int children = 0;
                EnumChildWindows(vbe, (window, state) => {
                    if (++children >= 2048) return false;
                    uint process; uint thread = GetWindowThreadProcessId(window, out process);
                    if (process != pid || thread != ownerThread || !IsWindowVisible(window)) return true;
                    var name = new StringBuilder(128); GetClassName(window, name, name.Capacity);
                    if (!name.ToString().StartsWith("WindowsForms", StringComparison.Ordinal)) return true;
                    var title = new StringBuilder(128); GetWindowText(window, title, title.Capacity);
                    if (title.ToString() != "VBAi — Your AI agent for VBA") return true;
                    var candidate = AutomationElement.FromHandle(window);
                    if (candidate.Current.ControlType != ControlType.Window && candidate.Current.ControlType != ControlType.Pane) return true;
                    var scope = candidate.FindAll(TreeScope.Descendants,
                        new PropertyCondition(AutomationElement.AutomationIdProperty, "scopePicker"));
                    if (scope.Count == 1 && IsChild(vbe, window)) found.Add(window);
                    return true;
                }, IntPtr.Zero);
                if (children >= 2048 || found.Count > 1) throw new InvalidOperationException("Ambiguous/bounded installed assistant inventory.");
                if (found.Count == 0) return false;
                chat = found[0]; root = AutomationElement.FromHandle(chat); Guard();
                record(new { Phase = "ActualEmbeddedAssistant", ProcessId = pid, Vbe = vbe.ToInt64(), Chat = chat.ToInt64(),
                    NativeThread = ownerThread, Desktop = desktop, ChildOfVbe = IsChild(vbe, chat) });
                return true;
            }, 45, "discover the installed embedded assistant");
        }

        private void Guard()
        {
            uint process; uint thread = GetWindowThreadProcessId(chat, out process);
            if (process != pid || thread != ownerThread || !IsChild(vbe, chat) || !IsWindowVisible(chat))
                throw new InvalidOperationException("Installed assistant ownership/visibility changed.");
            IsolatedTestDesktop.RequireCurrent(Environment.GetEnvironmentVariable("VBAi_TEST_DESKTOP_NAME"));
            IsolatedTestDesktop.RequireOfficeWindowInventory(Environment.GetEnvironmentVariable("VBAi_TEST_DESKTOP_NAME"), (uint)pid, true, chat);
        }

        internal AutomationElement Leaf(string id)
        {
            Guard();
            var found = root.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.AutomationIdProperty, id));
            if (found.Count != 1 || found[0].Current.ProcessId != pid)
                throw new InvalidOperationException("Expected one owned assistant control: " + id);
            return found[0];
        }

        internal void Click(string id)
        {
            var button = Leaf(id);
            record(new { Phase = "ButtonIntentOnce", Id = id, Name = button.Current.Name });
            PrivateDesktopUiAction.ClickButtonOnce(button, id, button.Current.Name, chat,
                new IntPtr(button.Current.NativeWindowHandle), pid, ownerThread);
        }

        internal void Select(string id, string expected)
        {
            var combo = Leaf(id);
            var value = combo.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
            if (value != null && value.Current.Value == expected) return;
            var expand = (ExpandCollapsePattern)combo.GetCurrentPattern(ExpandCollapsePattern.Pattern);
            record(new { Phase = "ComboExpandIntentOnce", Id = id }); expand.Expand();
            var choices = AutomationElement.RootElement.FindAll(TreeScope.Descendants,
                new AndCondition(new PropertyCondition(AutomationElement.ProcessIdProperty, pid),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem),
                    new PropertyCondition(AutomationElement.NameProperty, expected)));
            if (choices.Count != 1) throw new InvalidOperationException("Ambiguous/missing exact combo choice: " + expected);
            record(new { Phase = "ComboSelectionIntentOnce", Id = id, Value = expected });
            ((SelectionItemPattern)choices[0].GetCurrentPattern(SelectionItemPattern.Pattern)).Select();
            Wait(() => Convert.ToString(((ValuePattern)Leaf(id).GetCurrentPattern(ValuePattern.Pattern)).Current.Value) == expected,
                10, "read back " + id);
        }

        internal void RequireScope(string expectedLabel)
        {
            Wait(() => ((ValuePattern)Leaf("scopePicker").GetCurrentPattern(ValuePattern.Pattern)).Current.Value == expectedLabel,
                15, "exact owned project scope");
            record(new { Phase = "ScopeReadback", Value = expectedLabel });
        }

        private AutomationElement Composer()
        {
            Guard();
            var fields = root.FindAll(TreeScope.Descendants, new AndCondition(
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit),
                new PropertyCondition(AutomationElement.NameProperty, UiText.Get("Your request; Enter to send, Shift+Enter for a new line"))));
            if (fields.Count != 1 || fields[0].Current.ProcessId != pid) throw new InvalidOperationException("Exact owned composer not found.");
            return fields[0];
        }

        internal void SendOnce(string prompt)
        {
            if (SentUnsettled) throw new InvalidOperationException("Prior assistant turn has no observed terminal UI.");
            var value = (ValuePattern)Composer().GetCurrentPattern(ValuePattern.Pattern);
            if (value.Current.Value != "") throw new InvalidOperationException("Composer is not empty; no overwrite.");
            Exception writeError = null;
            record(new { Phase = "ComposerSetIntentOnce", PromptSha256 = EditorDocument.Hash(prompt) });
            try { value.SetValue(prompt); } catch (Exception error) { writeError = error; }
            if (((ValuePattern)Composer().GetCurrentPattern(ValuePattern.Pattern)).Current.Value != prompt)
                throw new InvalidOperationException("Composer independent readback differs; no SetValue replay.", writeError);
            record(new { Phase = "ComposerExactReadback", ApiErrorAfterApplication = writeError != null });
            StopEmitted = false; SentUnsettled = true; Click("send");
            Wait(() => IsBusy, 10, "observe active turn");
        }

        internal bool IsBusy => Leaf("send").Current.Name == UiText.Get("Stop ■");
        internal void StopOnce()
        {
            if (StopEmitted || !IsBusy) throw new InvalidOperationException("No first Stop is available; do not replay an action.");
            StopEmitted = true; Click("send");
        }
        internal bool ObserveAlreadyIdle()
        {
            if (IsBusy || !Leaf("providerPicker").Current.IsEnabled || !Leaf("modePicker").Current.IsEnabled) return false;
            SentUnsettled = false; return true;
        }
        internal void WaitIdle(int seconds)
        {
            Wait(() => !IsBusy && Leaf("send").Current.IsEnabled && Leaf("providerPicker").Current.IsEnabled &&
                Leaf("modePicker").Current.IsEnabled, seconds, "observe terminal assistant UI");
            SentUnsettled = false;
        }

        internal string[] VisibleTranscript()
        {
            Guard();
            var texts = root.FindAll(TreeScope.Descendants, new AndCondition(
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit),
                new PropertyCondition(AutomationElement.IsOffscreenProperty, false))).Cast<AutomationElement>();
            return texts.Where(item => item.Current.Name != UiText.Get("Your request; Enter to send, Shift+Enter for a new line"))
                .Select(item => { object pattern; return item.TryGetCurrentPattern(ValuePattern.Pattern, out pattern)
                    ? ((ValuePattern)pattern).Current.Value : ""; }).Where(text => text.Length > 0).ToArray();
        }

        internal void Wait(Func<bool> condition, int seconds, string stage)
        {
            var watch = Stopwatch.StartNew();
            while (!condition() && watch.Elapsed.TotalSeconds < seconds) Thread.Sleep(40);
            if (!condition()) throw new TimeoutException("Q028: " + stage + "; no action will be replayed.");
            record(new { Phase = "Observed", Stage = stage, ElapsedMilliseconds = watch.ElapsedMilliseconds });
        }
    }
}
