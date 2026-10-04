using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
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
        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr hwnd, uint flag);
        [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr hwnd, uint command);
        [StructLayout(LayoutKind.Sequential)] private struct Rect { internal int Left, Top, Right, Bottom; }
        [StructLayout(LayoutKind.Sequential)] private struct ComboInfo
        {
            internal int Size; internal Rect Item, Button; internal uint ButtonState;
            internal IntPtr Combo, Edit, List;
        }
        [DllImport("user32.dll")] private static extern bool GetComboBoxInfo(IntPtr hwnd, ref ComboInfo info);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr hwnd, StringBuilder text, int size);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int size);
        private readonly int pid;
        private readonly Action<object> record;
        private IntPtr vbe, chat;
        private IntPtr toolContainer, nativeSite;
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
            bool initialInventory = true;
            object lastInventory = null;
            try {
            Wait(() => {
                var roots = new List<IntPtr>(); var tops = new List<IntPtr>(); var topInventory = new List<object>(); int count = 0;
                EnumWindows((window, state) => {
                    if (++count >= 4096) return false;
                    uint process; GetWindowThreadProcessId(window, out process);
                    if (process != pid) return true;
                    var name = new StringBuilder(128); GetClassName(window, name, name.Capacity);
                    var title = new StringBuilder(128); GetWindowText(window, title, title.Capacity);
                    topInventory.Add(new { Hwnd = window.ToInt64(), Class = name.ToString(), Visible = IsWindowVisible(window),
                        Owner = GetWindow(window, 4).ToInt64(), CaptionMatches = title.ToString() == "VBAi — Your AI agent for VBA" });
                    if (!IsWindowVisible(window)) return true;
                    tops.Add(window);
                    if (name.ToString() == "wndclass_desked_gsk") roots.Add(window);
                    return true;
                }, IntPtr.Zero);
                if (count >= 4096 || roots.Count > 1) throw new InvalidOperationException("Ambiguous/bounded owned VBE inventory.");
                if (roots.Count == 0) return false;
                vbe = roots[0]; uint nativePid; ownerThread = GetWindowThreadProcessId(vbe, out nativePid);
                IsolatedTestDesktop.RequireOfficeWindowInventory(desktop, (uint)pid, true, vbe);
                var windows = new HashSet<IntPtr>(); int children = 0;
                foreach (var top in tops.Where(window => window == vbe || OwnedByVbe(window)))
                {
                    windows.Add(top);
                    EnumChildWindows(top, (window, state) => {
                        if (++children >= 2048) return false;
                        windows.Add(window); return true;
                    }, IntPtr.Zero);
                    if (children >= 2048) throw new InvalidOperationException("Bounded native tool site inventory.");
                }
                var found = new List<IntPtr>();
                var inventory = new List<object>();
                foreach (var window in windows)
                {
                    uint process; uint thread = GetWindowThreadProcessId(window, out process);
                    if (process != pid || thread != ownerThread) continue;
                    var name = new StringBuilder(128); GetClassName(window, name, name.Capacity);
                    if (!name.ToString().StartsWith("WindowsForms", StringComparison.Ordinal)) continue;
                    var title = new StringBuilder(128); GetWindowText(window, title, title.Capacity);
                    bool captionMatches = title.ToString() == "VBAi — Your AI agent for VBA";
                    var candidate = AutomationElement.FromHandle(window);
                    string automationId = candidate.Current.AutomationId;
                    bool knownControl = automationId == "ChatToolWindow" || automationId == "ChatWindow";
                    inventory.Add(new { Hwnd = window.ToInt64(), Class = name.ToString(), ProcessId = process,
                        NativeThread = thread, CaptionMatches = captionMatches, ChildOfVbe = IsChild(vbe, window),
                        Visible = IsWindowVisible(window), KnownAssistantControl = knownControl,
                        FixedControlId = knownControl ? automationId : null, Type = candidate.Current.ControlType.ProgrammaticName,
                        ControlIdSha256 = EditorDocument.Hash(automationId ?? ""),
                        Root = GetAncestor(window, 2).ToInt64(), Parent = GetAncestor(window, 1).ToInt64() });
                    // A hosted borderless Form need not expose its caption through cross-process
                    // GetWindowText. The COM container and exact control identities are the proof.
                    if (!IsWindowVisible(window) || (!captionMatches && !knownControl)) continue;
                    if (candidate.Current.ControlType != ControlType.Window && candidate.Current.ControlType != ControlType.Pane) continue;
                    var scope = candidate.FindAll(TreeScope.Descendants,
                        new PropertyCondition(AutomationElement.AutomationIdProperty, "scopePicker"));
                    var options = candidate.FindAll(TreeScope.Descendants,
                        new PropertyCondition(AutomationElement.AutomationIdProperty, "options"));
                    var send = candidate.FindAll(TreeScope.Descendants,
                        new PropertyCondition(AutomationElement.AutomationIdProperty, "send"));
                    IntPtr container, site;
                    bool hosted = TryHostedSite(window, out container, out site);
                    inventory.Add(new { Candidate = window.ToInt64(), ScopeCount = scope.Count, HostedToolWindow = hosted,
                        ToolContainer = container.ToInt64(), NativeSite = site.ToInt64() });
                    if (scope.Count == 1 && options.Count == 1 && send.Count == 1 && hosted) found.Add(window);
                }
                lastInventory = new { ProcessId = pid, Vbe = vbe.ToInt64(), Windows = inventory, TopWindows = topInventory };
                if (initialInventory) { record(new { Phase = "InitialAssistantInventory", Inventory = lastInventory }); initialInventory = false; }
                // A container and its inner Form can expose the same descendants. Prefer the
                // innermost native candidate, while retaining distinct tool sites as ambiguous.
                found = found.Where(window => !found.Any(other => other != window && IsChild(window, other))).ToList();
                if (children >= 2048 || found.Count > 1) throw new InvalidOperationException("Ambiguous/bounded installed assistant inventory.");
                if (found.Count == 0) return false;
                chat = found[0]; root = AutomationElement.FromHandle(chat);
                if (!TryHostedSite(chat, out toolContainer, out nativeSite)) throw new InvalidOperationException("Native tool site changed.");
                Guard();
                record(new { Phase = "ActualEmbeddedAssistant", ProcessId = pid, Vbe = vbe.ToInt64(), Chat = chat.ToInt64(),
                    NativeThread = ownerThread, Desktop = desktop, ChildOfVbe = IsChild(vbe, chat), HostedToolWindow = true,
                    ToolContainer = toolContainer.ToInt64(), NativeSite = nativeSite.ToInt64(), NativeSiteOwnedByVbe = OwnedByVbe(nativeSite) });
                return true;
            }, 45, "discover the installed embedded assistant");
            } catch { record(new { Phase = "AssistantDiscoveryFailed", ProcessId = pid, Vbe = vbe.ToInt64(),
                Chat = chat.ToInt64(), NativeThread = ownerThread, LastInventory = lastInventory }); throw; }
        }

        internal void DetectNativeLanguage()
        {
            Guard();
            var vbeElement = AutomationElement.FromHandle(vbe);
            var bars = vbeElement.FindAll(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuBar));
            if (bars.Count > 16) throw new InvalidOperationException("Bounded native VBE menu inventory.");
            var captions = new List<string>();
            foreach (AutomationElement bar in bars)
            {
                if (bar.Current.ProcessId != pid) throw new InvalidOperationException("Foreign native VBE menu.");
                var items = bar.FindAll(TreeScope.Children,
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem));
                if (items.Count > 64) throw new InvalidOperationException("Bounded native menu items.");
                foreach (AutomationElement item in items) captions.Add(item.Current.Name);
            }
            if (!captions.Any(caption => UiLanguages.IsMenu(caption, true) || UiLanguages.IsMenu(caption, false)))
                throw new InvalidOperationException("Native VBE menu language was not observed.");
            var culture = UiLanguages.FromMenus(captions, CultureInfo.CurrentUICulture);
            VBAi.Tests.Infrastructure.LocalizationScope.Set(culture.Name);
            record(new { Phase = "NativeVbeLanguageObserved", Culture = culture.Name, RecognizedMenu = true });
        }

        private bool OwnedByVbe(IntPtr window)
        {
            var seen = new HashSet<IntPtr>();
            for (int i = 0; i < 64 && window != IntPtr.Zero && seen.Add(window); i++)
            {
                uint process; uint thread = GetWindowThreadProcessId(window, out process);
                if (process != pid || thread != ownerThread) return false;
                window = GetWindow(window, 4); // GW_OWNER: never confuse a standalone form with a native child.
                if (window == vbe) return true;
            }
            return false;
        }

        private bool TryHostedSite(IntPtr window, out IntPtr container, out IntPtr site)
        {
            container = site = IntPtr.Zero;
            var seen = new HashSet<IntPtr>();
            for (IntPtr parent = window; parent != IntPtr.Zero && seen.Count < 64 && seen.Add(parent);
                parent = GetAncestor(parent, 1))
            {
                uint process; uint thread = GetWindowThreadProcessId(parent, out process);
                if (process != pid || thread != ownerThread) return false;
                var element = AutomationElement.FromHandle(parent);
                if (element.Current.AutomationId != "ChatToolWindow") continue;
                container = parent; site = GetAncestor(parent, 2);
                return site != IntPtr.Zero && (container == window || IsChild(container, window)) &&
                    (IsChild(vbe, window) || OwnedByVbe(site));
            }
            return false;
        }

        private void Guard()
        {
            uint process; uint thread = GetWindowThreadProcessId(chat, out process);
            IntPtr container, site;
            if (process != pid || thread != ownerThread || !IsWindowVisible(chat) ||
                !TryHostedSite(chat, out container, out site) || container != toolContainer || site != nativeSite)
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
            Wait(() => Leaf(id).Current.IsEnabled, 30, "enabled " + id);
            if (SelectedValue(combo) == expected) return;
            var comboHandle = new IntPtr(combo.Current.NativeWindowHandle);
            var info = new ComboInfo { Size = Marshal.SizeOf(typeof(ComboInfo)) };
            if (!IsChild(chat, comboHandle) || !GetComboBoxInfo(comboHandle, ref info) ||
                info.Combo != comboHandle || info.List == IntPtr.Zero)
                throw new InvalidOperationException("Exact native combo popup identity unavailable: " + id);
            var expand = (ExpandCollapsePattern)combo.GetCurrentPattern(ExpandCollapsePattern.Pattern);
            record(new { Phase = "ComboExpandIntentOnce", Id = id }); expand.Expand();
            var choices = AutomationElement.RootElement.FindAll(TreeScope.Descendants,
                new AndCondition(new PropertyCondition(AutomationElement.ProcessIdProperty, pid),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem),
                    new PropertyCondition(AutomationElement.NameProperty, expected)));
            if (choices.Count != 1) throw new InvalidOperationException("Ambiguous/missing exact combo choice: " + expected);
            uint popupPid; uint popupThread = GetWindowThreadProcessId(info.List, out popupPid);
            if (popupPid != pid || popupThread != ownerThread || !IsWindowVisible(info.List))
                throw new InvalidOperationException("Exact combo popup left its native owner.");
            IsolatedTestDesktop.RequireOfficeWindowInventory(Environment.GetEnvironmentVariable("VBAi_TEST_DESKTOP_NAME"), (uint)pid, true, info.List);
            bool inExactList = false; var node = choices[0];
            for (int depth = 0; depth < 32 && node != null; depth++, node = TreeWalker.RawViewWalker.GetParent(node))
                if (node.Current.NativeWindowHandle == info.List.ToInt64()) { inExactList = true; break; }
            if (!inExactList) throw new InvalidOperationException("Choice does not belong to the exact combo popup.");
            Guard();
            record(new { Phase = "ComboSelectionIntentOnce", Id = id, Value = expected });
            ((SelectionItemPattern)choices[0].GetCurrentPattern(SelectionItemPattern.Pattern)).Select();
            Wait(() => SelectedValue(Leaf(id)) == expected,
                10, "read back " + id);
        }

        private string SelectedValue(AutomationElement combo)
        {
            object pattern;
            if (combo.TryGetCurrentPattern(SelectionPattern.Pattern, out pattern))
            {
                var selected = ((SelectionPattern)pattern).Current.GetSelection();
                if (selected.Length == 0) return "";
                if (selected.Length != 1 || selected[0].Current.ProcessId != pid)
                    throw new InvalidOperationException("Ambiguous or foreign selected combo item.");
                return selected[0].Current.Name;
            }
            if (combo.TryGetCurrentPattern(ValuePattern.Pattern, out pattern)) return ((ValuePattern)pattern).Current.Value;
            throw new InvalidOperationException("Combo exposes neither exact selection nor value readback.");
        }

        internal void RequireScope(string expectedLabel)
        {
            Wait(() => SelectedValue(Leaf("scopePicker")) == expectedLabel,
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
                new OrCondition(new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Document)),
                new PropertyCondition(AutomationElement.IsOffscreenProperty, false))).Cast<AutomationElement>();
            if (texts.Count() > 256) throw new InvalidOperationException("Bounded visible transcript control inventory.");
            return texts.Where(item => item.Current.Name != UiText.Get("Your request; Enter to send, Shift+Enter for a new line"))
                .Select(ReadVisibleText).Where(text => text.Length > 0).ToArray();
        }

        private string ReadVisibleText(AutomationElement item)
        {
            if (item.Current.ProcessId != pid) throw new InvalidOperationException("Foreign transcript element.");
            object pattern;
            if (item.TryGetCurrentPattern(TextPattern.Pattern, out pattern))
            {
                var ranges = ((TextPattern)pattern).GetVisibleRanges();
                if (ranges.Length > 64) throw new InvalidOperationException("Bounded visible text ranges.");
                var text = new StringBuilder();
                foreach (var range in ranges)
                {
                    int remaining = 65536 - text.Length;
                    if (remaining <= 0) break;
                    text.Append(range.GetText(remaining));
                }
                return text.ToString();
            }
            if (item.TryGetCurrentPattern(ValuePattern.Pattern, out pattern))
            {
                string text = ((ValuePattern)pattern).Current.Value;
                if (text.Length > 65536) throw new InvalidOperationException("Bounded transcript value.");
                return text;
            }
            return "";
        }

        internal void Wait(Func<bool> condition, int seconds, string stage)
        {
            var watch = Stopwatch.StartNew();
            bool observed;
            while (!(observed = condition()) && watch.Elapsed.TotalSeconds < seconds) Thread.Sleep(40);
            if (!observed) throw new TimeoutException("Q028: " + stage + "; no action will be replayed.");
            record(new { Phase = "Observed", Stage = stage, ElapsedMilliseconds = watch.ElapsedMilliseconds });
        }
    }
}
