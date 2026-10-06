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
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr hwnd);
        [DllImport("user32.dll")] private static extern bool IsChild(IntPtr parent, IntPtr child);
        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr hwnd, uint flag);
        [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr hwnd, uint command);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr SendMessageTimeout(IntPtr hwnd, uint message, IntPtr wParam, string text,
            uint flags, uint timeout, out IntPtr result);
        [DllImport("user32.dll", EntryPoint = "SendMessageTimeoutW", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr ReadTextTimeout(IntPtr hwnd, uint message, IntPtr wParam, StringBuilder text,
            uint flags, uint timeout, out IntPtr result);
        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")]
        private static extern IntPtr GetWindowLongPtr(IntPtr hwnd, int index);
        [StructLayout(LayoutKind.Sequential)] private struct Rect { internal int Left, Top, Right, Bottom; }
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr hwnd, out Rect rect);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr hwnd, StringBuilder text, int size);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int size);
        private readonly int pid;
        private readonly Action<object> record;
        private IntPtr vbe, chat;
        private IntPtr toolContainer, nativeSite;
        private uint ownerThread;
        private AutomationElement root;
        private IntPtr transcriptPanel;
        private readonly Dictionary<string, AutomationElement> fixedControls = new Dictionary<string, AutomationElement>();
        private readonly HashSet<IntPtr> recordedAncestors = new HashSet<IntPtr>();
        internal bool SentUnsettled { get; private set; }
        internal bool StopEmitted { get; private set; }

        internal OllamaOfficeUi(int processId, Action<object> record)
        { pid = processId; this.record = record; }

        internal void Discover()
        {
            string desktop = OllamaOfficeDesktop.MainEnabled ? "Default" : Environment.GetEnvironmentVariable("VBAi_TEST_DESKTOP_NAME");
            OllamaOfficeDesktop.Require(desktop);
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
                OllamaOfficeDesktop.RequireWindow(desktop, (uint)pid, true, vbe);
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
                    ToolContainer = toolContainer.ToInt64(), NativeSite = nativeSite.ToInt64(),
                    NativeSiteOwnedByVbe = OwnedByVbe(GetAncestor(nativeSite, 2)),
                    ContainerIdentity = AutomationElement.FromHandle(toolContainer).Current.AutomationId,
                    NativeSiteClass = "GenericPane", NativeToolCaptionMatches = true });
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
                var cls = new StringBuilder(128); GetClassName(parent, cls, cls.Capacity);
                var title = new StringBuilder(128); GetWindowText(parent, title, title.Capacity);
                if (recordedAncestors.Add(parent)) record(new { Phase = "AssistantAncestorIdentity", Hwnd = parent.ToInt64(),
                    ProcessId = process, NativeThread = thread, Class = cls.ToString(),
                    AutomationId = element.Current.AutomationId, Type = element.Current.ControlType.ProgrammaticName,
                    ContainerNameMatches = element.Current.Name == "ChatToolWindow", NativeToolCaptionMatches = title.ToString() == "VBAi",
                    Parent = GetAncestor(parent, 1).ToInt64() });
                if (element.Current.AutomationId != "ChatToolWindow" && element.Current.AutomationId != "ControlAxSourcingSite") continue;
                // VBIDE supplies ControlAxSourcingSite as the ActiveX container's runtime
                // name. Its immediate native tool pane must still be the exact VBAi site.
                if (!cls.ToString().StartsWith("WindowsForms", StringComparison.Ordinal)) return false;
                container = parent; site = GetAncestor(parent, 1);
                uint sitePid; uint siteThread = GetWindowThreadProcessId(site, out sitePid);
                var siteClass = new StringBuilder(128); GetClassName(site, siteClass, siteClass.Capacity);
                var siteTitle = new StringBuilder(128); GetWindowText(site, siteTitle, siteTitle.Capacity);
                return site != IntPtr.Zero && sitePid == pid && siteThread == ownerThread && IsWindowVisible(site) &&
                    siteClass.ToString() == "GenericPane" && siteTitle.ToString() == "VBAi" &&
                    (container == window || IsChild(container, window)) &&
                    (IsChild(vbe, site) || OwnedByVbe(GetAncestor(site, 2)));
            }
            return false;
        }

        internal static bool HostedNodesMatch(int expectedPid, uint expectedThread,
            int containerPid, uint containerThread, int sitePid, uint siteThread,
            string containerClass, string siteClass, string siteCaption,
            bool containerVisible, bool siteVisible, bool parentMatches, bool ownedByVbe)
            => expectedPid > 0 && expectedThread != 0 && containerPid == expectedPid && sitePid == expectedPid &&
                containerThread == expectedThread && siteThread == expectedThread &&
                containerClass != null && containerClass.StartsWith("WindowsForms", StringComparison.Ordinal) &&
                siteClass == "GenericPane" && siteCaption == "VBAi" && containerVisible && siteVisible &&
                parentMatches && ownedByVbe;

        private bool OriginalHostedSiteMatches()
        {
            // Discovery admits the ActiveX identities once through UIA. Subsequent
            // guards verify those original native HWNDs, never rebind a provider or
            // rediscover a replacement container during an unsettled Send.
            if (!IsWindow(toolContainer) || !IsWindow(nativeSite)) return false;
            uint containerPid, sitePid;
            uint containerThread = GetWindowThreadProcessId(toolContainer, out containerPid);
            uint siteThread = GetWindowThreadProcessId(nativeSite, out sitePid);
            var containerClass = new StringBuilder(128); var siteClass = new StringBuilder(128);
            var caption = new StringBuilder(128);
            if (GetClassName(toolContainer, containerClass, containerClass.Capacity) == 0 ||
                GetClassName(nativeSite, siteClass, siteClass.Capacity) == 0) return false;
            GetWindowText(nativeSite, caption, caption.Capacity);
            return HostedNodesMatch(pid, ownerThread, (int)containerPid, containerThread, (int)sitePid, siteThread,
                containerClass.ToString(), siteClass.ToString(), caption.ToString(),
                IsWindowVisible(toolContainer), IsWindowVisible(nativeSite),
                GetAncestor(toolContainer, 1) == nativeSite && (toolContainer == chat || IsChild(toolContainer, chat)),
                IsChild(vbe, nativeSite) || OwnedByVbe(GetAncestor(nativeSite, 2)));
        }

        private void Guard()
        {
            uint process; uint thread = GetWindowThreadProcessId(chat, out process);
            if (process != pid || thread != ownerThread || !IsWindowVisible(chat) ||
                !OriginalHostedSiteMatches())
                throw new InvalidOperationException("Installed assistant ownership/visibility changed.");
            OllamaOfficeDesktop.RequireWindow(Environment.GetEnvironmentVariable("VBAi_TEST_DESKTOP_NAME"), (uint)pid, true, chat);
        }

        internal AutomationElement Leaf(string id)
        {
            Guard();
            AutomationElement cached;
            if (fixedControls.TryGetValue(id, out cached))
            {
                var handle = new IntPtr(cached.Current.NativeWindowHandle);
                uint process; uint thread = GetWindowThreadProcessId(handle, out process);
                if (cached.Current.AutomationId != id || cached.Current.ProcessId != pid || process != pid ||
                    thread != ownerThread || !IsChild(chat, handle) || !IsWindowVisible(handle))
                    throw new InvalidOperationException("Fixed assistant control ownership changed: " + id);
                return cached;
            }
            var found = root.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.AutomationIdProperty, id));
            if (found.Count != 1 || found[0].Current.ProcessId != pid)
                throw new InvalidOperationException("Expected one owned assistant control: " + id);
            fixedControls.Add(id, found[0]);
            return found[0];
        }

        internal void BindTranscript()
        {
            if (SentUnsettled) throw new InvalidOperationException("Bind the fixed transcript before sending.");
            transcriptPanel = new IntPtr(Leaf("transcriptPanel").Current.NativeWindowHandle);
            if (!IsChild(chat, transcriptPanel)) throw new InvalidOperationException("Transcript is outside the owned chat.");
            record(new { Phase = "NativeTranscriptBound", ProcessId = pid, Panel = transcriptPanel.ToInt64(), NativeThread = ownerThread });
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
            Wait(() => ExactChoicePresent(id, expected), 30, "exact native choice ready " + id);
            var expand = (ExpandCollapsePattern)combo.GetCurrentPattern(ExpandCollapsePattern.Pattern);
            record(new { Phase = "ComboExpandIntentOnce", Id = id }); expand.Expand();
            string desktop = OllamaOfficeDesktop.MainEnabled ? "Default" : Environment.GetEnvironmentVariable("VBAi_TEST_DESKTOP_NAME");
            var choice = VBAi.Desktop.Helper.NativeComboListDiscovery.RequireExactItem(comboHandle, expected,
                pid, ownerThread, desktop, thread => {
                    if (thread != ownerThread) throw new InvalidOperationException("Foreign combo/list thread.");
                    // Desktop identity comes from the exact owned window inventory. Querying
                    // GetThreadDesktop for a foreign process is not a reliable proof here.
                    Guard(); OllamaOfficeDesktop.RequireWindow(desktop, (uint)pid, true, comboHandle);
                    return desktop;
                }, record);
            Guard();
            record(new { Phase = "ComboSelectionIntentOnce", Id = id, Value = expected });
            ((SelectionItemPattern)choice.GetCurrentPattern(SelectionItemPattern.Pattern)).Select();
            Wait(() => SelectedValue(Leaf(id)) == expected,
                10, "read back " + id);
        }

        private bool ExactChoicePresent(string id, string expected)
        {
            var combo = Leaf(id);
            var hwnd = new IntPtr(combo.Current.NativeWindowHandle);
            uint process; uint thread = GetWindowThreadProcessId(hwnd, out process);
            if (!IsChild(chat, hwnd) || process != pid || thread != ownerThread || expected.Length > 2048)
                throw new InvalidOperationException("Exact owned combo identity changed.");
            IntPtr index;
            // CB_FINDSTRINGEXACT is a read, marshalled by user32 across the owned process.
            // It lets project/catalogue refresh settle before the one popup expansion.
            if (SendMessageTimeout(hwnd, 0x0158, new IntPtr(-1), expected, 3, 3000, out index) == IntPtr.Zero)
                throw new InvalidOperationException("Bounded native combo readiness read did not return.");
            if (index.ToInt64() == -1) return false;
            if (index.ToInt64() < 0 || index.ToInt64() >= 4096) throw new InvalidOperationException("Bounded native combo index.");
            return true;
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
            var texts = ReadNativeTranscript(transcriptPanel, pid, ownerThread);
            Guard();
            return texts;
        }

        // The transcript recycles WinForms RichEdit children inside WPF. Enumerating
        // its changing UIA tree can block until generation ends or return stale nodes.
        // Observe only visible read-only native RichEdit controls inside the fixed
        // transcript panel, with bounded marshalled WM_GETTEXT reads; emit no action.
        internal static string[] ReadNativeTranscript(IntPtr panel, int processId, uint nativeThread)
        {
            uint panelPid;
            if (panel == IntPtr.Zero || GetWindowThreadProcessId(panel, out panelPid) != nativeThread ||
                panelPid != processId || !IsWindowVisible(panel))
                throw new InvalidOperationException("Missing or foreign native transcript panel.");
            Rect viewport;
            if (!GetWindowRect(panel, out viewport)) throw new InvalidOperationException("Transcript viewport unavailable.");
            var handles = new List<IntPtr>(); int count = 0;
            EnumChildWindows(panel, (window, state) => {
                if (++count >= 2048) return false;
                var name = new StringBuilder(128); GetClassName(window, name, name.Capacity);
                if (name.ToString().IndexOf("RichEdit", StringComparison.OrdinalIgnoreCase) >= 0) handles.Add(window);
                return true;
            }, IntPtr.Zero);
            if (count >= 2048 || handles.Count > 256) throw new InvalidOperationException("Bounded native transcript inventory.");
            var texts = new List<string>();
            foreach (var handle in handles)
            {
                if (!IsWindow(handle)) continue; // A recycled child may disappear between observations.
                uint childPid; uint childThread = GetWindowThreadProcessId(handle, out childPid);
                if (childPid != processId || childThread != nativeThread || !IsChild(panel, handle))
                    throw new InvalidOperationException("Foreign native transcript child.");
                Rect bounds;
                if (!IsWindowVisible(handle) || (GetWindowLongPtr(handle, -16).ToInt64() & 0x800) == 0 ||
                    !GetWindowRect(handle, out bounds) || Math.Min(bounds.Right, viewport.Right) <= Math.Max(bounds.Left, viewport.Left) ||
                    Math.Min(bounds.Bottom, viewport.Bottom) <= Math.Max(bounds.Top, viewport.Top)) continue;
                IntPtr length;
                if (SendMessageTimeout(handle, 0x000E, IntPtr.Zero, null, 3, 1000, out length) == IntPtr.Zero)
                    throw new InvalidOperationException("Bounded native transcript length read failed.");
                if (length.ToInt64() < 0 || length.ToInt64() > 65536) throw new InvalidOperationException("Bounded native transcript text.");
                // Leave room for a fragment appended between the two independent reads.
                var text = new StringBuilder(65537); IntPtr copied;
                if (ReadTextTimeout(handle, 0x000D, new IntPtr(text.Capacity), text, 3, 1000, out copied) == IntPtr.Zero)
                    throw new InvalidOperationException("Bounded native transcript text read failed.");
                if (copied.ToInt64() < 0 || copied.ToInt64() > 65536) throw new InvalidOperationException("Invalid native transcript length.");
                if (!IsWindow(handle)) continue;
                if (GetWindowThreadProcessId(handle, out childPid) != nativeThread || childPid != processId || !IsChild(panel, handle))
                    throw new InvalidOperationException("Native transcript child changed during read.");
                if (text.Length > 0) texts.Add(text.ToString());
            }
            return texts.ToArray();
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
