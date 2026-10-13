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
    internal sealed partial class OllamaOfficeUi
    {
        private delegate bool Visitor(IntPtr hwnd, IntPtr state);
        [DllImport("user32.dll")] private static extern bool EnumWindows(Visitor visitor, IntPtr state);
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, Visitor visitor, IntPtr state);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr hwnd);
        [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr hwnd);
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
        private readonly Action requireObserverOwner;
        private readonly int observerThread = Thread.CurrentThread.ManagedThreadId;
        private IntPtr vbe, chat;
        private IntPtr toolContainer, nativeSite;
        private uint ownerThread;
        private AutomationElement root;
        private IntPtr transcriptPanel;
        private readonly Dictionary<string, AutomationElement> fixedControls = new Dictionary<string, AutomationElement>();
        private readonly Dictionary<string, IntPtr> fixedHandles = new Dictionary<string, IntPtr>();
        private readonly HashSet<IntPtr> recordedAncestors = new HashSet<IntPtr>();
        internal bool SentUnsettled { get; private set; }
        internal bool StopEmitted { get; private set; }

        internal OllamaOfficeUi(int processId, Action<object> record, Action requireObserverOwner = null)
        { pid = processId; this.record = record; this.requireObserverOwner = requireObserverOwner; }

        internal void Discover()
        {
            string desktop = OllamaOfficeDesktop.MainEnabled ? "Default" : Environment.GetEnvironmentVariable("VBAi_TEST_DESKTOP_NAME");
            OllamaOfficeDesktop.Require(desktop);
            bool initialInventory = true;
            object lastInventory = null;
            try
            {
                Wait(() =>
                {
                    var roots = new List<IntPtr>(); var tops = new List<IntPtr>(); var topInventory = new List<object>(); int count = 0;
                    EnumWindows((window, state) =>
                    {
                        if (++count >= 4096) return false;
                        uint process; GetWindowThreadProcessId(window, out process);
                        if (process != pid) return true;
                        var name = new StringBuilder(128); GetClassName(window, name, name.Capacity);
                        var title = new StringBuilder(128); GetWindowText(window, title, title.Capacity);
                        topInventory.Add(new
                        {
                            Hwnd = window.ToInt64(),
                            Class = name.ToString(),
                            Visible = IsWindowVisible(window),
                            Owner = GetWindow(window, 4).ToInt64(),
                            CaptionMatches = title.ToString() == "VBAi — Your AI agent for VBA"
                        });
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
                        EnumChildWindows(top, (window, state) =>
                        {
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
                        var candidate = OfficeUiSnapshot.Refresh(AutomationElement.FromHandle(window), requireObserverOwner);
                        string automationId = candidate.Cached.AutomationId;
                        bool knownControl = automationId == "ChatToolWindow" || automationId == "ChatWindow";
                        inventory.Add(new
                        {
                            Hwnd = window.ToInt64(),
                            Class = name.ToString(),
                            ProcessId = process,
                            NativeThread = thread,
                            CaptionMatches = captionMatches,
                            ChildOfVbe = IsChild(vbe, window),
                            Visible = IsWindowVisible(window),
                            KnownAssistantControl = knownControl,
                            FixedControlId = knownControl ? automationId : null,
                            Type = candidate.Cached.ControlType.ProgrammaticName,
                            ControlIdSha256 = EditorDocument.Hash(automationId ?? ""),
                            Root = GetAncestor(window, 2).ToInt64(),
                            Parent = GetAncestor(window, 1).ToInt64()
                        });
                        // A hosted borderless Form need not expose its caption through cross-process
                        // GetWindowText. The COM container and exact control identities are the proof.
                        if (!IsWindowVisible(window) || (!captionMatches && !knownControl)) continue;
                        if (candidate.Cached.ControlType != ControlType.Window && candidate.Cached.ControlType != ControlType.Pane) continue;
                        var scope = candidate.FindAll(TreeScope.Descendants,
                            new PropertyCondition(AutomationElement.AutomationIdProperty, "scopePicker"));
                        var options = candidate.FindAll(TreeScope.Descendants,
                            new PropertyCondition(AutomationElement.AutomationIdProperty, AssistantControlId("options")));
                        var send = candidate.FindAll(TreeScope.Descendants,
                            new PropertyCondition(AutomationElement.AutomationIdProperty, AssistantControlId("send")));
                        IntPtr container, site;
                        bool hosted = TryHostedSite(window, out container, out site);
                        inventory.Add(new
                        {
                            Candidate = window.ToInt64(),
                            ScopeCount = scope.Count,
                            HostedToolWindow = hosted,
                            ToolContainer = container.ToInt64(),
                            NativeSite = site.ToInt64()
                        });
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
                    record(new
                    {
                        Phase = "ActualEmbeddedAssistant",
                        ProcessId = pid,
                        Vbe = vbe.ToInt64(),
                        Chat = chat.ToInt64(),
                        NativeThread = ownerThread,
                        Desktop = desktop,
                        ChildOfVbe = IsChild(vbe, chat),
                        HostedToolWindow = true,
                        ToolContainer = toolContainer.ToInt64(),
                        NativeSite = nativeSite.ToInt64(),
                        NativeSiteOwnedByVbe = OwnedByVbe(GetAncestor(nativeSite, 2)),
                        ContainerIdentity = OfficeUiSnapshot.Refresh(AutomationElement.FromHandle(toolContainer), requireObserverOwner).Cached.AutomationId,
                        NativeSiteClass = "GenericPane",
                        NativeToolCaptionMatches = true
                    });
                    return true;
                }, 45, "discover the installed embedded assistant");
            }
            catch
            {
                record(new
                {
                    Phase = "AssistantDiscoveryFailed",
                    ProcessId = pid,
                    Vbe = vbe.ToInt64(),
                    Chat = chat.ToInt64(),
                    NativeThread = ownerThread,
                    LastInventory = lastInventory
                }); throw;
            }
        }

        internal void DetectNativeLanguage()
        {
            Guard();
            var vbeElement = AutomationElement.FromHandle(vbe);
            var bars = OfficeUiSnapshot.FindAll(vbeElement, TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuBar));
            if (bars.Count > 16) throw new InvalidOperationException("Bounded native VBE menu inventory.");
            var captions = new List<string>();
            foreach (AutomationElement bar in bars)
            {
                if (bar.Cached.ProcessId != pid) throw new InvalidOperationException("Foreign native VBE menu.");
                var items = OfficeUiSnapshot.FindAll(bar, TreeScope.Children,
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem));
                if (items.Count > 64) throw new InvalidOperationException("Bounded native menu items.");
                foreach (AutomationElement item in items) captions.Add(item.Cached.Name);
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
                var element = OfficeUiSnapshot.Refresh(AutomationElement.FromHandle(parent), requireObserverOwner);
                var cls = new StringBuilder(128); GetClassName(parent, cls, cls.Capacity);
                var title = new StringBuilder(128); GetWindowText(parent, title, title.Capacity);
                if (recordedAncestors.Add(parent)) record(new
                {
                    Phase = "AssistantAncestorIdentity",
                    Hwnd = parent.ToInt64(),
                    ProcessId = process,
                    NativeThread = thread,
                    Class = cls.ToString(),
                    AutomationId = element.Cached.AutomationId,
                    Type = element.Cached.ControlType.ProgrammaticName,
                    ContainerNameMatches = element.Cached.Name == "ChatToolWindow",
                    NativeToolCaptionMatches = title.ToString() == "VBAi",
                    Parent = GetAncestor(parent, 1).ToInt64()
                });
                if (element.Cached.AutomationId != "ChatToolWindow" && element.Cached.AutomationId != "ControlAxSourcingSite") continue;
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
            requireObserverOwner?.Invoke();
            uint process; uint thread = GetWindowThreadProcessId(chat, out process);
            if (process != pid || thread != ownerThread || !IsWindowVisible(chat) ||
                !OriginalHostedSiteMatches())
                throw new InvalidOperationException("Installed assistant ownership/visibility changed.");
            OllamaOfficeDesktop.RequireWindow(Environment.GetEnvironmentVariable("VBAi_TEST_DESKTOP_NAME"), (uint)pid, true, chat);
        }

        /// <summary>Requires the original native control, not a replacement peer with the same automation ID.</summary>
        internal static bool FixedControlIdentityMatches(long originalHandle, long currentHandle, string id,
            string observedId, int expectedPid, uint expectedThread, int uiPid, int nativePid,
            uint nativeThread, bool child, bool visible, bool allowHiddenObservation = false)
            => originalHandle != 0 && originalHandle == currentHandle && id == observedId && expectedPid > 0 &&
                expectedThread != 0 && uiPid == expectedPid && nativePid == expectedPid &&
                nativeThread == expectedThread && child && (visible || allowHiddenObservation);

        internal AutomationElement Leaf(string id, bool allowHiddenObservation = false, bool allowInitialAbsentStop = false)
        {
            Guard();
            AutomationElement cached;
            bool bound = fixedControls.TryGetValue(id, out cached);
            if (!bound)
            {
                cached = FindFixedNativeLeaf(id, allowHiddenObservation, allowInitialAbsentStop);
                if (cached == null) return null;
            }
            // Bound peers are refreshed once per inspection; no stale state survives a poll.
            var snapshot = bound ? OfficeUiSnapshot.Refresh(cached, Guard) : cached;
            var current = snapshot.Cached;
            var handle = new IntPtr(current.NativeWindowHandle);
            IntPtr original = bound ? fixedHandles[id] : handle;
            string automationId = current.AutomationId;
            int uiPid = current.ProcessId;
            uint process; uint thread = GetWindowThreadProcessId(handle, out process);
            bool child = IsChild(chat, handle), visible = IsWindowVisible(handle);
            if (!FixedControlIdentityMatches(original.ToInt64(), handle.ToInt64(), AssistantControlId(id), automationId,
                pid, ownerThread, uiPid, (int)process, thread, child, visible, allowHiddenObservation))
            {
                record(new { Phase = "FixedControlIdentityRefused", Id = id, OriginalHandle = original.ToInt64(),
                    ObservedHandle = handle.ToInt64(), AutomationIdMatches = automationId == AssistantControlId(id),
                    ExpectedPid = pid, UiPid = uiPid, NativePid = process, ExpectedThread = ownerThread,
                    NativeThread = thread, OriginalParentMatches = child, NativeVisible = visible,
                    NativeWindowExists = IsWindow(handle), SentUnsettled, Rebound = false, ActionEmitted = false });
                throw new InvalidOperationException("Fixed assistant control ownership changed: " + id);
            }
            if (!bound) { fixedControls.Add(id, cached); fixedHandles.Add(id, handle); }
            return snapshot;
        }

        /// <summary>Binds a real WinForms HWND once; the host's unrelated or virtual accessibility tree is not enumerated.</summary>
        private AutomationElement FindFixedNativeLeaf(string id, bool allowHiddenObservation, bool allowInitialAbsentStop)
        {
            Guard();
            var peers = new Dictionary<long, AutomationElement>();
            int count = 0;
            var timer = Stopwatch.StartNew();
            record(new { Phase = "NativeFixedControlLookupStarted", Id = id, Utc = DateTime.UtcNow.ToString("o") });
            EnumChildWindows(chat, (window, unused) =>
            {
                if (++count > 2048) return false;
                uint process; uint thread = GetWindowThreadProcessId(window, out process);
                if (process != pid || thread != ownerThread || (!allowHiddenObservation && !IsWindowVisible(window))) return true;
                var name = new StringBuilder(128); GetClassName(window, name, name.Capacity);
                if (!FixedControlClassCanMatch(id, name.ToString())) return true;
                var element = OfficeUiSnapshot.RefreshIdentity(AutomationElement.FromHandle(window), Guard);
                var identity = element.Cached;
                if (identity.AutomationId != AssistantControlId(id)) return true;
                if (identity.ProcessId != pid || unchecked((uint)identity.NativeWindowHandle) != unchecked((uint)window.ToInt64()))
                    throw new InvalidOperationException("Native assistant HWND and UIA peer disagree.");
                peers.Add(window.ToInt64(), element);
                return true;
            }, IntPtr.Zero);
            Guard();
            if (count > 2048) throw new InvalidOperationException("Native assistant control inventory exceeds its bound.");
            // WinForms may defer creating a hidden Stop HWND until the first turn.
            // Once bound, Leaf always validates the retained HWND; it never adopts
            // a replacement or treats a missing original as idle.
            if (allowInitialAbsentStop && InitialStopCanBeAbsent(id, peers.Count)) return null;
            long handle = RequireUniqueFixedWindow(id, peers.Keys.ToArray());
            var result = OfficeUiSnapshot.Refresh(peers[handle], Guard);
            record(new { Phase = "NativeFixedControlLookupObserved", Id = id, Handle = handle,
                ElapsedMilliseconds = timer.ElapsedMilliseconds, NativeWindows = count });
            return result;
        }

        /// <summary>Uses native class metadata to avoid unrelated Word UIA providers while retaining every eligible peer.</summary>
        internal static bool FixedControlClassCanMatch(string id, string nativeClass)
        {
            if (nativeClass == null || !nativeClass.StartsWith("WindowsForms", StringComparison.Ordinal)) return false;
            switch (id)
            {
                case "scopePicker":
                case "providerPicker":
                case "modelPicker":
                case "effortPicker":
                case "modePicker":
                case "approvalPicker":
                    return nativeClass.StartsWith("WindowsForms10.COMBOBOX.", StringComparison.OrdinalIgnoreCase);
                case "modelSummary":
                case "send":
                case "stop":
                case "refreshModels":
                    return nativeClass.StartsWith("WindowsForms10.BUTTON.", StringComparison.OrdinalIgnoreCase);
                case "promptHost":
                case "transcriptPanel":
                    // UserControl and Panel use the WinForms Window class. Neither
                    // can be an edit, combo or button peer; inspect every eligible
                    // panel and still require one exact AutomationId/owned HWND.
                    return nativeClass.StartsWith("WindowsForms10.Window.", StringComparison.OrdinalIgnoreCase);
                default:
                    return true;
            }
        }

        /// <summary>Refuses absent, duplicate or handle-free controls before any action can use a peer.</summary>
        internal static long RequireUniqueFixedWindow(string id, long[] handles)
        {
            if (string.IsNullOrEmpty(id) || handles == null || handles.Length != 1 || handles[0] == 0)
                throw new InvalidOperationException("Expected one original native assistant control: " + id);
            return handles[0];
        }

        internal void BindTranscript()
        {
            if (SentUnsettled) throw new InvalidOperationException("Bind the fixed transcript before sending.");
            transcriptPanel = new IntPtr(Leaf("transcriptPanel").Cached.NativeWindowHandle);
            if (!IsChild(chat, transcriptPanel)) throw new InvalidOperationException("Transcript is outside the owned chat.");
            record(new { Phase = "NativeTranscriptBound", ProcessId = pid, Panel = transcriptPanel.ToInt64(), NativeThread = ownerThread });
        }

        internal void Click(string id)
        {
            var button = Leaf(id);
            IntPtr target = new IntPtr(button.Cached.NativeWindowHandle);
            uint nativePid; uint nativeThread = GetWindowThreadProcessId(target, out nativePid);
            record(new
            {
                Phase = "ButtonPreflight",
                Id = id,
                AutomationId = button.Cached.AutomationId,
                Name = button.Cached.Name,
                UiProcessId = button.Cached.ProcessId,
                NativeProcessId = nativePid,
                NativeThread = nativeThread,
                ExpectedNativeThread = ownerThread,
                Handle = target.ToInt64(),
                UiEnabled = button.Cached.IsEnabled,
                UiOffscreen = button.Cached.IsOffscreen,
                NativeVisible = IsWindowVisible(target),
                NativeEnabled = IsWindowEnabled(target),
                OriginalParentMatches = IsChild(chat, target)
            });
            record(new { Phase = "ButtonIntentOnce", Id = id, Name = button.Cached.Name });
            PrivateDesktopUiAction.ClickButtonOnce(button, AssistantControlId(id), button.Cached.Name, chat,
                new IntPtr(button.Cached.NativeWindowHandle), pid, ownerThread, requireObserverOwner);
        }

        internal void Select(string id, string expected)
        {
            var combo = Leaf(id);
            Wait(() => Leaf(id).Cached.IsEnabled, 30, "enabled " + id);
            if (SelectedValue(combo) == expected) return;
            var comboHandle = new IntPtr(combo.Cached.NativeWindowHandle);
            Wait(() => ExactChoicePresent(id, expected), 30, "exact native choice ready " + id);
            OfficeNativeComboSelect.Select(comboHandle, chat, pid, ownerThread, expected, Guard,
                phase => record(new { Phase = phase, Id = id, Value = expected }));
            Wait(() => SelectedValue(Leaf(id)) == expected,
                10, "read back " + id);
        }

        /// <summary>Focuses the exact owned selector before its single expansion and rechecks ownership between the actions.</summary>
        /// <remarks>A failed or uncertain focus or expansion is never replayed. This uses the control's accessibility provider without keyboard or mouse synthesis.</remarks>
        internal static void ExpandOwnedComboOnce(Action requireOwner, Action focus, Action expand, Action<string> recordIntent)
        {
            if (requireOwner == null || focus == null || expand == null || recordIntent == null)
                throw new ArgumentNullException(nameof(requireOwner));
            requireOwner();
            recordIntent("ComboFocusIntentOnce");
            focus();
            requireOwner();
            recordIntent("ComboExpandIntentOnce");
            expand();
        }

        private bool ExactChoicePresent(string id, string expected)
        {
            var combo = Leaf(id);
            var hwnd = new IntPtr(combo.Cached.NativeWindowHandle);
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
            // UIA SelectionPattern can stall on Word's hosted accessibility tree.
            // The same original HWND exposes its actual displayed selection through
            // bounded native read-only messages. Selection actions remain one-shot.
            return OfficeNativeComboRead.Read(new IntPtr(combo.Cached.NativeWindowHandle), chat,
                pid, ownerThread, Guard);
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
            // Search only the fixed composer host; completed transcript rows must
            // not make later composer reads traverse the whole WPF conversation.
            var fields = OfficeUiSnapshot.FindAll(Leaf("promptHost"), TreeScope.Descendants, new AndCondition(
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit),
                new PropertyCondition(AutomationElement.NameProperty, UiText.Get("Your request; Enter to send, Shift+Enter for a new line"))), Guard);
            if (fields.Count != 1 || fields[0].Cached.ProcessId != pid) throw new InvalidOperationException("Exact owned composer not found.");
            return fields[0];
        }

        internal void SendOnce(string prompt)
        {
            if (SentUnsettled) throw new InvalidOperationException("Prior assistant turn has no observed terminal UI.");
            var originalComposer = Composer();
            var originalRuntimeId = originalComposer.GetRuntimeId();
            var value = (ValuePattern)originalComposer.GetCurrentPattern(ValuePattern.Pattern);
            Func<string> readOriginal = () =>
            {
                var observed = Composer();
                if (!originalRuntimeId.SequenceEqual(observed.GetRuntimeId()))
                    throw new InvalidOperationException("Original composer identity changed; no replacement is adopted.");
                return ((ValuePattern)observed.GetCurrentPattern(ValuePattern.Pattern)).Current.Value;
            };
            WriteComposerOnce(prompt, Guard, readOriginal, value.SetValue, record);
            Wait(() => SendReady(),
                30, "ready original send button after composer update");
            if (OfficeUiSnapshot.ReadOnce(Guard, readOriginal) != prompt)
                throw new InvalidOperationException("Composer changed while waiting for readiness; no overwrite or send.");
            StopEmitted = false; SentUnsettled = true; Click("send");
            Wait(() => IsBusy, 10, "observe active turn");
        }

        /// <summary>Writes once and independently verifies the original editor; diagnostics never contain composer text.</summary>
        internal static void WriteComposerOnce(string prompt, Action requireOwner, Func<string> read,
            Action<string> write, Action<object> record)
        {
            if (prompt == null || requireOwner == null || read == null || write == null || record == null)
                throw new ArgumentNullException(nameof(prompt));
            if (OfficeUiSnapshot.ReadOnce(requireOwner, read) != "")
                throw new InvalidOperationException("Composer is not empty; no overwrite.");
            requireOwner();
            record(new { Phase = "ComposerSetIntentOnce", PromptSha256 = EditorDocument.Hash(prompt), PromptLength = prompt.Length });
            requireOwner();
            Exception writeError = null;
            var timer = Stopwatch.StartNew();
            try { write(prompt); } catch (Exception error) { writeError = error; }
            long writeElapsed = timer.ElapsedMilliseconds;
            string observed;
            try { observed = OfficeUiSnapshot.ReadOnce(requireOwner, read); }
            catch (Exception error)
            {
                record(new { Phase = "ComposerReadbackFailed", ErrorType = error.GetType().FullName,
                    ErrorHResult = error.HResult, WriteErrorType = writeError?.GetType().FullName,
                    WriteErrorHResult = writeError?.HResult, WriteElapsedMilliseconds = writeElapsed });
                throw;
            }
            bool exact = string.Equals(prompt, observed, StringComparison.Ordinal);
            record(new { Phase = "ComposerIndependentReadback", Exact = exact, PromptLength = prompt.Length,
                ObservedLength = observed?.Length, ObservedIsNull = observed == null,
                PromptSha256 = EditorDocument.Hash(prompt), ObservedSha256 = observed == null ? null : EditorDocument.Hash(observed),
                WriteErrorType = writeError?.GetType().FullName, WriteErrorHResult = writeError?.HResult,
                WriteElapsedMilliseconds = writeElapsed, ReadElapsedMilliseconds = timer.ElapsedMilliseconds - writeElapsed });
            if (!exact)
                throw new InvalidOperationException("Composer independent readback differs; no SetValue replay.", writeError);
            record(new { Phase = "ComposerExactReadback", ApiErrorAfterApplication = writeError != null });
        }

        private AssistantActionState ReadActionState(bool includeSelectors)
        {
            var send = Leaf("send", true).Cached;
            var stop = Leaf("stop", true, true);
            return new AssistantActionState
            {
                SendVisible = IsWindowVisible(new IntPtr(send.NativeWindowHandle)) && !send.IsOffscreen,
                SendEnabled = send.IsEnabled && IsWindowEnabled(new IntPtr(send.NativeWindowHandle)),
                StopVisible = stop != null && IsWindowVisible(new IntPtr(stop.Cached.NativeWindowHandle)),
                ProviderEnabled = !includeSelectors || Leaf("providerPicker").Cached.IsEnabled,
                ModeEnabled = !includeSelectors || Leaf("modePicker").Cached.IsEnabled
            };
        }

        private bool SendReady() => CanSend(ReadActionState(false));

        internal bool IsBusy => ReadActionState(false).StopVisible;
        internal void StopOnce()
        {
            if (StopEmitted || !IsBusy) throw new InvalidOperationException("No first Stop is available; do not replay an action.");
            StopEmitted = true; Click("stop");
        }
        internal bool ObserveAlreadyIdle()
        {
            if (!TerminalReady()) return false;
            SentUnsettled = false; return true;
        }
        private bool TerminalReady() => CanFinish(ReadActionState(true));

        internal void WaitIdle(int seconds)
        {
            Wait(TerminalReady, seconds, "observe terminal assistant UI");
            SentUnsettled = false;
        }

        /// <summary>Drops the observer's managed UIA references after terminal UI proof, before host shutdown.</summary>
        internal void ReleaseSettledReferences()
        {
            RequireReferenceRelease(observerThread, Thread.CurrentThread.ManagedThreadId, SentUnsettled);
            fixedControls.Clear();
            fixedHandles.Clear();
            root = null;
            record(new
            {
                Phase = "ObserverReferencesDropped",
                ObserverThread = observerThread,
                NativeMutationEntries = 0,
                NativeRcwReleaseProven = false
            });
        }

        /// <summary>Never abandons pending observation or releases another observer's references.</summary>
        internal static void RequireReferenceRelease(int expectedThread, int currentThread, bool unsettled)
        {
            if (expectedThread <= 0 || expectedThread != currentThread || unsettled)
                throw new InvalidOperationException("The original observer and terminal assistant state are required before reference release.");
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
            EnumChildWindows(panel, (window, state) =>
            {
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
                if (!TranscriptChildCanBeObserved(processId, nativeThread, (int)childPid, childThread, IsChild(panel, handle)))
                    continue;
                Rect bounds;
                if (!IsWindowVisible(handle) || (GetWindowLongPtr(handle, -16).ToInt64() & 0x800) == 0 ||
                    !GetWindowRect(handle, out bounds) || Math.Min(bounds.Right, viewport.Right) <= Math.Max(bounds.Left, viewport.Left) ||
                    Math.Min(bounds.Bottom, viewport.Bottom) <= Math.Max(bounds.Top, viewport.Top)) continue;
                IntPtr length;
                if (SendMessageTimeout(handle, 0x000E, IntPtr.Zero, null, 3, 1000, out length) == IntPtr.Zero)
                {
                    int error = Marshal.GetLastWin32Error();
                    childThread = GetWindowThreadProcessId(handle, out childPid);
                    RequireDeferredTranscriptRead(error, processId, nativeThread, (int)childPid, childThread, IsChild(panel, handle), "length");
                    return new string[0]; // No partial inventory may satisfy an oracle.
                }
                if (length.ToInt64() < 0 || length.ToInt64() > 65536) throw new InvalidOperationException("Bounded native transcript text.");
                // Leave room for a fragment appended between the two independent reads.
                var text = new StringBuilder(65537); IntPtr copied;
                if (ReadTextTimeout(handle, 0x000D, new IntPtr(text.Capacity), text, 3, 1000, out copied) == IntPtr.Zero)
                {
                    int error = Marshal.GetLastWin32Error();
                    childThread = GetWindowThreadProcessId(handle, out childPid);
                    RequireDeferredTranscriptRead(error, processId, nativeThread, (int)childPid, childThread, IsChild(panel, handle), "text");
                    return new string[0];
                }
                if (copied.ToInt64() < 0 || copied.ToInt64() > 65536) throw new InvalidOperationException("Invalid native transcript length.");
                if (!IsWindow(handle)) continue;
                childThread = GetWindowThreadProcessId(handle, out childPid);
                if (!TranscriptChildCanBeObserved(processId, nativeThread, (int)childPid, childThread, IsChild(panel, handle)))
                    continue;
                // Streaming and disclosure changes recycle transcript children. A
                // missing or reparented row supplies no oracle text; a later bounded
                // observation may read the newly rendered row. Never admit foreign
                // ownership or a reused HWND whose role is no longer read-only text.
                var currentClass = new StringBuilder(128); GetClassName(handle, currentClass, currentClass.Capacity);
                if (currentClass.ToString().IndexOf("RichEdit", StringComparison.OrdinalIgnoreCase) < 0 ||
                    !IsWindowVisible(handle) || (GetWindowLongPtr(handle, -16).ToInt64() & 0x800) == 0) continue;
                if (text.Length > 0) texts.Add(text.ToString());
            }
            return texts.ToArray();
        }

        /// <summary>Allows an unavailable read-only snapshot within the original oracle deadline; rejects foreign ownership and other native errors.</summary>
        internal static void RequireDeferredTranscriptRead(int error, int expectedPid, uint expectedThread,
            int actualPid, uint actualThread, bool stillChild, string phase)
        {
            if (!TranscriptChildCanBeObserved(expectedPid, expectedThread, actualPid, actualThread, stillChild)) return;
            if (error == 1460) return; // ERROR_TIMEOUT: the complete observation supplies no text and cannot pass.
            throw new InvalidOperationException("Bounded native transcript " + phase + " read failed on a live child (Win32 " + error + ").");
        }

        /// <summary>Discard recycled or detached read-only rows, while preserving strict foreign-owner rejection.</summary>
        internal static bool TranscriptChildCanBeObserved(int expectedPid, uint expectedThread,
            int actualPid, uint actualThread, bool stillChild)
        {
            if (expectedPid <= 0 || expectedThread == 0)
                throw new InvalidOperationException("Original transcript ownership is required.");
            if (actualPid == 0 && actualThread == 0) return false;
            if (actualPid != expectedPid || actualThread != expectedThread)
                throw new InvalidOperationException("Foreign native transcript child.");
            return stillChild;
        }

        internal void Wait(Func<bool> condition, int seconds, string stage)
        {
            var watch = Stopwatch.StartNew();
            ObserveWithinDeadline(condition, checked(seconds * 1000L), () => watch.ElapsedMilliseconds, Thread.Sleep, stage);
            record(new { Phase = "Observed", Stage = stage, ElapsedMilliseconds = watch.ElapsedMilliseconds });
        }

        /// <summary>Rejects a late successful synchronous read; a blocked UIA call cannot silently enlarge its acceptance window.</summary>
        internal static void ObserveWithinDeadline(Func<bool> condition, long limitMilliseconds,
            Func<long> elapsed, Action<int> pause, string stage)
        {
            if (condition == null || elapsed == null || pause == null) throw new ArgumentNullException(nameof(condition));
            if (limitMilliseconds <= 0) throw new ArgumentOutOfRangeException(nameof(limitMilliseconds));
            while (true)
            {
                if (elapsed() > limitMilliseconds) break;
                bool observed = condition();
                long after = elapsed();
                if (after > limitMilliseconds) break;
                if (observed) return;
                if (after >= limitMilliseconds) break;
                pause((int)Math.Min(40, limitMilliseconds - after));
            }
            throw new TimeoutException("Q028: " + stage + "; no action will be replayed.");
        }
    }
}
