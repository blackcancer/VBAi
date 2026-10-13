using System;
using System.Collections.Generic;
using System.IO;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Automation;

/// <summary>Qualification-only handling of the observed, path-named SOLIDWORKS DirectUI save questions.</summary>
/// <remarks>Every input requires the caller's owned-project, process and desktop guard. Delivery alone is never completion.</remarks>
public static class Q020TaskDialogTeardown
{
    /// <summary>Known construction boundary before any command binding or input exists; never a reason to retry a delivered action.</summary>
    public sealed class NotReadyException : InvalidOperationException
    { public NotReadyException(string message) : base(message) { } }
    /// <summary>A complete read is temporarily unavailable; waiting may repeat observation, never a delivered command.</summary>
    public sealed class ObservationUnavailableException : InvalidOperationException
    { public ObservationUnavailableException(string message) : base(message) { } }
    /// <summary>Immutable identity of the one already-owned macro and its original native/VBE roots.</summary>
    public sealed class Scope
    {
        public readonly int Pid;
        public readonly uint Thread;
        public readonly long VbeRoot, NativeRoot;
        public readonly string Macro;
        public Scope(int pid, uint thread, long vbeRoot, long nativeRoot, string macro)
        {
            if (pid <= 0 || thread == 0 || vbeRoot <= 0 || nativeRoot <= 0 || vbeRoot == nativeRoot)
                throw new ArgumentException("Original owned process, UI thread and distinct roots are required.");
            if (string.IsNullOrWhiteSpace(macro) || !Path.IsPathRooted(macro) || macro.StartsWith("\\\\", StringComparison.Ordinal) ||
                !string.Equals(Path.GetFullPath(macro), macro, StringComparison.Ordinal) ||
                !string.Equals(Path.GetExtension(macro), ".swp", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("One canonical local absolute disposable SWP path is required.");
            Pid = pid; Thread = thread; VbeRoot = vbeRoot; NativeRoot = nativeRoot; Macro = macro;
        }
    }

    /// <summary>Read-only-at-use semantic node; the plan retains validated immutable bindings, not caller-owned arrays.</summary>
    public sealed class Element
    {
        public string Id, Name, Type;
        public long Handle;
        public int Pid;
        public bool Enabled, Offscreen, Actionable;
    }
    /// <summary>Physical HWND identity independently checked against an accessible command node.</summary>
    public sealed class Button
    {
        public long Handle, Root;
        public int Pid, Id;
        public uint Thread;
        public string Class, Caption;
        public bool Visible, Enabled;
    }
    /// <summary>Complete bounded snapshot of one existing DirectUI modal, including its original owner root.</summary>
    public sealed class Snapshot
    {
        public long Dialog, OwnerRoot;
        public int Pid;
        public uint Thread;
        public string Class, Caption;
        public bool Visible, Enabled;
        public Element[] Elements;
        public Button[] Buttons;
    }
    /// <summary>Immutable binding to the exact original command controls and full path-named question.</summary>
    public sealed class Binding
    {
        public readonly long Dialog, OwnerRoot, Yes, No, Cancel;
        public readonly string Prompt, YesCaption, NoCaption, CancelCaption;
        internal Binding(Snapshot s, string prompt, Button yes, Button no, Button cancel)
        {
            Dialog = s.Dialog; OwnerRoot = s.OwnerRoot; Yes = yes.Handle; No = no.Handle; Cancel = cancel.Handle;
            Prompt = prompt; YesCaption = yes.Caption; NoCaption = no.Caption; CancelCaption = cancel.Caption;
        }
    }

    /// <summary>Validates complete ownership, exact localized question, command set and physical/semantic HWND agreement.</summary>
    public static Binding[] Admit(Scope scope, Snapshot[] snapshots)
    {
        if (scope == null) throw new ArgumentNullException("scope");
        if (snapshots == null || snapshots.Length < 1 || snapshots.Length > 2)
            throw new InvalidOperationException("One or two complete captured save questions are required.");
        var dialogs = new HashSet<long>(); var handles = new HashSet<long>(); var result = new List<Binding>();
        foreach (var s in snapshots)
        {
            if (s == null || s.Dialog <= 0 || !dialogs.Add(s.Dialog) || s.Dialog == scope.VbeRoot || s.Dialog == scope.NativeRoot ||
                s.Pid != scope.Pid || s.Thread != scope.Thread || s.Class != "#32770" || s.Caption != "SOLIDWORKS" ||
                !s.Visible || !s.Enabled || (s.OwnerRoot != scope.VbeRoot && s.OwnerRoot != scope.NativeRoot))
                throw new InvalidOperationException("Captured modal ownership or original root differs.");
            if (s.Elements == null || s.Elements.Length == 0 || s.Elements.Length > 128 ||
                s.Elements.Any(e => e == null || e.Pid != scope.Pid || string.IsNullOrEmpty(e.Type)))
                throw new InvalidOperationException("Complete bounded accessible inventory required.");
            var texts = s.Elements.Where(e => e.Id == "ContentText").ToArray();
            if (texts.Length != 1 || texts[0].Type != "ControlType.Text" || texts[0].Offscreen || !texts[0].Enabled)
                throw new InvalidOperationException("One visible readable native question is required.");
            bool french = texts[0].Name == "Enregistrer les modifications apportées à " + scope.Macro + " ?";
            bool english = texts[0].Name == "Save changes to " + scope.Macro + "?";
            if (!french && !english) throw new InvalidOperationException("The exact owned macro path is not the save question.");
            var chrome = s.Elements.Where(e => e.Id == "TitleBar" || e.Id == "SystemMenuBar" || e.Id == "Item 1" || e.Id == "Close").ToArray();
            if (chrome.Length != 0 && (chrome.Length != 4 || chrome.Select(e => e.Id).Distinct().Count() != 4 ||
                chrome.Any(e => !KnownChrome(e, french))))
                throw new InvalidOperationException("Native title-bar inventory differs; no discard.");
            var icons = s.Elements.Where(e => e.Id == "MainIcon").ToArray();
            if (icons.Length > 1 || icons.Any(e => !KnownIcon(e)))
                throw new InvalidOperationException("Native instruction icon differs; no discard.");
            var commands = s.Elements.Where(e => e.Id != null && e.Id.StartsWith("CommandButton_", StringComparison.Ordinal)).ToArray();
            if (commands.Length != 3 || commands.Any(e => e.Offscreen || !e.Enabled ||
                (e.Type != "ControlType.Pane" && e.Type != "ControlType.Button")))
                throw new InvalidOperationException("Exactly three available native commands are required.");
            if (s.Elements.Any(e => !e.Offscreen && !commands.Contains(e) && e != texts[0] &&
                !KnownIcon(e) && !KnownChrome(e, french) && (e.Actionable || !string.IsNullOrEmpty(e.Name))))
                throw new InvalidOperationException("An additional named or actionable control prevents discard.");
            if (s.Buttons == null || s.Buttons.Length != 3 || s.Buttons.Any(b => b == null || b.Handle <= 0 ||
                b.Root != s.Dialog || b.Pid != scope.Pid || b.Thread != scope.Thread || b.Class != "Button" || b.Id != 0 ||
                !b.Visible || !b.Enabled || !handles.Add(b.Handle)))
                throw new InvalidOperationException("Physical native command identity differs.");
            var labels = french ? new[] { "Oui", "Non", "Annuler" } : new[] { "Yes", "No", "Cancel" };
            var ids = new[] { "CommandButton_6", "CommandButton_7", "CommandButton_2" };
            var bound = new Button[3];
            for (int i = 0; i < 3; i++)
            {
                var nodes = commands.Where(e => e.Id == ids[i] && e.Name == labels[i]).ToArray();
                if (nodes.Length != 1 || nodes[0].Handle <= 0) throw new InvalidOperationException("Exact named command unavailable.");
                var physical = s.Buttons.Where(b => b.Handle == nodes[0].Handle && (b.Caption ?? "").Replace("&", "") == labels[i]).ToArray();
                if (physical.Length != 1) throw new InvalidOperationException("Accessible command and physical HWND differ.");
                bound[i] = physical[0];
            }
            if (bound.Select(b => b.Handle).Distinct().Count() != 3)
                throw new InvalidOperationException("Save command handles are ambiguous.");
            result.Add(new Binding(s, texts[0].Name, bound[0], bound[1], bound[2]));
        }
        if (result.Any(b => dialogs.Contains(b.Yes) || dialogs.Contains(b.No) || dialogs.Contains(b.Cancel)))
            throw new InvalidOperationException("A dialog cannot be its own or another dialog's command HWND.");
        return result.ToArray();
    }

    // UI Automation includes nonclient controls and a Value pattern on the instruction image.
    // They have no command HWND and must match the complete observed Windows title-bar set.
    private static bool KnownIcon(Element e)
    { return e.Id == "MainIcon" && e.Name == "MainInstructionIcon" && e.Type == "ControlType.Image" && e.Handle == 0 && e.Enabled && !e.Offscreen; }
    private static bool KnownChrome(Element e, bool french)
    {
        if (e.Handle != 0 || !e.Enabled) return false;
        if (e.Id == "TitleBar") return e.Name == "SOLIDWORKS" && e.Type == "ControlType.TitleBar" && !e.Offscreen && !e.Actionable;
        if (e.Id == "SystemMenuBar") return e.Name == (french ? "Barre de menus système" : "System menu bar") && e.Type == "ControlType.MenuBar" && e.Offscreen && !e.Actionable;
        if (e.Id == "Item 1") return e.Name == (french ? "Système" : "System") && e.Type == "ControlType.MenuItem" && !e.Offscreen && e.Actionable;
        return e.Id == "Close" && e.Name == (french ? "Fermer" : "Close") && e.Type == "ControlType.Button" && !e.Offscreen && e.Actionable;
    }

    /// <summary>Requires a quiet, repeatedly validated inventory before binding input or accepting an empty close boundary.</summary>
    /// <remarks>Uses caller-supplied monotonic elapsed time. This observer never delivers input or alters an existing discard batch.</remarks>
    public sealed class StableInventory
    {
        private readonly Scope scope;
        private readonly int quietMilliseconds;
        private Binding[] candidate;
        private long since, last = -1;
        private int reads;
        public bool Ready { get; private set; }
        public bool HasQuestions { get { return candidate != null && candidate.Length != 0; } }
        public StableInventory(Scope scope, int quietMilliseconds)
        {
            if (scope == null) throw new ArgumentNullException("scope");
            if (quietMilliseconds < 500 || quietMilliseconds > 5000) throw new ArgumentOutOfRangeException("quietMilliseconds");
            this.scope = scope; this.quietMilliseconds = quietMilliseconds;
        }
        /// <summary>Resets the quiet interval on any changed binding; refuses incomplete snapshots and a nonmonotonic clock.</summary>
        public bool Observe(Snapshot[] snapshots, long elapsedMilliseconds)
        {
            Ready = false;
            Binding[] current;
            try {
                if (elapsedMilliseconds < 0 || elapsedMilliseconds < last) throw new InvalidOperationException("A monotonic admission clock is required.");
                if (snapshots == null) throw new InvalidOperationException("Complete modal inventory required.");
                current = snapshots.Length == 0 ? new Binding[0] : Admit(scope, snapshots);
            } catch { Invalidate(); throw; }
            last = elapsedMilliseconds;
            if (candidate == null || current.Length != candidate.Length || current.Where((b, i) => !SameBinding(b, candidate[i])).Any())
            { candidate = current; since = elapsedMilliseconds; reads = 0; }
            reads++;
            Ready = reads >= 3 && elapsedMilliseconds - since >= quietMilliseconds;
            return Ready;
        }
        /// <summary>Breaks the quiet interval after an unavailable read, retaining monotonic-clock validation.</summary>
        public void Invalidate() { Ready = false; candidate = null; reads = 0; }
    }

    private static bool SameBinding(Binding a, Binding b)
    { return a.Dialog == b.Dialog && a.OwnerRoot == b.OwnerRoot && a.Yes == b.Yes && a.No == b.No && a.Cancel == b.Cancel &&
        a.Prompt == b.Prompt && a.YesCaption == b.YesCaption && a.NoCaption == b.NoCaption && a.CancelCaption == b.CancelCaption; }

    /// <summary>One immutable-at-dispatch request for a read-only, same-user observer on the real desktop.</summary>
    public sealed class ReaderRequest
    {
        public int Version, ParentPid, HostPid;
        public uint Thread;
        public long VbeRoot, NativeRoot;
        public string Nonce, UserSid, ParentBirthUtc, ParentImage, HostBirthUtc, HostImage;
        public string ReaderImage, Tests, TestsSha256, TestsMvid, Macro, Receipt;
    }
    /// <summary>A complete primitive snapshot or an explicit read failure; it authorizes no input by itself.</summary>
    public sealed class ReaderResponse
    {
        public int Version, ReaderPid, HostPid;
        public string Nonce, RequestSha256, TestsSha256, TestsMvid, HostBirthUtc, ReaderBirthUtc, ReaderImage;
        public bool Completed, ReadOnly, NativeInputSent, Success;
        public string ErrorKind, Error;
        public Snapshot[] Snapshots;
    }
    /// <summary>Checks the entire request before loading an accessibility provider; no arbitrary entry point is accepted.</summary>
    public static void ValidateReaderRequest(ReaderRequest r, string manifest, string image, string actualTestsHash, string actualTestsMvid)
    {
        if (r == null) throw new ArgumentNullException("request");
        Guid id;
        if (r.Version != 1 || !Guid.TryParseExact(r.Nonce, "N", out id) || string.IsNullOrWhiteSpace(r.UserSid) ||
            r.ParentPid <= 0 || r.ParentPid == r.HostPid || !UtcBirth(r.ParentBirthUtc) || !UtcBirth(r.HostBirthUtc) ||
            !Guid.TryParseExact(r.TestsMvid, "D", out id) || !HashShape(r.TestsSha256) ||
            r.TestsSha256 != actualTestsHash || r.TestsMvid != actualTestsMvid)
            throw new InvalidOperationException("Complete original observer identities and frozen tests are required.");
        new Scope(r.HostPid, r.Thread, r.VbeRoot, r.NativeRoot, r.Macro);
        foreach (string path in new[] { manifest, image, r.ReaderImage, r.ParentImage, r.HostImage, r.Tests, r.Receipt })
            if (!LocalPath(path)) throw new InvalidOperationException("Canonical local observer paths required.");
        if (!string.Equals(image, r.ReaderImage, StringComparison.OrdinalIgnoreCase) ||
            !string.Equals(Path.GetFileName(image), "powershell.exe", StringComparison.OrdinalIgnoreCase) ||
            !string.Equals(Path.GetFileName(r.ParentImage), "powershell.exe", StringComparison.OrdinalIgnoreCase) ||
            !string.Equals(Path.GetFileName(r.HostImage), "SLDWORKS.exe", StringComparison.OrdinalIgnoreCase) ||
            Path.GetFileName(r.Tests) != "VBAi.Tests.dll" || Path.GetFileName(manifest) != r.Nonce + ".request.json" ||
            Path.GetFileName(r.Receipt) != r.Nonce + ".response.json" ||
            !string.Equals(Path.GetDirectoryName(manifest), Path.GetDirectoryName(r.Receipt), StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Fixed observer entry, binaries and one-shot receipt required.");
    }
    /// <summary>Requires original observer exit and a matching terminal receipt before accepting even an empty snapshot.</summary>
    public static Snapshot[] RequireReaderTerminal(ReaderRequest request, ReaderResponse result, string requestSha256,
        int readerPid, string readerBirthUtc, bool originalExitObserved, uint exitCode)
    {
        if (request == null || result == null) throw new ArgumentNullException("Observer terminal inputs");
        if (!HashShape(requestSha256) || readerPid <= 0 || readerPid == request.ParentPid || readerPid == request.HostPid ||
            !UtcBirth(readerBirthUtc) || !originalExitObserved || !result.Completed || !result.ReadOnly || result.NativeInputSent ||
            result.Version != 1 || result.ReaderPid != readerPid || result.ReaderBirthUtc != readerBirthUtc ||
            !string.Equals(result.ReaderImage, request.ReaderImage, StringComparison.OrdinalIgnoreCase) ||
            result.Nonce != request.Nonce || result.RequestSha256 != requestSha256 ||
            result.HostPid != request.HostPid || result.HostBirthUtc != request.HostBirthUtc ||
            result.TestsSha256 != request.TestsSha256 || result.TestsMvid != request.TestsMvid)
            throw new InvalidOperationException("Original observer exit or terminal identity unproven; no input.");
        if (result.Success)
        {
            if (exitCode != 0 || result.Error != null || result.ErrorKind != null || result.Snapshots == null)
                throw new InvalidOperationException("Complete observer result and original exit disagree.");
            if (result.Snapshots.Length > 0) Admit(new Scope(request.HostPid, request.Thread, request.VbeRoot, request.NativeRoot, request.Macro), result.Snapshots);
            return result.Snapshots;
        }
        if (string.IsNullOrWhiteSpace(result.Error) || result.Snapshots != null)
            throw new InvalidOperationException("A failed read must never publish partial snapshots.");
        if (exitCode == 2 && result.ErrorKind == "NotReady") throw new NotReadyException(result.Error);
        if (exitCode == 2 && result.ErrorKind == "ObservationUnavailable") throw new ObservationUnavailableException(result.Error);
        if (exitCode == 1 && (result.ErrorKind == "Rejected" || result.ErrorKind == "RuntimeError"))
            throw new InvalidOperationException("Owned observer rejected the read: " + result.Error);
        throw new InvalidOperationException("Observer failure and original exit disagree.");
    }
    private static bool LocalPath(string value)
    { return !string.IsNullOrWhiteSpace(value) && Path.IsPathRooted(value) && !value.StartsWith("\\\\", StringComparison.Ordinal) && string.Equals(Path.GetFullPath(value), value, StringComparison.OrdinalIgnoreCase); }
    private static bool HashShape(string value)
    { return value != null && value.Length == 64 && value.All(Uri.IsHexDigit) && value == value.ToUpperInvariant(); }
    private static bool UtcBirth(string value)
    { DateTime parsed; return DateTime.TryParseExact(value, "o", CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out parsed) && parsed.Kind == DateTimeKind.Utc; }
    /// <summary>Reads birth through the retained creation handle, including after exit; a PID lookup cannot replace this proof.</summary>
    public static string OriginalReaderBirthUtc(long processHandle)
    {
        long birth, exit, kernel, user;
        if (processHandle <= 0 || !GetProcessTimes(new IntPtr(processHandle), out birth, out exit, out kernel, out user))
            throw new InvalidOperationException("Original reader creation time unavailable.");
        return DateTime.FromFileTimeUtc(birth).ToString("o", CultureInfo.InvariantCulture);
    }

    /// <summary>One delivery per originally captured question; a pending or failed attempt cannot be retried.</summary>
    public sealed class DiscardBatch
    {
        private readonly Scope scope;
        private readonly Binding[] original;
        private readonly Func<Snapshot[]> read;
        private readonly Action context;
        private readonly Func<long, bool> post;
        private int settled;
        private bool entered, delivered;
        public int Settled { get { return settled; } }
        public bool Complete { get { return settled == original.Length; } }
        /// <summary>True only after delivery of the last captured response, when original normal process exit can settle it.</summary>
        public bool LastOriginalQuestionPending { get { return entered && delivered && settled == original.Length - 1; } }
        public long PendingDialog { get { return entered && settled < original.Length ? original[settled].Dialog : 0; } }
        /// <summary>Binds native reads and delivery while retaining the caller's live ownership and desktop guard.</summary>
        public DiscardBatch(Scope scope, Snapshot[] approved, Action context)
            : this(scope, approved, () => Observe(scope), context) { }
        /// <summary>Uses an owned, exited observation worker; only the parent can deliver the fully guarded native response.</summary>
        public DiscardBatch(Scope scope, Snapshot[] approved, Func<Snapshot[]> read, Action context)
            : this(scope, approved, read, context, handle => true)
        {
            post = handle => {
                context(); Require(handle, scope);
                var window = new IntPtr(handle);
                if (handle != original[settled].No || Class(handle) != "Button" || GetDlgCtrlID(window) != 0 ||
                    !IsWindowVisible(window) || !IsWindowEnabled(window) ||
                    GetAncestor(window, 2).ToInt64() != original[settled].Dialog || Text(handle) != original[settled].NoCaption)
                    throw new InvalidOperationException("Original native No changed immediately before delivery.");
                return PostMessage(window, 0xF5, UIntPtr.Zero, IntPtr.Zero);
            };
        }
        public DiscardBatch(Scope scope, Snapshot[] approved, Func<Snapshot[]> read, Action context, Func<long, bool> post)
        {
            if (read == null || context == null || post == null) throw new ArgumentNullException("Discard boundaries");
            this.scope = scope; original = Admit(scope, approved); this.read = read; this.context = context; this.post = post;
        }
        /// <summary>Checks two fresh complete reads and the caller's scope before requesting the next captured No.</summary>
        public void RequestNext()
        {
            if (Complete || entered) throw new InvalidOperationException("This discard is complete or already entered; no retry.");
            entered = true;
            context(); RequireRemaining(read(), settled);
            context(); RequireRemaining(read(), settled);
            context();
            delivered = post(original[settled].No);
            if (!delivered) throw new InvalidOperationException("Discard enqueue failed; outcome unproven, no retry.");
        }
        /// <summary>Observes only; advancement requires disappearance of the addressed modal with every other binding unchanged.</summary>
        public bool ObserveCompletion()
        {
            if (!entered || !delivered) throw new InvalidOperationException("There is no successfully queued original discard to observe.");
            context(); var current = read();
            if (current == null) throw new InvalidOperationException("Complete owned modal inventory required.");
            if (current.Any(s => s != null && s.Dialog == original[settled].Dialog))
            { RequireRemaining(current, settled); return false; }
            RequireRemaining(current, settled + 1);
            settled++; entered = false; delivered = false; return true;
        }
        /// <summary>Settles the last queued question only when the caller proves normal exit through its original process handle and pending ExitApp result.</summary>
        /// <remarks>No native read is possible after exit. This observation never requests another input or retries ExitApp.</remarks>
        public bool ObserveOriginalOwnerExit(Func<bool> originalExitProof)
        {
            if (originalExitProof == null) throw new ArgumentNullException("originalExitProof");
            if (!entered || !delivered || settled != original.Length - 1)
                throw new InvalidOperationException("Only the last successfully queued original question can settle by original owner exit.");
            if (!originalExitProof()) return false;
            settled++; entered = false; delivered = false; return true;
        }
        private void RequireRemaining(Snapshot[] snapshots, int start)
        {
            if (start == original.Length)
            {
                if (snapshots == null || snapshots.Length != 0) throw new InvalidOperationException("A new modal appeared after discard.");
                return;
            }
            var current = Admit(scope, snapshots);
            if (current.Length != original.Length - start) throw new InvalidOperationException("Original modal set changed; no replacement.");
            for (int i = 0; i < current.Length; i++)
            {
                var a = original[start + i]; var b = current[i];
                if (a.Dialog != b.Dialog || a.OwnerRoot != b.OwnerRoot || a.Yes != b.Yes || a.No != b.No || a.Cancel != b.Cancel ||
                    a.Prompt != b.Prompt || a.YesCaption != b.YesCaption || a.NoCaption != b.NoCaption || a.CancelCaption != b.CancelCaption)
                    throw new InvalidOperationException("Original save question or command binding changed; no delivery.");
            }
        }
    }

    /// <summary>Reads the owned desktop, physical command HWNDs and fresh UI Automation nodes without selecting or focusing a control.</summary>
    public static Snapshot[] Observe(Scope scope)
    {
        if (scope == null) throw new ArgumentNullException("scope");
        if (!IsWindow(new IntPtr(scope.NativeRoot))) throw new NotReadyException("Original root is absent; no input can be admitted while its owner's exit is observed.");
        Require(scope.NativeRoot, scope);
        if (GetAncestor(new IntPtr(scope.NativeRoot), 2).ToInt64() != scope.NativeRoot || !Class(scope.NativeRoot).StartsWith("Afx:", StringComparison.Ordinal))
            throw new InvalidOperationException("Original native root identity differs.");
        var dialogs = new List<long>(); Exception failure = null; int count = 0;
        bool complete = EnumDesktopWindows(GetThreadDesktop(GetCurrentThreadId()), (w, state) => {
            try {
                if (++count > 8192) throw new InvalidOperationException("Desktop inventory bound exceeded.");
                uint pid; uint thread = GetWindowThreadProcessId(w, out pid);
                if (pid == scope.Pid && IsWindowVisible(w) && Class(w.ToInt64()) == "#32770")
                { if (thread != scope.Thread) throw new InvalidOperationException("Foreign owned modal thread."); dialogs.Add(w.ToInt64()); }
                return true;
            } catch (Exception error) { failure = error; return false; }
        }, IntPtr.Zero);
        if (failure != null) throw failure;
        if (!complete || dialogs.Count > 2) throw new InvalidOperationException("Complete bounded owned modal set required.");
        var snapshots = new List<Snapshot>();
        foreach (long dialog in dialogs)
        {
            Require(dialog, scope);
            long owner = GetAncestor(new IntPtr(dialog), 3).ToInt64(); Require(owner, scope);
            if (owner == scope.VbeRoot && Class(owner) != "wndclass_desked_gsk") throw new InvalidOperationException("Original VBE root identity differs.");
            var buttons = new List<Button>(); int children = 0, direct = 0; failure = null;
            complete = EnumChildWindows(new IntPtr(dialog), (w, state) => {
                try {
                    long handle = w.ToInt64(); Require(handle, scope);
                    if (++children > 64) throw new InvalidOperationException("Native control inventory bound exceeded.");
                    string cls = Class(handle);
                    if (cls == "DirectUIHWND") direct++;
                    if (IsWindowVisible(w) && cls == "Button") buttons.Add(new Button { Handle = handle, Root = GetAncestor(w, 2).ToInt64(),
                        Pid = scope.Pid, Thread = scope.Thread, Id = GetDlgCtrlID(w), Class = cls, Caption = Text(handle), Visible = true, Enabled = IsWindowEnabled(w) });
                    else if (IsWindowVisible(w) && cls != "DirectUIHWND" && cls != "CtrlNotifySink" && cls != "ScrollBar" && cls != "SysLink")
                        throw new InvalidOperationException("Unknown visible task-dialog control.");
                    return true;
                } catch (Exception error) { failure = error; return false; }
            }, IntPtr.Zero);
            if (failure != null) throw failure;
            if (!complete || direct != 1) throw new InvalidOperationException("Complete DirectUI layout required; no generic alert response.");
            if (buttons.Count != 3) throw new NotReadyException("The observed DirectUI command set is not yet complete; no input.");
            AutomationElementCollection nodes;
            try { nodes = AutomationElement.FromHandle(new IntPtr(dialog)).FindAll(TreeScope.Descendants, Condition.TrueCondition); }
            catch (ElementNotAvailableException) { throw new ObservationUnavailableException("The captured accessible dialog disappeared during observation."); }
            if (nodes.Count == 0) throw new NotReadyException("DirectUI accessibility is not yet ready; no input.");
            if (nodes.Count > 128) throw new InvalidOperationException("Complete bounded accessible inventory required.");
            var elements = new List<Element>();
            foreach (AutomationElement node in nodes)
            {
                AutomationElement.AutomationElementInformation c; AutomationPattern[] patterns;
                try { c = node.Current; patterns = node.GetSupportedPatterns(); }
                catch (ElementNotAvailableException) { throw new ObservationUnavailableException("A captured accessible control disappeared during observation."); }
                if ((c.Name ?? "").Length > 2048 || (c.AutomationId ?? "").Length > 128 || c.ControlType.ProgrammaticName.Length > 64)
                    throw new InvalidOperationException("Accessible text exceeds the complete snapshot bounds.");
                elements.Add(new Element { Id = c.AutomationId, Name = c.Name, Type = c.ControlType.ProgrammaticName,
                    Handle = unchecked((uint)c.NativeWindowHandle), Pid = c.ProcessId, Enabled = c.IsEnabled, Offscreen = c.IsOffscreen,
                    Actionable = patterns.Contains(InvokePattern.Pattern) || patterns.Contains(TogglePattern.Pattern) ||
                        patterns.Contains(SelectionItemPattern.Pattern) || patterns.Contains(ExpandCollapsePattern.Pattern) || patterns.Contains(ValuePattern.Pattern) });
            }
            if (elements.Count(e => e.Id == "ContentText") != 1 || elements.Count(e => e.Id != null && e.Id.StartsWith("CommandButton_", StringComparison.Ordinal)) != 3)
                throw new NotReadyException("The observed accessible question is still constructing; no input.");
            Require(dialog, scope);
            snapshots.Add(new Snapshot { Dialog = dialog, OwnerRoot = owner, Pid = scope.Pid, Thread = scope.Thread,
                Class = Class(dialog), Caption = Text(dialog), Visible = IsWindowVisible(new IntPtr(dialog)), Enabled = IsWindowEnabled(new IntPtr(dialog)),
                Elements = elements.ToArray(), Buttons = buttons.ToArray() });
        }
        return snapshots.ToArray();
    }
    private static void Require(long handle, Scope scope)
    {
        if (handle <= 0) throw new InvalidOperationException("Owned HWND identity differs.");
        if (!IsWindow(new IntPtr(handle))) throw new ObservationUnavailableException("A captured HWND disappeared during observation.");
        uint pid; uint thread = GetWindowThreadProcessId(new IntPtr(handle), out pid);
        if (pid != scope.Pid || thread != scope.Thread)
            throw new InvalidOperationException("Owned HWND identity differs.");
    }
    private static string Class(long handle)
    { var value = new StringBuilder(256); if (GetClassName(new IntPtr(handle), value, value.Capacity) == 0) throw new ObservationUnavailableException("Window class unavailable during observation."); return value.ToString(); }
    private static string Text(long handle)
    { var value = new StringBuilder(2048); UIntPtr length; if (SendMessageTimeout(new IntPtr(handle), 13, new UIntPtr((uint)value.Capacity), value, 0x23, 100, out length) == IntPtr.Zero) throw new ObservationUnavailableException("Native text is not currently readable."); if (length.ToUInt64() >= 2047) throw new InvalidOperationException("Complete bounded native text required."); return value.ToString(); }
    private delegate bool Visitor(IntPtr window, IntPtr state);
    [DllImport("user32.dll", SetLastError = true)] private static extern bool EnumDesktopWindows(IntPtr desktop, Visitor visitor, IntPtr state);
    [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr root, Visitor visitor, IntPtr state);
    [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr window, uint flags);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
    [DllImport("user32.dll")] private static extern IntPtr GetThreadDesktop(uint thread);
    [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
    [DllImport("kernel32.dll", SetLastError = true)] private static extern bool GetProcessTimes(IntPtr process, out long birth, out long exit, out long kernel, out long user);
    [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr window);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
    [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr window);
    [DllImport("user32.dll")] private static extern int GetDlgCtrlID(IntPtr window);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder text, int size);
    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern IntPtr SendMessageTimeout(IntPtr window, uint message, UIntPtr first, StringBuilder text, uint flags, uint timeout, out UIntPtr result);
    [DllImport("user32.dll", SetLastError = true)] private static extern bool PostMessage(IntPtr window, uint message, UIntPtr first, IntPtr second);
}
