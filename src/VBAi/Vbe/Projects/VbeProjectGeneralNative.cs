using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace VBAi
{
    // General controls are the observed VBIDE dialog contract, never a host-specific Application API.
    /// <summary>Reads and updates the existing VBIDE Project Properties General page through verified controls on the original VBE UI thread.</summary>
    internal sealed class VbeProjectGeneralNative : VbeProjectGeneralOperation.INative
    {

        /// <summary>Original VBE root window; every dialog and field operation must remain in the root owner's process and UI thread.</summary>
        private readonly IntPtr root;

        /// <summary>Process and UI-thread IDs captured when this wrapper is created; later calls reject a different owner.</summary>
        private readonly uint pid, thread;

        /// <summary>Additional caller-supplied check for private/input/sentinel context, run after the native owner checks.</summary>
        private readonly Action requireNativeContext;

        /// <summary>Canonical project name recorded only after a successful General-page snapshot; used to revalidate the same dialog.</summary>
        private string projectName;

        /// <summary>One-shot guards: a field setter and a close command can each be attempted at most once per wrapper.</summary>
        private bool fieldConsumed, closeConsumed;

        /// <summary>Binds the wrapper to the current x64 VBE UI thread and validates the supplied root window.</summary>
        /// <param name="originalVbeRoot">Root VBE window; it must be a top-level window owned by the current process and UI thread.</param>
        /// <param name="requireNativeContext">Additional caller check for qualification-specific context; null is rejected.</param>
        internal VbeProjectGeneralNative(IntPtr originalVbeRoot, Action requireNativeContext)
        {
            root = originalVbeRoot; pid = GetCurrentProcessId(); thread = GetCurrentThreadId();
            this.requireNativeContext = requireNativeContext ?? throw new ArgumentNullException(nameof(requireNativeContext));
            RequireOwner();
        }

        /// <summary>Rejects calls unless this is the original x64 STA thread and the root still belongs to its captured process and thread.</summary>
        public void RequireOwner()
        {
            uint owner;
            Require(IntPtr.Size == 8 && Thread.CurrentThread.GetApartmentState() == ApartmentState.STA &&
                GetCurrentProcessId() == pid && GetCurrentThreadId() == thread && root != IntPtr.Zero &&
                GetWindowThreadProcessId(root, out owner) == thread && owner == pid && GetAncestor(root, 2) == root,
                "General must remain on its original current-process x64 VBE UI STA/root.");
            requireNativeContext(); // Caller can additionally bind private/input/sentinel proof for qualification.
        }

        /// <summary>Requires an enabled VBE root and no existing owned modal before starting a General-page operation.</summary>
        public void Prepare()
        {
            RequireOwner(); Require(IsWindowEnabled(root) && Dialogs().Count == 0, "Existing owned modal/disabled VBE refuses a new General operation.");
        }

        /// <summary>Enumerates visible top-level windows owned by the captured process and accepts only known modal dialog classes on the captured UI thread.</summary>
        /// <returns>Owned modal window handles found during the bounded enumeration.</returns>
        private List<IntPtr> Dialogs()
        {
            RequireOwner(); var found = new List<IntPtr>(); var watch = Stopwatch.StartNew(); int visited = 0; Exception failure = null;
            bool complete = EnumWindows((window, state) => {
                try
                {
                    if (++visited > 8192 || watch.ElapsedMilliseconds > 5000) return false;
                    uint owner; uint tid = GetWindowThreadProcessId(window, out owner);
                    if (owner == pid && IsWindowVisible(window))
                    {
                        string kind = Class(window);
                        if (kind == "#32770" || kind.StartsWith("bosa_sdm", StringComparison.OrdinalIgnoreCase))
                        {
                            Require(tid == thread && kind == "#32770", "Unknown owned modal/thread refuses General automation."); found.Add(window);
                        }
                    }
                    return true;
                }
                catch (Exception error) { failure = error; return false; }
            }, IntPtr.Zero);
            if (failure != null) throw failure;
            Require(complete && visited <= 8192 && watch.ElapsedMilliseconds <= 5000, "Owned modal inventory is incomplete."); RequireOwner(); return found;
        }

        /// <summary>Rejects a zero or stale target and any window not owned by the captured process and VBE UI thread.</summary>
        /// <param name="window">Candidate HWND to validate against the captured process and UI thread.</param>
        private void RequireWindow(IntPtr window)
        {
            RequireOwner(); uint owner;
            Require(window != IntPtr.Zero && GetWindowThreadProcessId(window, out owner) == thread && owner == pid,
                "General native target changed PID/UI thread.");
        }

        /// <summary>Checks the shared snapshot stopwatch and owner context against the five-second cooperative budget.</summary>
        /// <param name="watch">Shared stopwatch for the current bounded UI operation.</param>
        private void Budget(Stopwatch watch) { RequireOwner(); Require(watch.ElapsedMilliseconds <= 5000, "General snapshot exceeded its cooperative five-second budget."); }

        /// <summary>Reads a window caption with a bounded buffer and rechecks its owner before and after the native call.</summary>
        /// <param name="window">Owned window whose Unicode caption is read.</param>
        /// <param name="watch">Shared stopwatch for the current bounded UI operation.</param>
        /// <returns>Exact window text; unavailable or truncated text causes the operation to fail.</returns>
        private string Text(IntPtr window, Stopwatch watch)
        {
            Budget(watch); RequireWindow(window);
            var text = new StringBuilder(4096); UIntPtr result;
            // This is an owning-thread trusted native getter; SendMessageTimeout is not a hard wall timeout on that same thread.
            Require(SendMessageTimeoutText(window, 13, new UIntPtr(4096), text, 0x23, 150, out result) != IntPtr.Zero && result.ToUInt64() < 4095,
                "General text getter is unavailable/truncated.");
            Budget(watch); RequireWindow(window); return text.ToString();
        }

        /// <summary>Requires the unique visible owned Project Properties dialog, its original-root ownership, and an exact English or French title.</summary>
        /// <param name="dialog">Candidate Project Properties dialog to validate.</param>
        /// <param name="name">Exact VBIDE project name expected in the dialog title.</param>
        /// <param name="watch">Shared stopwatch for the current bounded UI operation.</param>
        private void RequireDialog(IntPtr dialog, string name, Stopwatch watch)
        {
            Budget(watch); RequireWindow(dialog);
            Require(Class(dialog) == "#32770" && IsWindowVisible(dialog) && Dialogs().SequenceEqual(new[] { dialog }), "General modal is ambiguous or changed.");
            IntPtr owner = GetWindow(dialog, 4);
            Require(owner == root || GetAncestor(owner, 2) == root, "General is not owned by the original VBE root.");
            string caption = Text(dialog, watch);
            Require(caption == name + " - Propriétés du projet" || caption == name + " - Project Properties", "General exact project dialog title changed.");
            Budget(watch);
        }

        /// <summary>Captures the existing General dialog and its visible project fields without switching tabs or opening a dialog.</summary>
        /// <param name="exactProjectName">Canonical project name used to match the unique owned Project Properties dialog and verify its Name field.</param>
        /// <returns>A snapshot of the exact page, tab, edit controls, and their current text, or <see langword="null"/> when no owned modal is open.</returns>
        public VbeProjectGeneralOperation.Snapshot Capture(string exactProjectName)
        {
            RequireOwner(); var watch = Stopwatch.StartNew(); var dialogs = Dialogs(); Budget(watch);
            if (dialogs.Count == 0) return null;
            Require(dialogs.Count == 1, "Multiple owned modals refuse General."); IntPtr dialog = dialogs[0]; RequireDialog(dialog, exactProjectName, watch);
            var children = EnumerateChildren(dialog, watch);
            var snapshot = new VbeProjectGeneralOperation.Snapshot
            {
                Dialog = dialog,
                Tab = FindSelectedGeneralTab(dialog, children),
                NameEdit = FindVisibleEdit(children, 4941, watch),
                DescriptionEdit = FindVisibleEdit(children, 4940, watch),
                HelpFileEdit = FindVisibleEdit(children, 4948, watch),
                Context = FindVisibleEdit(children, 4949, watch),
                CompilationEdit = FindVisibleEdit(children, 4958, watch)
            };
            snapshot.Page = GetParent(snapshot.Context); RequireWindow(snapshot.Page);
            Require(Class(snapshot.Page) == "#32770" && GetDlgCtrlID(snapshot.Page) == 0 && GetParent(snapshot.Page) == dialog && IsWindowVisible(snapshot.Page), "General child page identity changed.");
            foreach (IntPtr window in new[] { snapshot.NameEdit, snapshot.DescriptionEdit, snapshot.HelpFileEdit, snapshot.Context, snapshot.CompilationEdit })
                Require(GetParent(window) == snapshot.Page && GetAncestor(window, 2) == dialog, "General edits belong to another native page.");
            ReadSnapshotFields(snapshot, watch);
            Require(snapshot.Name == exactProjectName, "General native project name differs from the canonical approved project.");
            RequireDialog(dialog, exactProjectName, watch); Budget(watch); projectName = exactProjectName; return snapshot;
        }

        /// <summary>Enumerates the dialog's child windows within the cooperative time and count limits.</summary>
        /// <param name="dialog">Unique owned Project Properties dialog previously checked by <see cref="RequireDialog"/>.</param>
        /// <param name="watch">Stopwatch shared by the complete snapshot operation.</param>
        /// <returns>Child handles that remain owned by the original VBE UI thread.</returns>
        private List<IntPtr> EnumerateChildren(IntPtr dialog, Stopwatch watch)
        {
            var children = new List<IntPtr>();
            bool exceeded = false;
            Exception childFailure = null;
            EnumChildWindows(dialog, (child, state) =>
            {
                try
                {
                    Budget(watch);
                    if (children.Count >= 1024)
                    {
                        exceeded = true;
                        return false;
                    }
                    RequireWindow(child);
                    children.Add(child);
                    return true;
                }
                catch (Exception error)
                {
                    childFailure = error;
                    return false;
                }
            }, IntPtr.Zero);
            if (childFailure != null) throw childFailure;
            Require(!exceeded && children.Count != 0, "General child inventory is incomplete.");
            return children;
        }

        /// <summary>Resolves one visible, enabled General edit with the expected control ID and editable style.</summary>
        /// <param name="children">Previously enumerated child windows of the unique General dialog.</param>
        /// <param name="id">Expected VBIDE control ID for the field.</param>
        /// <param name="watch">Stopwatch shared by the complete snapshot operation.</param>
        /// <returns>The unique matching edit handle on the original VBE UI thread.</returns>
        private IntPtr FindVisibleEdit(IList<IntPtr> children, int id, Stopwatch watch)
        {
            Budget(watch);
            var matches = children.Where(child => GetDlgCtrlID(child) == id
                && Class(child) == "Edit" && IsWindowVisible(child)).ToArray();
            Require(matches.Length == 1, "General expected visible edit is missing or ambiguous.");
            var edit = matches[0];
            RequireWindow(edit);
            Require(MatchesEditableContract(GetDlgCtrlID(edit), id, Class(edit),
                GetWindowLongPtr(edit, -16).ToInt64(), IsWindowVisible(edit), IsWindowEnabled(edit)),
                "General edit is disabled/password/readonly.");
            return edit;
        }

        /// <summary>Finds the sole enabled General tab and confirms tab zero is already selected.</summary>
        /// <param name="dialog">Unique owned Project Properties dialog.</param>
        /// <param name="children">Previously enumerated dialog children.</param>
        /// <returns>The original tab handle; this method never changes the selected tab.</returns>
        private IntPtr FindSelectedGeneralTab(IntPtr dialog, IList<IntPtr> children)
        {
            var tabs = children.Where(child => Class(child) == "SysTabControl32"
                && GetParent(child) == dialog && IsWindowVisible(child) && IsWindowEnabled(child)).ToArray();
            Require(tabs.Length == 1, "General exact native tab is unavailable.");
            RequireWindow(tabs[0]);
            UIntPtr selected;
            Require(SendMessageTimeoutScalar(tabs[0], 0x130B, UIntPtr.Zero, IntPtr.Zero,
                0x23, 150, out selected) != IntPtr.Zero && selected.ToUInt64() == 0,
                "General tab0 must already be selected; tab changes are forbidden.");
            return tabs[0];
        }

        /// <summary>Reads the five text values from the edit handles already bound to the captured General page.</summary>
        /// <param name="snapshot">Snapshot whose page and edit ancestry were verified by <see cref="Capture"/>.</param>
        /// <param name="watch">Stopwatch shared by the complete snapshot operation.</param>
        private void ReadSnapshotFields(VbeProjectGeneralOperation.Snapshot snapshot, Stopwatch watch)
        {
            snapshot.Name = Text(snapshot.NameEdit, watch);
            snapshot.Description = Text(snapshot.DescriptionEdit, watch);
            snapshot.HelpFile = Text(snapshot.HelpFileEdit, watch);
            snapshot.ContextText = Text(snapshot.Context, watch);
            snapshot.Compilation = Text(snapshot.CompilationEdit, watch);
        }

        /// <summary>Re-captures the dialog and requires its native control identities and, optionally, all captured text to match the expected snapshot.</summary>
        /// <param name="expected">Previously captured snapshot used as the concurrency token.</param>
        /// <param name="compareValues">When true, also compares the options version derived from the captured field values.</param>
        public void RequireSame(VbeProjectGeneralOperation.Snapshot expected, bool compareValues)
        {
            Require(expected != null && projectName != null, "General snapshot was not captured.");
            var current = Capture(projectName);
            Require(expected.SameNative(current) && (!compareValues || expected.OptionsVersion == current.OptionsVersion), "Original General native handles/values changed.");
        }

        /// <summary>Writes one HelpContextID value through the captured General edit after the final authorization callback.</summary>
        /// <param name="expected">Current snapshot whose dialog and field identity must still match.</param>
        /// <param name="value">HelpContextID value formatted with invariant decimal digits.</param>
        /// <param name="beforeEntry">Final authorization check run immediately before the single native text mutation.</param>
        public void WriteContext(VbeProjectGeneralOperation.Snapshot expected, int value, Action beforeEntry)
        { WriteField(expected, expected.Context, value.ToString(CultureInfo.InvariantCulture), beforeEntry); }

        /// <summary>Verifies the current snapshot and proves the requested HelpFile round-trips through the target edit's text encoding.</summary>
        /// <param name="expected">Snapshot whose native controls and current option values must remain unchanged.</param>
        /// <param name="value">HelpFile text to encode and decode without best-fit substitution or normalization.</param>
        public void RequireHelpFileRepresentable(VbeProjectGeneralOperation.Snapshot expected, string value)
        {
            RequireSame(expected, true); RequireTextRepresentable(expected.HelpFileEdit, value);
        }

        /// <summary>Checks that text can be represented exactly by the target edit's Unicode or active ANSI code-page contract.</summary>
        /// <param name="target">Original-process edit control whose encoding mode is inspected.</param>
        /// <param name="value">Text to round-trip without best-fit replacement.</param>
        private void RequireTextRepresentable(IntPtr target, string value)
        {
            RequireWindow(target); bool unicode = IsWindowUnicode(target); uint codePage = GetACP();
            RequireWindow(target);
            RequireExactTextRepresentation(value, checked((int)codePage), unicode);
            RequireWindow(target);
        }

        /// <summary>Refuses unrepresentable text rather than changing it through replacement, best-fit conversion, or normalization.</summary>
        /// <param name="value">Text whose exact round-trip representation is required.</param>
        /// <param name="codePage">Active Windows ANSI code page; must be positive even when the control is Unicode.</param>
        /// <param name="unicode">Whether the edit control uses UTF-16 instead of the supplied ANSI code page.</param>
        internal static void RequireExactTextRepresentation(string value, int codePage, bool unicode)
        {
            if (value == null) throw new ArgumentNullException(nameof(value));
            Require(codePage > 0, "Native General text code page is unavailable.");
            // Unsupported/unknown code pages are infrastructure failures, not a
            // known representation refusal eligible for a guarded Cancel.
            var encoding = Encoding.GetEncoding(unicode ? 1200 : codePage,
                EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);
            try
            {
                // Exception fallbacks disable replacement and best-fit mappings. Unicode controls
                // still require valid UTF-16; no normalization or path substitution is performed.
                string roundTrip = encoding.GetString(encoding.GetBytes(value));
                if (!string.Equals(value, roundTrip, StringComparison.Ordinal))
                    throw new VbeProjectGeneralOperation.TextRepresentationRefusedException();
            }
            catch (EncoderFallbackException error) { throw new VbeProjectGeneralOperation.TextRepresentationRefusedException(error); }
            catch (DecoderFallbackException error) { throw new VbeProjectGeneralOperation.TextRepresentationRefusedException(error); }
        }

        /// <summary>Writes the HelpFile field once after checking the captured snapshot and final authorization.</summary>
        /// <param name="expected">Snapshot whose target dialog and option values must remain current.</param>
        /// <param name="value">New HelpFile text; null is refused before native dispatch.</param>
        /// <param name="beforeEntry">Final authorization callback run directly before the native edit message.</param>
        public void WriteHelpFile(VbeProjectGeneralOperation.Snapshot expected, string value, Action beforeEntry)
        { WriteField(expected, expected.HelpFileEdit, value ?? throw new ArgumentNullException(nameof(value)), beforeEntry); }

        /// <summary>Consumes the one permitted field-write attempt, revalidates the dialog and text encoding, then sends one native setter message.</summary>
        /// <param name="expected">Snapshot used to reject stale native handles or changed project options.</param>
        /// <param name="target">Exact captured edit control for HelpFile or HelpContextID.</param>
        /// <param name="value">Exact text to set after representation checks.</param>
        /// <param name="beforeEntry">Required final authorization callback; a timeout or uncertain setter result is never retried.</param>
        private void WriteField(VbeProjectGeneralOperation.Snapshot expected, IntPtr target, string value, Action beforeEntry)
        {
            Require(!fieldConsumed, "Original General field write is already consumed."); fieldConsumed = true;
            RequireSame(expected, true); RequireOwner(); UIntPtr result;
            RequireTextRepresentable(target, value);
            Require(beforeEntry != null, "Final General field authorization is required."); beforeEntry(); RequireOwner();
            Require(SendMessageTimeoutWrite(target, 12, UIntPtr.Zero, value, 0x23, 150, out result) != IntPtr.Zero && result.ToUInt64() != 0,
                "The single General field write is uncertain; no alternative write or retry.");
        }

        /// <summary>Posts exactly one verified OK or Cancel command to the captured dialog after a final identity and authorization check.</summary>
        /// <param name="expected">Snapshot identifying the original dialog and controls.</param>
        /// <param name="buttonId">Native button ID 1 for OK or 2 for Cancel.</param>
        /// <param name="beforeEnqueue">Final authorization callback invoked immediately before posting the close command.</param>
        public void Close(VbeProjectGeneralOperation.Snapshot expected, int buttonId, Action beforeEnqueue)
        {
            Require((buttonId == 1 || buttonId == 2) && !closeConsumed, "Only one original General OK/Cancel is allowed."); closeConsumed = true;
            RequireSame(expected, true); var watch = Stopwatch.StartNew(); IntPtr button = GetDlgItem(expected.Dialog, buttonId); RequireWindow(button);
            Require(Class(button) == "Button" && GetParent(button) == expected.Dialog && GetDlgCtrlID(button) == buttonId && IsWindowVisible(button) && IsWindowEnabled(button), "General exact close button changed.");
            string caption = Text(button, watch);
            Require(MatchesCloseCaption(buttonId, caption), "General close caption is unsupported.");
            RequireDialog(expected.Dialog, projectName, watch); RequireWindow(button);
            Require(GetDlgItem(expected.Dialog, buttonId) == button && GetParent(button) == expected.Dialog && GetDlgCtrlID(button) == buttonId && Class(button) == "Button" && IsWindowVisible(button) && IsWindowEnabled(button), "General final close identity changed.");
            Budget(watch); Require(beforeEnqueue != null, "Final General close authorization is required."); beforeEnqueue(); RequireOwner();
            Require(PostMessage(expected.Dialog, 0x111, new UIntPtr((uint)buttonId), button), "General close enqueue failed; no retry.");
        }

        /// <summary>Verifies that the original dialog is gone, no owned modal remains, and the VBE root is enabled.</summary>
        /// <param name="expected">Snapshot identifying the dialog whose closure is being checked.</param>
        /// <returns>True when closure is proved; false while the original dialog remains open.</returns>
        public bool Closed(VbeProjectGeneralOperation.Snapshot expected)
        {
            RequireOwner(); if (IsWindow(expected.Dialog)) { RequireSame(expected, false); return false; }
            Require(Dialogs().Count == 0 && IsWindowEnabled(root), "General closure/root enabled state is unproved."); return true;
        }

        /// <summary>Throws when a required native or authorization condition is false.</summary>
        /// <param name="value">Condition that must hold; false throws InvalidOperationException with the supplied reason.</param>
        /// <param name="reason">Failure explanation included in the InvalidOperationException.</param>
        private static void Require(bool value, string reason) { if (!value) throw new InvalidOperationException(reason); }

        /// <summary>Reports whether a control has the requested ID, is a visible enabled Edit, and has neither password nor read-only style.</summary>
        /// <param name="id">Control ID reported by the candidate window.</param>
        /// <param name="expectedId">VBIDE control ID required for the requested General field.</param>
        /// <param name="kind">Native class name; the editable contract requires <c>Edit</c>.</param>
        /// <param name="style">Window style bits checked for password and read-only flags.</param>
        /// <param name="visible">Whether the candidate control is visible.</param>
        /// <param name="enabled">Whether the candidate control accepts input.</param>
        /// <returns><see langword="true"/> only when ID, class, visibility, enabled state, and editable style all match.</returns>
        internal static bool MatchesEditableContract(int id, int expectedId, string kind, long style, bool visible, bool enabled)
            => id == expectedId && kind == "Edit" && visible && enabled && (style & (0x20 | 0x800)) == 0;

        /// <summary>Accepts only the supported localized caption for the requested OK or Cancel button.</summary>
        /// <param name="buttonId">Dialog command ID: 1 for OK or 2 for Cancel.</param>
        /// <param name="caption">Exact caption read from the native button.</param>
        /// <returns><see langword="true"/> when the caption matches one of the supported English or French forms for that ID.</returns>
        internal static bool MatchesCloseCaption(int buttonId, string caption)
            => buttonId == 1 ? caption == "OK" || caption == "&OK" : buttonId == 2 && (caption == "Annuler" || caption == "&Annuler" || caption == "Cancel" || caption == "&Cancel");

        /// <summary>Reads a window's Win32 class name, failing if the native API returns no class.</summary>
        /// <param name="window">Window whose native class name is queried.</param>
        /// <returns>Class name copied into the bounded 256-character buffer.</returns>
        private static string Class(IntPtr window) { var text = new StringBuilder(256); Require(GetClassName(window, text, text.Capacity) > 0, "Native class unavailable."); return text.ToString(); }

        /// <summary>Defines the visitor callback.</summary>
        /// <param name="window">Window visited by EnumWindows or EnumChildWindows.</param>
        /// <param name="state">Caller context passed through the Win32 enumeration API.</param>
        /// <returns><see langword="true"/> to continue enumeration; <see langword="false"/> to stop it.</returns>
        private delegate bool Visitor(IntPtr window, IntPtr state);

        /// <summary>Gets the process ID of the process running this wrapper.</summary>
        /// <returns>Current process ID.</returns>
        [DllImport("kernel32.dll")] private static extern uint GetCurrentProcessId();

        /// <summary>Gets the native ID of the calling VBE UI thread.</summary>
        /// <returns>Current native thread ID.</returns>
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();

        /// <summary>Gets the active Windows ANSI code page used to check legacy edit controls.</summary>
        /// <returns>Active ANSI code-page identifier.</returns>
        [DllImport("kernel32.dll")] private static extern uint GetACP();

        /// <summary>Gets the owning thread ID for a window and writes its process ID to the out parameter.</summary>
        /// <param name="window">Window whose owner is queried.</param>
        /// <param name="pid">Receives the owning process ID.</param>
        /// <returns>Owning thread ID, or zero when the window has no owner thread.</returns>
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);

        /// <summary>Gets a window ancestor, including the root ancestor used to bind controls to the original VBE.</summary>
        /// <param name="window">Window whose ancestor is requested.</param>
        /// <param name="flags">Win32 ancestor selector; callers pass GA_ROOT (2).</param>
        /// <returns>Ancestor HWND, or zero when no matching ancestor exists.</returns>
        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr window, uint flags);

        /// <summary>Gets a related window, used here to read a dialog's owner.</summary>
        /// <param name="window">Starting window.</param>
        /// <param name="command">Win32 relationship selector; callers pass GW_OWNER (4).</param>
        /// <returns>Related HWND, or zero when there is no related window.</returns>
        [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr window, uint command);

        /// <summary>Gets the parent window of a control to verify its dialog-page ancestry.</summary>
        /// <param name="window">Child window whose parent is requested.</param>
        /// <returns>Parent HWND, or zero when the window has no parent.</returns>
        [DllImport("user32.dll")] private static extern IntPtr GetParent(IntPtr window);

        /// <summary>Finds a child control by its dialog command ID.</summary>
        /// <param name="window">Dialog that owns the control.</param>
        /// <param name="id">Button command ID, 1 for OK or 2 for Cancel.</param>
        /// <returns>Matching child HWND, or zero if that ID is absent.</returns>
        [DllImport("user32.dll")] private static extern IntPtr GetDlgItem(IntPtr window, int id);

        /// <summary>Gets a child control's dialog ID for matching the known VBIDE General fields.</summary>
        /// <param name="window">Control whose ID is queried.</param>
        /// <returns>Dialog control ID, or zero when no ID is assigned.</returns>
        [DllImport("user32.dll")] private static extern int GetDlgCtrlID(IntPtr window);

        /// <summary>Checks whether an HWND still identifies a live window.</summary>
        /// <param name="window">Window handle to check.</param>
        /// <returns><see langword="true"/> while the handle identifies a window.</returns>
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr window);

        /// <summary>Checks whether a window is visible, including visibility inherited from its parent.</summary>
        /// <param name="window">Window handle to check.</param>
        /// <returns><see langword="true"/> when the window is visible.</returns>
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);

        /// <summary>Checks whether a window is enabled to receive user input.</summary>
        /// <param name="window">Window handle to check.</param>
        /// <returns><see langword="true"/> when the window is enabled.</returns>
        [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr window);

        /// <summary>Checks whether a window uses the Unicode character set for its messages.</summary>
        /// <param name="window">Window handle to check.</param>
        /// <returns><see langword="true"/> when Unicode messages are used.</returns>
        [DllImport("user32.dll")] private static extern bool IsWindowUnicode(IntPtr window);

        /// <summary>Reads a pointer-sized window value, used here to inspect edit-control style flags.</summary>
        /// <param name="window">Control whose window data is queried.</param>
        /// <param name="index">Window-data index; callers pass GWL_STYLE (-16).</param>
        /// <returns>Pointer-sized value stored at the requested index.</returns>
        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")] private static extern IntPtr GetWindowLongPtr(IntPtr window, int index);

        /// <summary>Copies a window's Unicode class name into the caller-provided buffer.</summary>
        /// <param name="window">Window whose class is queried.</param>
        /// <param name="text">Buffer that receives the null-terminated class name.</param>
        /// <param name="capacity">Buffer capacity in characters.</param>
        /// <returns>Number of characters copied, excluding the terminator; zero indicates failure.</returns>
        [DllImport("user32.dll", EntryPoint = "GetClassNameW", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder text, int capacity);

        /// <summary>Enumerates top-level windows, stopping when the visitor returns false.</summary>
        /// <param name="visitor">Callback invoked for each top-level window.</param>
        /// <param name="state">Opaque value passed unchanged to the callback.</param>
        /// <returns><see langword="true"/> when enumeration completes; <see langword="false"/> when it fails or is stopped.</returns>
        [DllImport("user32.dll")] private static extern bool EnumWindows(Visitor visitor, IntPtr state);

        /// <summary>Enumerates child windows of the supplied dialog, stopping when the visitor returns false.</summary>
        /// <param name="window">Dialog whose children are enumerated.</param>
        /// <param name="visitor">Callback invoked for each child window.</param>
        /// <param name="state">Opaque value passed unchanged to the callback.</param>
        /// <returns><see langword="true"/> when enumeration completes; <see langword="false"/> when it fails or is stopped.</returns>
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr window, Visitor visitor, IntPtr state);

        /// <summary>Queues the verified WM_COMMAND close message without waiting for the dialog procedure.</summary>
        /// <param name="window">Captured Project Properties dialog.</param>
        /// <param name="message">Window message; callers pass WM_COMMAND (0x111).</param>
        /// <param name="parameter">Command ID, 1 for OK or 2 for Cancel.</param>
        /// <param name="data">Verified button HWND placed in the message's lParam.</param>
        /// <returns><see langword="true"/> when the message was queued; this does not prove the dialog closed.</returns>
        [DllImport("user32.dll", EntryPoint = "PostMessageW", SetLastError = true)] private static extern bool PostMessage(IntPtr window, uint message, UIntPtr parameter, IntPtr data);

        /// <summary>Sends a Unicode text getter with abort-if-hung and error-on-exit flags.</summary>
        /// <param name="window">Owned window whose text is requested.</param>
        /// <param name="message">Getter message; callers pass WM_GETTEXT (0x0D).</param>
        /// <param name="parameter">Text buffer capacity in characters.</param>
        /// <param name="text">Receives the returned text.</param>
        /// <param name="flags">SendMessageTimeout behavior flags; callers pass 0x23.</param>
        /// <param name="timeout">Wait limit in milliseconds; callers pass 150.</param>
        /// <param name="result">Receives the message result, including copied character count.</param>
        /// <returns>Nonzero when the message call returns; zero indicates timeout or failure.</returns>
        [DllImport("user32.dll", EntryPoint = "SendMessageTimeoutW", CharSet = CharSet.Unicode, SetLastError = true)] private static extern IntPtr SendMessageTimeoutText(IntPtr window, uint message, UIntPtr parameter, StringBuilder text, uint flags, uint timeout, out UIntPtr result);

        /// <summary>Sends one Unicode edit-control setter message with abort-if-hung and error-on-exit flags.</summary>
        /// <param name="window">Captured HelpFile or HelpContextID edit.</param>
        /// <param name="message">Setter message; callers pass WM_SETTEXT (0x0C).</param>
        /// <param name="parameter">Unused WM_SETTEXT wParam, passed as zero.</param>
        /// <param name="text">Exact text to set.</param>
        /// <param name="flags">SendMessageTimeout behavior flags; callers pass 0x23.</param>
        /// <param name="timeout">Wait limit in milliseconds; callers pass 150.</param>
        /// <param name="result">Receives the setter result; zero means the control refused the text.</param>
        /// <returns>Nonzero when the message call returns; zero indicates timeout or failure.</returns>
        [DllImport("user32.dll", EntryPoint = "SendMessageTimeoutW", CharSet = CharSet.Unicode, SetLastError = true)] private static extern IntPtr SendMessageTimeoutWrite(IntPtr window, uint message, UIntPtr parameter, string text, uint flags, uint timeout, out UIntPtr result);

        /// <summary>Sends a scalar control message, used to read the currently selected tab index.</summary>
        /// <param name="window">Captured tab control.</param>
        /// <param name="message">Control message; callers pass TCM_GETCURSEL (0x130B).</param>
        /// <param name="parameter">Message wParam; callers pass zero.</param>
        /// <param name="data">Message lParam; callers pass zero.</param>
        /// <param name="flags">SendMessageTimeout behavior flags; callers pass 0x23.</param>
        /// <param name="timeout">Wait limit in milliseconds; callers pass 150.</param>
        /// <param name="result">Receives the selected tab index.</param>
        /// <returns>Nonzero when the message call returns; zero indicates timeout or failure.</returns>
        [DllImport("user32.dll", EntryPoint = "SendMessageTimeoutW", SetLastError = true)] private static extern IntPtr SendMessageTimeoutScalar(IntPtr window, uint message, UIntPtr parameter, IntPtr data, uint flags, uint timeout, out UIntPtr result);
    }
}
