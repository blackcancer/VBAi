using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace VBAi
{

    /// <summary>Implements the Access save-confirmation gate and its exact native-dialog reader.</summary>
    internal sealed partial class VbeProjectComponents
    {

        /// <summary>Separates the one native confirmation from saved-state verification.</summary>
        internal interface IAccessSaveConfirmation
        {

            /// <summary>Rejects preexisting dialogs and arms this save-scoped confirmation gate once.</summary>
            void Prepare();

            /// <summary>Consumes the immediate pre-Save check and verifies no Access dialog predates the save.</summary>
            void RequireBeforeSave();

            /// <summary>Reads and freezes the one qualified dialog produced after the pre-Save check.</summary>
            /// <returns>The frozen candidate, or <see langword="null"/> when no owned dialog is present.</returns>
            AccessSaveConfirmationCandidate Observe();

            /// <summary>Revalidates the approved save context and queues at most one response to the frozen dialog.</summary>
            /// <param name="candidate">Candidate returned by this gate's observation; candidates from another gate are refused.</param>
            /// <param name="revalidateApprovedContext">Callback that rechecks the caller's current approval and project identity.</param>
            /// <param name="requireDeliveryDeadline">Optional final deadline check immediately before native delivery.</param>
            void Confirm(AccessSaveConfirmationCandidate candidate, Action revalidateApprovedContext, Action requireDeliveryDeadline = null);

            /// <summary>Gets the number of native response attempts claimed by this gate.</summary>
            /// <value>Zero before delivery is claimed; one once the single enqueue is attempted, including uncertain failure.</value>
            int ConfirmationAttempts { get; }

            /// <summary>Gets whether the native queue accepted the one response request.</summary>
            /// <value><see langword="true"/> only after enqueue reports acceptance.</value>
            bool ConfirmationQueued { get; }

            /// <summary>Gets whether a qualified dialog is still being tracked for this save.</summary>
            /// <value>Set after observation and retained through uncertain or queued delivery.</value>
            bool ConfirmationPending { get; }
        }

        /// <summary>Only explicitly approved standard modules and classes can be confirmed.</summary>
        internal sealed class AccessSaveApprovedComponent
        {

            /// <summary>Exact Access object-list label approved for inclusion in the save prompt.</summary>
            internal readonly string Name;

            /// <summary>Access VBComponent type; only standard modules (1) and class modules (2) are admissible.</summary>
            internal readonly int Type;

            /// <summary>Captures the approved component's exact Access list label and VBComponent type.</summary>
            /// <param name="name">Component name used to form the exact <c>Module: name</c> dialog entry.</param>
            /// <param name="type">VBComponent type, validated by the gate as 1 or 2.</param>
            internal AccessSaveApprovedComponent(string name, int type) { Name = name; Type = type; }
        }

        /// <summary>Immutable native target and identity frozen by this specific save gate.</summary>
        internal sealed class AccessSaveConfirmationCandidate
        {

            /// <summary>Native handles of the frozen dialog and its affirmative button.</summary>
            internal readonly IntPtr Window, YesButton;

            /// <summary>Process and native owner-thread IDs captured with the dialog identity.</summary>
            internal readonly uint ProcessId, ThreadId;

            /// <summary>Deterministic digest of the validated dialog and child-control snapshot.</summary>
            internal readonly string Fingerprint;

            /// <summary>Gate instance that created this candidate; confirmation requires reference identity.</summary>
            internal readonly object Owner;

            /// <summary>Freezes dialog identity, affirmative-button handle, owner gate, and snapshot fingerprint.</summary>
            /// <param name="owner">Gate instance that owns and later validates this candidate.</param>
            /// <param name="dialog">Verified native dialog snapshot with exactly one affirmative button.</param>
            /// <param name="fingerprint">Digest used to detect replacement or mutation before response delivery.</param>
            internal AccessSaveConfirmationCandidate(object owner, AccessSaveDialogSnapshot dialog, string fingerprint)
            {
                Owner = owner; Window = dialog.Window; ProcessId = dialog.ProcessId; ThreadId = dialog.ThreadId;
                YesButton = dialog.Controls.Single(control => control.Id == 1).Window; Fingerprint = fingerprint;
            }
        }

        /// <summary>Read-only injectable native inventory; incomplete enumeration is never an absence proof.</summary>
        internal sealed class AccessSaveDialogInventory
        {

            /// <summary>Whether enumeration completed without errors; false never proves that no dialog exists.</summary>
            internal bool Complete;

            /// <summary>Process and owner native-thread IDs for the enumeration scope.</summary>
            internal uint ProcessId, OwnerThreadId;

            /// <summary>All dialogs observed in the scoped Access process, including dialogs that fail qualification.</summary>
            internal AccessSaveDialogSnapshot[] Dialogs;
        }

            /// <summary>Read-only native identity and visible state captured for one Access dialog.</summary>
        internal sealed class AccessSaveDialogSnapshot
        {

            /// <summary>Native HWND of the captured top-level dialog.</summary>
            internal IntPtr Window;

            /// <summary>Owning process, native UI thread, and window style captured from the dialog.</summary>
            internal uint ProcessId, ThreadId, Style;

            /// <summary>Win32 class name and exact caption observed on the dialog.</summary>
            internal string Class, Caption;

            /// <summary>Visibility, enabled state, and completeness of the child-control enumeration.</summary>
            internal bool Visible, Enabled, ChildrenComplete;

            /// <summary>Captured child controls used to compare the prompt with the qualified Access dialog.</summary>
            internal AccessSaveDialogControl[] Controls;
        }

            /// <summary>Native identity, style, text, and item state captured for one dialog child control.</summary>
        internal sealed class AccessSaveDialogControl
        {

            /// <summary>Native HWND of this child control.</summary>
            internal IntPtr Window;

            /// <summary>Owning process, native UI thread, and Win32 style of this child control.</summary>
            internal uint ProcessId, ThreadId, Style;

            /// <summary>Dialog control identifier, such as 1 for the affirmative button or 5142 for the object list.</summary>
            internal int Id;

            /// <summary>Win32 class name and captured control text.</summary>
            internal string Class, Text;

            /// <summary>Visibility, enabled state, and completeness of any list-item enumeration.</summary>
            internal bool Visible, Enabled, ItemsComplete;

            /// <summary>Captured list entries when this control exposes selectable items.</summary>
            internal AccessSaveDialogItem[] Items;
        }

            /// <summary>Text and selection state of one item in the Access save prompt's object list.</summary>
        internal sealed class AccessSaveDialogItem
        {

            /// <summary>Exact object label displayed for this list entry.</summary>
            internal string Text;

            /// <summary>Whether Access marked this object for the pending save confirmation.</summary>
            internal bool Selected;
        }

        /// <summary>Creates the production gate bound to the supplied VBE window, process, owner thread, and approved modules.</summary>
        /// <param name="vbeWindow">VBE top-level HWND used to establish the native owner thread.</param>
        /// <param name="processId">Process that must own every observed dialog and child.</param>
        /// <param name="approvedComponents">Exact standard-module and class-module names permitted in the save list.</param>
        /// <returns>A gate whose reads and one response enqueue execute through the native owner-thread reader.</returns>
        internal static IAccessSaveConfirmation CreateNativeAccessSaveConfirmation(
            IntPtr vbeWindow, int processId, IEnumerable<AccessSaveApprovedComponent> approvedComponents)
        {
            var reader = new NativeAccessSaveConfirmationReader(vbeWindow, processId);
            return new AccessSaveConfirmation(reader.Read, reader.Enqueue, reader.RequireOwner,
                (uint)processId, reader.OwnerThreadId, approvedComponents);
        }

        /// <summary>Refuses unknown dialogs and consumes every confirmation call before any enqueue.</summary>
        internal sealed class AccessSaveConfirmation : IAccessSaveConfirmation
        {

            /// <summary>Reads a bounded inventory from the exact Access owner thread.</summary>
            private readonly Func<AccessSaveDialogInventory> read;

            /// <summary>Queues the affirmative response for a candidate after all identity checks pass.</summary>
            private readonly Func<AccessSaveConfirmationCandidate, bool> enqueue;

            /// <summary>Throws unless the caller is executing on the captured native owner thread.</summary>
            private readonly Action requireOwner;

            /// <summary>Immutable process and UI-thread identity required for every inventory read.</summary>
            private readonly uint processId, ownerThreadId;

            /// <summary>Ordinal set of exact <c>Module: name</c> labels allowed in the Access prompt.</summary>
            private readonly HashSet<string> approved;

            /// <summary>One-use gate flags; any identity, inventory, or delivery failure faults the gate permanently.</summary>
            private bool preparationClaimed, prepared, beforeSaveClaimed, confirmationClaimed, faulted;

            /// <summary>First qualified candidate, retained so later reads must describe the same native dialog.</summary>
            private AccessSaveConfirmationCandidate observed;

            /// <summary>Gets or sets the confirmation attempts.</summary>
            /// <value>Current confirmation attempts exposed by access save confirmation.</value>
            public int ConfirmationAttempts { get; private set; }

            /// <summary>Gets or sets the confirmation queued.</summary>
            /// <value>Current confirmation queued exposed by access save confirmation.</value>
            public bool ConfirmationQueued { get; private set; }

            /// <summary>Gets or sets the confirmation pending.</summary>
            /// <value>Current confirmation pending exposed by access save confirmation.</value>
            public bool ConfirmationPending { get; private set; }

            /// <summary>Creates a save gate with injected native I/O and a fixed approved-component allowlist.</summary>
            /// <param name="read">Bounded dialog inventory reader for the captured owner thread.</param>
            /// <param name="enqueue">Single-response queue operation; a false result is treated as uncertain and never retried.</param>
            /// <param name="requireOwner">Owner-thread assertion run before native reads and delivery.</param>
            /// <param name="processId">Nonzero process ID that must own every observed native window.</param>
            /// <param name="ownerThreadId">Nonzero native UI-thread ID that must own every observed window.</param>
            /// <param name="approvedComponents">Distinct Type 1/2 component names admitted to the exact prompt list.</param>
            internal AccessSaveConfirmation(Func<AccessSaveDialogInventory> read,
                Func<AccessSaveConfirmationCandidate, bool> enqueue, Action requireOwner,
                uint processId, uint ownerThreadId, IEnumerable<AccessSaveApprovedComponent> approvedComponents)
            {
                this.read = read ?? throw new ArgumentNullException(nameof(read));
                this.enqueue = enqueue ?? throw new ArgumentNullException(nameof(enqueue));
                this.requireOwner = requireOwner ?? throw new ArgumentNullException(nameof(requireOwner));
                if (processId == 0 || ownerThreadId == 0) throw new ArgumentException("Access confirmation requires exact process and native owner thread.");
                this.processId = processId; this.ownerThreadId = ownerThreadId;
                if (approvedComponents == null) throw new ArgumentNullException(nameof(approvedComponents));
                approved = new HashSet<string>(StringComparer.Ordinal);
                var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var component in approvedComponents)
                {
                    if (component == null || (component.Type != 1 && component.Type != 2) ||
                        string.IsNullOrWhiteSpace(component.Name) || component.Name.Length > 255 ||
                        component.Name.IndexOf('\0') >= 0 || component.Name.IndexOf('\r') >= 0 ||
                        component.Name.IndexOf('\n') >= 0 || !names.Add(component.Name) || !approved.Add("Module: " + component.Name))
                        throw new InvalidOperationException("Access confirmation requires distinct approved Type1/Type2 component names.");
                }
            }

            /// <summary>Consumes preparation and refuses to arm while any dialog already exists.</summary>
            public void Prepare()
            {
                if (preparationClaimed) throw new InvalidOperationException("Access confirmation preparation was already claimed.");
                preparationClaimed = true;
                RequireNoDialog();
                prepared = true;
            }

            /// <summary>Consumes the immediate pre-Save check and rejects a dialog that predates that Save.</summary>
            public void RequireBeforeSave()
            {
                RequirePrepared();
                if (beforeSaveClaimed) throw new InvalidOperationException("The Access pre-Save confirmation gate was already consumed.");
                beforeSaveClaimed = true;
                RequireNoDialog(); // No preexisting dialog may be accepted as the result of this Save.
            }

            /// <summary>Observes  for access save confirmation.</summary>
            /// <returns>access save confirmation candidate produced by the operation for observe on access save confirmation.</returns>
            public AccessSaveConfirmationCandidate Observe()
            {
                RequirePrepared();
                if (!beforeSaveClaimed) throw new InvalidOperationException("Access confirmation requires the immediate pre-Save gate.");
                try
                {
                    var inventory = ReadVerifiedInventory();
                    if (inventory.Dialogs.Length == 0)
                    {
                        ConfirmationPending = false;
                        return null;
                    }
                    var dialog = RequireKnownDialog(inventory);
                    string fingerprint = Fingerprint(dialog);
                    if (observed != null && fingerprint != observed.Fingerprint)
                        throw new InvalidOperationException("The observed Access save dialog changed; no further confirmation is permitted.");
                    if (observed == null) observed = new AccessSaveConfirmationCandidate(this, dialog, fingerprint);
                    ConfirmationPending = true;
                    return observed; // The original modal may remain visible after its one queued response.
                }
                catch { faulted = true; throw; }
            }

            /// <summary>Claims the only response attempt, rechecks authority and dialog identity, then queues once.</summary>
            /// <param name="candidate">Exact candidate most recently observed by this instance.</param>
            /// <param name="revalidateApprovedContext">Revalidation callback for project identity and the approved save scope.</param>
            /// <param name="requireDeliveryDeadline">Optional final deadline assertion before the native enqueue call.</param>
            public void Confirm(AccessSaveConfirmationCandidate candidate, Action revalidateApprovedContext, Action requireDeliveryDeadline = null)
            {
                RequirePrepared();
                if (confirmationClaimed) throw new InvalidOperationException("The Access confirmation call was already consumed; do not retry.");
                confirmationClaimed = true; // Includes a context/identity refusal before native delivery.
                ConfirmationPending = true;
                try
                {
                    if (candidate == null || !ReferenceEquals(candidate.Owner, this) || !ReferenceEquals(candidate, observed))
                        throw new InvalidOperationException("The Access confirmation candidate belongs to another save gate.");
                    if (revalidateApprovedContext == null) throw new ArgumentNullException(nameof(revalidateApprovedContext));
                    RequireSameCandidate(candidate);
                    revalidateApprovedContext();
                    RequireSameCandidate(candidate); // Approval validation itself must not substitute the dialog.
                    requireOwner();
                    requireDeliveryDeadline?.Invoke(); // The final native read must not outlive this Save's delivery deadline.
                    ConfirmationAttempts = 1; // Claim before the one enqueue, including false/throw uncertainty.
                    if (!enqueue(candidate))
                        throw new InvalidOperationException("The Access confirmation enqueue was not accepted; its outcome is uncertain and must not be retried.");
                    ConfirmationQueued = true;
                }
                catch { faulted = true; throw; }
            }

            /// <summary>Rejects use before preparation or after any earlier check has faulted the gate.</summary>
            private void RequirePrepared()
            {
                if (!prepared || faulted) throw new InvalidOperationException("The Access confirmation gate is unprepared or refused.");
            }

            /// <summary>Reads inventory on the owner thread and verifies it is complete and bound to the captured process/thread.</summary>
            /// <returns>Verified dialog inventory; incomplete enumeration or any foreign/hidden dialog throws.</returns>
            private AccessSaveDialogInventory ReadVerifiedInventory()
            {
                requireOwner();
                var inventory = read();
                if (inventory == null || !inventory.Complete || inventory.ProcessId != processId ||
                    inventory.OwnerThreadId != ownerThreadId || inventory.Dialogs == null)
                    throw new InvalidOperationException("The exact Access native dialog inventory is incomplete or changed.");
                foreach (var dialog in inventory.Dialogs)
                    if (dialog == null || dialog.Window == IntPtr.Zero || dialog.ProcessId != processId ||
                        dialog.ThreadId != ownerThreadId || dialog.Class != "#32770" || !dialog.Visible)
                        throw new InvalidOperationException("An Access dialog belongs to another native UI thread or has unverified identity.");
                return inventory;
            }

            /// <summary>Requires a complete owner-scoped inventory containing no preexisting dialogs.</summary>
            private void RequireNoDialog()
            {
                try
                {
                    if (ReadVerifiedInventory().Dialogs.Length != 0)
                        throw new InvalidOperationException("A preexisting owned Access dialog prevents this Save.");
                }
                catch { faulted = true; throw; }
            }

            /// <summary>Re-reads the only dialog and compares its validated fingerprint with the frozen candidate.</summary>
            /// <param name="candidate">Candidate whose native dialog identity must remain unchanged.</param>
            private void RequireSameCandidate(AccessSaveConfirmationCandidate candidate)
            {
                var inventory = ReadVerifiedInventory();
                if (inventory.Dialogs.Length != 1 ||
                    Fingerprint(RequireKnownDialog(inventory)) != candidate.Fingerprint)
                    throw new InvalidOperationException("The frozen Access save confirmation candidate changed before delivery.");
            }

            /// <summary>Validates the exact Access save prompt, required controls, and fully selected approved object list.</summary>
            /// <param name="inventory">Complete inventory containing exactly one dialog from the approved Access process/thread.</param>
            /// <returns>The sole snapshot after caption, controls, list style, item labels, and selections pass validation.</returns>
            private AccessSaveDialogSnapshot RequireKnownDialog(AccessSaveDialogInventory inventory)
            {
                if (inventory.Dialogs.Length != 1) throw new InvalidOperationException("Multiple owned Access dialogs prevent confirmation.");
                var dialog = inventory.Dialogs[0];
                if (dialog.Caption != "Enregistrer" || !dialog.Enabled || !dialog.ChildrenComplete || dialog.Controls == null)
                    throw new InvalidOperationException("The Access save dialog does not match the qualified native prompt.");
                var ids = new HashSet<int>(); var handles = new HashSet<IntPtr>();
                foreach (var control in dialog.Controls)
                {
                    if (control == null || control.Window == IntPtr.Zero || !handles.Add(control.Window) ||
                        control.ProcessId != processId || control.ThreadId != ownerThreadId || !ids.Add(control.Id))
                        throw new InvalidOperationException("Access save dialog child identity or IDs are ambiguous.");
                    bool required = control.Id == 1 || control.Id == 7 || control.Id == 2 || control.Id == 5271 || control.Id == 5142;
                    if (!required && (control.Visible || control.Class != "Static" || control.Text != "DAL=on"))
                        throw new InvalidOperationException("An unknown Access save dialog child prevents confirmation.");
                }
                RequireControl(dialog, 1, "Button", "&Oui");
                RequireControl(dialog, 7, "Button", "&Non pour tout");
                RequireControl(dialog, 2, "Button", "Annuler");
                RequireControl(dialog, 5271, "Static", "Enregistrer les modifications apportées aux objets suivants\u00A0?");
                var list = RequireControl(dialog, 5142, "ListBox", null);
                bool ownerDraw = (list.Style & 0x30) != 0, hasStrings = (list.Style & 0x40) != 0;
                bool multiSelect = (list.Style & (0x8 | 0x800)) != 0;
                if ((ownerDraw && !hasStrings) || (list.Style & 0x2000) != 0 || !multiSelect ||
                    !list.ItemsComplete || list.Items == null || list.Items.Length < 1 || list.Items.Length > 1000)
                    throw new InvalidOperationException("Access save dialog list strings or complete multiselection are unverified.");
                var entries = new HashSet<string>(StringComparer.Ordinal);
                foreach (var item in list.Items)
                    if (item == null || item.Text == null || item.Text.Length > 512 || !item.Selected ||
                        !approved.Contains(item.Text) || !entries.Add(item.Text))
                        throw new InvalidOperationException("Access save confirmation includes an unapproved, duplicate or unselected object.");
                return dialog;
            }

            /// <summary>Requires one visible, enabled child with the requested control ID and exact expected class/text.</summary>
            /// <param name="dialog">Validated prompt whose child-control snapshot is searched.</param>
            /// <param name="id">Win32 dialog control ID expected for the child.</param>
            /// <param name="type">Exact Win32 class name required, such as <c>Button</c> or <c>Static</c>.</param>
            /// <param name="text">Exact expected caption, or <see langword="null"/> when the caption is not constrained.</param>
            /// <returns>The unique matching control; zero, duplicate, hidden, disabled, or changed controls throw.</returns>
            private static AccessSaveDialogControl RequireControl(AccessSaveDialogSnapshot dialog, int id, string type, string text)
            {
                var matches = dialog.Controls.Where(control => control.Id == id).ToArray();
                if (matches.Length != 1 || matches[0].Class != type || !matches[0].Visible || !matches[0].Enabled ||
                    (text != null && matches[0].Text != text))
                    throw new InvalidOperationException("The qualified Access save control " + id + " changed.");
                return matches[0];
            }

            /// <summary>Builds a stable fingerprint from dialog identity, controls, text, visibility, and list items.</summary>
            /// <param name="dialog">Captured dialog snapshot to encode, with child controls ordered by HWND.</param>
            /// <returns>Length-delimited invariant representation used to detect any snapshot change before delivery.</returns>
            private static string Fingerprint(AccessSaveDialogSnapshot dialog)
            {
                var value = new StringBuilder();
                AddFingerprint(value, dialog.Window.ToInt64(), dialog.ProcessId, dialog.ThreadId, dialog.Style,
                    dialog.Class, dialog.Caption, dialog.Visible, dialog.Enabled, dialog.ChildrenComplete);
                foreach (var control in dialog.Controls.OrderBy(control => control.Window.ToInt64()))
                {
                    AddFingerprint(value, control.Window.ToInt64(), control.ProcessId, control.ThreadId, control.Style,
                        control.Id, control.Class, control.Text, control.Visible, control.Enabled, control.ItemsComplete);
                    if (control.Items != null)
                        foreach (var item in control.Items) AddFingerprint(value, item.Text, item.Selected);
                }
                return value.ToString();
            }

            /// <summary>Appends values with invariant formatting and length prefixes to avoid delimiter ambiguity.</summary>
            /// <param name="result">Fingerprint buffer receiving each encoded value.</param>
            /// <param name="values">Snapshot values to append; null is represented explicitly.</param>
            private static void AddFingerprint(StringBuilder result, params object[] values)
            {
                foreach (object value in values)
                {
                    string text = value == null ? "<null>" : Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture);
                    result.Append(text.Length).Append(':').Append(text).Append(';');
                }
            }
        }

        /// <summary>Uses only exact current-process owner-thread native getters and one WM_COMMAND enqueue.</summary>
        private sealed class NativeAccessSaveConfirmationReader
        {

            /// <summary>VBE HWND used to bind this reader to its owning process and native UI thread.</summary>
            private readonly IntPtr vbeWindow;

            /// <summary>Process ID every enumerated dialog and control must belong to.</summary>
            private readonly uint processId;

            /// <summary>Native UI thread captured at construction and required for all subsequent reads and enqueue.</summary>
            internal readonly uint OwnerThreadId;

            /// <summary>Enumeration safety limits: top-level windows, child controls, and completed snapshot time in milliseconds.</summary>
            private const int WindowBound = 8192, ChildBound = 128, SnapshotReadBoundMilliseconds = 5000;

            /// <summary>Elapsed-time guard for one inventory snapshot; it cannot cancel an in-progress synchronous Win32 call.</summary>
            private Stopwatch snapshotReadWatch;

            /// <summary>Callback shape used by Win32 top-level and child window enumerators.</summary>
            /// <param name="window">Current HWND provided by the enumerator.</param>
            /// <param name="state">Caller context pointer forwarded unchanged by Win32.</param>
            /// <returns><see langword="true"/> to continue enumeration; <see langword="false"/> to stop.</returns>
            private delegate bool Visitor(IntPtr window, IntPtr state);

            /// <summary>Reads the calling native thread ID from Kernel32.</summary>
            /// <returns>Current Win32 thread ID.</returns>
            [DllImport("kernel32.dll", EntryPoint = "GetCurrentThreadId")] private static extern uint NativeGetCurrentThreadId();

            /// <summary>Reads the calling process ID from Kernel32.</summary>
            /// <returns>Current Win32 process ID.</returns>
            [DllImport("kernel32.dll", EntryPoint = "GetCurrentProcessId")] private static extern uint NativeGetCurrentProcessId();

            /// <summary>Reads the owning thread and process IDs for an HWND.</summary>
            /// <param name="window">HWND whose owner is queried.</param>
            /// <param name="process">Receives the owning process ID.</param>
            /// <returns>Owning native thread ID, or zero if the HWND is invalid.</returns>
            [DllImport("user32.dll", EntryPoint = "GetWindowThreadProcessId")] private static extern uint NativeGetWindowThreadProcessId(IntPtr window, out uint process);

            /// <summary>Reads the requested ancestor HWND, used here with GA_ROOT (2) to verify dialog ownership.</summary>
            /// <param name="window">Starting HWND.</param>
            /// <param name="flags">GetAncestor relationship selector, normally GA_ROOT (2).</param>
            /// <returns>Ancestor HWND, or zero when none exists.</returns>
            [DllImport("user32.dll", EntryPoint = "GetAncestor")] private static extern IntPtr NativeGetAncestor(IntPtr window, uint flags);

            /// <summary>Enumerates top-level windows; callback cancellation or API failure makes the inventory incomplete.</summary>
            /// <param name="visitor">Callback invoked for each top-level HWND.</param>
            /// <param name="state">Opaque callback context forwarded by Win32.</param>
            /// <returns>Win32 completion status; callers also track the bounded callback count and captured errors.</returns>
            [DllImport("user32.dll", SetLastError = true, EntryPoint = "EnumWindows")] private static extern bool NativeEnumWindows(Visitor visitor, IntPtr state);

            /// <summary>Enumerates descendants of a dialog; the API return value is not treated as completeness evidence.</summary>
            /// <param name="parent">Dialog HWND whose descendants are enumerated.</param>
            /// <param name="visitor">Callback invoked for each descendant HWND.</param>
            /// <param name="state">Opaque callback context forwarded by Win32.</param>
            /// <returns>Win32 API result, which this reader deliberately does not use as its completeness flag.</returns>
            [DllImport("user32.dll", EntryPoint = "EnumChildWindows")] private static extern bool NativeEnumChildWindows(IntPtr parent, Visitor visitor, IntPtr state);

            /// <summary>Copies an HWND's Unicode window-class name into the supplied buffer.</summary>
            /// <param name="window">Window whose class name is requested.</param>
            /// <param name="text">Destination buffer for the class name.</param>
            /// <param name="max">Destination capacity in characters.</param>
            /// <returns>Characters copied, excluding the terminator; zero indicates failure.</returns>
            [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "GetClassNameW")] private static extern int NativeGetClassNameW(IntPtr window, StringBuilder text, int max);

            /// <summary>Tests the WS_VISIBLE state for an HWND.</summary>
            /// <param name="window">Window whose visibility is checked.</param>
            /// <returns><see langword="true"/> when the window has the visible style.</returns>
            [DllImport("user32.dll", EntryPoint = "IsWindowVisible")] private static extern bool NativeIsWindowVisible(IntPtr window);

            /// <summary>Tests whether an HWND is enabled for user interaction.</summary>
            /// <param name="window">Window whose enabled state is checked.</param>
            /// <returns><see langword="true"/> when the window is enabled.</returns>
            [DllImport("user32.dll", EntryPoint = "IsWindowEnabled")] private static extern bool NativeIsWindowEnabled(IntPtr window);

            /// <summary>Reads a child window's dialog control identifier.</summary>
            /// <param name="window">Child HWND whose control ID is requested.</param>
            /// <returns>Control identifier, or zero when the window has no identifier.</returns>
            [DllImport("user32.dll", EntryPoint = "GetDlgCtrlID")] private static extern int NativeGetDlgCtrlID(IntPtr window);

            /// <summary>Reads a pointer-sized window attribute, used with GWL_STYLE (-16).</summary>
            /// <param name="window">Window whose attribute is read.</param>
            /// <param name="index">Negative index selecting the window attribute; this reader requests GWL_STYLE.</param>
            /// <returns>Attribute value, pointer-sized on x64.</returns>
            [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")] private static extern IntPtr NativeGetWindowLongPtrW(IntPtr window, int index);

            /// <summary>Sends a synchronous Unicode Win32 message to a window on the current owner thread.</summary>
            /// <param name="window">Target HWND.</param>
            /// <param name="message">Win32 message ID.</param>
            /// <param name="wParam">Message-specific first argument.</param>
            /// <param name="lParam">Message-specific second argument.</param>
            /// <returns>Message-specific pointer-sized result.</returns>
            [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "SendMessageW")] private static extern IntPtr NativeSendMessageW(IntPtr window, uint message, UIntPtr wParam, IntPtr lParam);

            /// <summary>Posts a Unicode Win32 message asynchronously to the window queue.</summary>
            /// <param name="window">Target HWND.</param>
            /// <param name="message">Win32 message ID.</param>
            /// <param name="wParam">Message-specific first argument.</param>
            /// <param name="lParam">Message-specific second argument.</param>
            /// <returns><see langword="true"/> when queued; false does not establish whether a side effect occurred.</returns>
            [DllImport("user32.dll", SetLastError = true, EntryPoint = "PostMessageW")] private static extern bool NativePostMessageW(IntPtr window, uint message, UIntPtr wParam, IntPtr lParam);


            // This bounds completed snapshot work; same-thread SendMessageW is not
            // represented as a cancellable call or a guaranteed wall-clock timeout.
            /// <summary>Throws when completed snapshot work has reached the five-second budget.</summary>
            private void RequireReadBudget()
            {
                if (snapshotReadWatch != null && snapshotReadWatch.ElapsedMilliseconds >= SnapshotReadBoundMilliseconds)
                    throw new InvalidOperationException("Access confirmation snapshot exceeded its read budget; no confirmation is queued.");
            }

            /// <summary>Checks the snapshot budget immediately before and after one synchronous native getter.</summary>
            /// <typeparam name="T">Value returned by the native getter.</typeparam>
            /// <param name="getter">Native read to execute once within the current inventory operation.</param>
            /// <returns>The getter's value if both budget checks pass.</returns>
            private T CheckedRead<T>(Func<T> getter)
            {
                RequireReadBudget(); T result = getter(); RequireReadBudget(); return result;
            }

            /// <summary>Reads the current thread ID through the snapshot-budget guard.</summary>
            /// <returns>Current Win32 thread ID.</returns>
            private uint GetCurrentThreadId() => CheckedRead(() => NativeGetCurrentThreadId());

            /// <summary>Reads the current process ID through the snapshot-budget guard.</summary>
            /// <returns>Current Win32 process ID.</returns>
            private uint GetCurrentProcessId() => CheckedRead(() => NativeGetCurrentProcessId());

            /// <summary>Reads an HWND's owner thread and process IDs with budget checks around the native call.</summary>
            /// <param name="window">HWND whose owner is checked.</param>
            /// <param name="process">Receives the HWND's owner process ID.</param>
            /// <returns>The HWND's owner thread ID.</returns>
            private uint GetWindowThreadProcessId(IntPtr window, out uint process)
            {
                RequireReadBudget(); uint thread = NativeGetWindowThreadProcessId(window, out process); RequireReadBudget(); return thread;
            }

            /// <summary>Reads an ancestor HWND through the snapshot-budget guard.</summary>
            /// <param name="window">Starting HWND.</param>
            /// <param name="flags">Ancestor relationship selector, such as GA_ROOT.</param>
            /// <returns>The requested ancestor HWND.</returns>
            private IntPtr GetAncestor(IntPtr window, uint flags) => CheckedRead(() => NativeGetAncestor(window, flags));

            /// <summary>Enumerates top-level windows while enforcing the active snapshot budget.</summary>
            /// <param name="visitor">Callback that inspects each HWND and may stop enumeration.</param>
            /// <param name="state">Opaque context passed to the callback.</param>
            /// <returns>Native completion result; callback errors are captured and rethrown by the reader.</returns>
            private bool EnumWindows(Visitor visitor, IntPtr state) => CheckedRead(() => NativeEnumWindows(visitor, state));

            /// <summary>Enumerates descendant windows within the active snapshot budget.</summary>
            /// <param name="parent">Parent dialog HWND.</param>
            /// <param name="visitor">Callback that inspects each descendant.</param>
            /// <param name="state">Opaque context passed to the callback.</param>
            /// <returns>Native result, ignored for completeness because Win32 does not define it as such here.</returns>
            private bool EnumChildWindows(IntPtr parent, Visitor visitor, IntPtr state) => CheckedRead(() => NativeEnumChildWindows(parent, visitor, state));

            /// <summary>Reads a window's class name with budget checks around the native call.</summary>
            /// <param name="window">HWND to inspect.</param>
            /// <param name="text">Destination buffer.</param>
            /// <param name="max">Buffer capacity in characters.</param>
            /// <returns>Number of characters copied, excluding the null terminator.</returns>
            private int GetClassNameW(IntPtr window, StringBuilder text, int max) => CheckedRead(() => NativeGetClassNameW(window, text, max));

            /// <summary>Reads an HWND's visible state through the snapshot-budget guard.</summary>
            /// <param name="window">Window to inspect.</param>
            /// <returns>Whether the window has WS_VISIBLE set.</returns>
            private bool IsWindowVisible(IntPtr window) => CheckedRead(() => NativeIsWindowVisible(window));

            /// <summary>Reads an HWND's enabled state through the snapshot-budget guard.</summary>
            /// <param name="window">Window to inspect.</param>
            /// <returns>Whether the window is enabled for interaction.</returns>
            private bool IsWindowEnabled(IntPtr window) => CheckedRead(() => NativeIsWindowEnabled(window));

            /// <summary>Reads the dialog-control ID through the snapshot-budget guard.</summary>
            /// <param name="window">Child HWND to inspect.</param>
            /// <returns>Dialog-control ID.</returns>
            private int GetDlgCtrlID(IntPtr window) => CheckedRead(() => NativeGetDlgCtrlID(window));

            /// <summary>Reads a pointer-sized window attribute through the snapshot-budget guard.</summary>
            /// <param name="window">HWND whose attribute is read.</param>
            /// <param name="index">Attribute index; callers use GWL_STYLE (-16).</param>
            /// <returns>Pointer-sized attribute value.</returns>
            private IntPtr GetWindowLongPtrW(IntPtr window, int index) => CheckedRead(() => NativeGetWindowLongPtrW(window, index));

            /// <summary>Sends one synchronous message through the snapshot-budget guard.</summary>
            /// <param name="window">Target HWND.</param>
            /// <param name="message">Win32 message ID.</param>
            /// <param name="wParam">Message-specific first argument.</param>
            /// <param name="lParam">Message-specific second argument.</param>
            /// <returns>Pointer-sized message result.</returns>
            private IntPtr SendMessageW(IntPtr window, uint message, UIntPtr wParam, IntPtr lParam) =>
                CheckedRead(() => NativeSendMessageW(window, message, wParam, lParam));

            /// <summary>Captures the VBE's process and current native owner-thread identity.</summary>
            /// <param name="vbeWindow">Nonzero top-level VBE HWND whose native owner thread is captured.</param>
            /// <param name="processId">Positive host process ID that must match both the current process and VBE window.</param>
            internal NativeAccessSaveConfirmationReader(IntPtr vbeWindow, int processId)
            {
                if (processId <= 0 || vbeWindow == IntPtr.Zero) throw new ArgumentException("Access confirmation requires the native VBE HWND and host PID.");
                this.vbeWindow = vbeWindow; this.processId = (uint)processId; OwnerThreadId = GetCurrentThreadId();
                RequireOwner();
            }

            /// <summary>Requires current process/thread and the VBE HWND owner to match the captured identity.</summary>
            internal void RequireOwner()
            {
                uint actual;
                if (GetCurrentProcessId() != processId || GetCurrentThreadId() != OwnerThreadId ||
                    GetWindowThreadProcessId(vbeWindow, out actual) != OwnerThreadId || actual != processId)
                    throw new InvalidOperationException("Access save confirmation left the exact native VBE process/UI thread.");
            }

            /// <summary>Verifies a window and its root ancestor belong to the captured Access process and UI thread.</summary>
            /// <param name="window">Nonzero HWND whose process, thread, and root ownership must be checked.</param>
            private void RequireWindow(IntPtr window)
            {
                RequireOwner();
                uint actual, rootPid; IntPtr root = GetAncestor(window, 2);
                if (window == IntPtr.Zero || GetWindowThreadProcessId(window, out actual) != OwnerThreadId ||
                    actual != processId || GetWindowThreadProcessId(root, out rootPid) != OwnerThreadId || rootPid != processId)
                    throw new InvalidOperationException("The Access confirmation window or root identity changed.");
            }

            /// <summary>Reads the Unicode Win32 class name and rejects missing class information.</summary>
            /// <param name="window">Owner-verified HWND to inspect.</param>
            /// <returns>Class name copied from the native window.</returns>
            private string ClassName(IntPtr window)
            {
                var text = new StringBuilder(256);
                if (GetClassNameW(window, text, text.Capacity) <= 0) throw new InvalidOperationException("Access native control class is unavailable.");
                return text.ToString();
            }

            /// <summary>Reads the GWL_STYLE bits after verifying the HWND's process, thread, and root.</summary>
            /// <param name="window">Owner-verified HWND whose style is captured.</param>
            /// <returns>Window-style bit field as an unsigned 32-bit value.</returns>
            private uint Style(IntPtr window) { RequireWindow(window); return unchecked((uint)GetWindowLongPtrW(window, -16).ToInt64()); }

            /// <summary>Reads message for native access save confirmation reader.</summary>
            /// <param name="window">Owner-verified HWND receiving the message.</param>
            /// <param name="message">Win32 message ID used to read text or list-box state.</param>
            /// <param name="wParam">Message-specific value, such as a list item index.</param>
            /// <param name="buffer">Optional native output buffer passed as LPARAM.</param>
            /// <returns>Pointer-sized native message result converted to a 64-bit integer.</returns>
            private long ReadMessage(IntPtr window, uint message, ulong wParam, IntPtr buffer)
            {
                RequireWindow(window);
                return SendMessageW(window, message, new UIntPtr(wParam), buffer).ToInt64();
            }

            /// <summary>Reads at most 1023 UTF-16 characters from a verified control using WM_GETTEXT.</summary>
            /// <param name="window">Owner-verified control HWND.</param>
            /// <returns>Text produced by the operation for text on native access save confirmation reader.</returns>
            private string Text(IntPtr window)
            {
                IntPtr memory = Marshal.AllocHGlobal(1024 * 2);
                try
                {
                    Marshal.WriteInt16(memory, 0);
                    long copied = ReadMessage(window, 0xD, 1024, memory);
                    if (copied < 0 || copied >= 1024) throw new InvalidOperationException("Access confirmation text is truncated or unavailable.");
                    return Marshal.PtrToStringUni(memory, (int)copied);
                }
                finally { Marshal.FreeHGlobal(memory); }
            }

            /// <summary>Reads  for native access save confirmation reader.</summary>
            /// <returns>access save dialog inventory produced by the operation for read on native access save confirmation reader.</returns>
            internal AccessSaveDialogInventory Read()
            {
                if (snapshotReadWatch != null) throw new InvalidOperationException("Access confirmation snapshot is already in progress.");
                snapshotReadWatch = Stopwatch.StartNew();
                try
                {
                RequireOwner();
                var dialogs = new List<AccessSaveDialogSnapshot>(); int visited = 0;
                bool bounded = true; Exception error = null;
                Visitor visitor = (window, state) => {
                    try
                    {
                        if (++visited > WindowBound) { bounded = false; return false; }
                        uint pid; uint thread = GetWindowThreadProcessId(window, out pid);
                        if (pid != processId || !IsWindowVisible(window) || ClassName(window) != "#32770") return true;
                        dialogs.Add(new AccessSaveDialogSnapshot { Window = window, ProcessId = pid, ThreadId = thread,
                            Class = "#32770", Visible = true, Enabled = IsWindowEnabled(window) });
                        return true;
                    }
                    catch (Exception failure) { error = failure; return false; }
                };
                bool complete = EnumWindows(visitor, IntPtr.Zero);
                GC.KeepAlive(visitor);
                if (error != null) throw error;
                var inventory = new AccessSaveDialogInventory { Complete = complete && bounded,
                    ProcessId = processId, OwnerThreadId = OwnerThreadId, Dialogs = dialogs.ToArray() };
                if (!inventory.Complete || dialogs.Count != 1 || dialogs[0].ThreadId != OwnerThreadId) return inventory;
                var dialog = dialogs[0]; RequireWindow(dialog.Window);
                dialog.Caption = Text(dialog.Window); dialog.Style = Style(dialog.Window);
                if (dialog.Caption != "Enregistrer") return inventory;
                dialog.Controls = ReadChildren(dialog.Window, out bool childrenComplete);
                dialog.ChildrenComplete = childrenComplete;
                var lists = dialog.Controls.Where(control => control.Id == 5142 && control.Class == "ListBox").ToArray();
                // Do not interpret owner-draw item data as strings, or read an unknown prompt's list.
                bool knownPrompt = dialog.Controls.Count(control => control.Id == 5271 && control.Class == "Static" &&
                    control.Text == "Enregistrer les modifications apportées aux objets suivants\u00A0?") == 1;
                if (childrenComplete && knownPrompt && lists.Length == 1) ReadItems(lists[0]);
                return inventory;

                }
                finally
                {
                    try { RequireReadBudget(); }
                    finally { snapshotReadWatch = null; }
                }
            }

            /// <summary>Enumerates at most the configured child bound and captures each child's verified native state.</summary>
            /// <param name="parent">Qualified top-level dialog HWND whose direct descendants are read.</param>
            /// <param name="complete">Receives false if the child limit is exceeded; Win32's enum return is not used as completeness evidence.</param>
            /// <returns>Captured controls; an identity/read error throws and incomplete bounded enumeration is reported separately.</returns>
            private AccessSaveDialogControl[] ReadChildren(IntPtr parent, out bool complete)
            {
                var controls = new List<AccessSaveDialogControl>(); int visited = 0;
                bool bounded = true; Exception error = null;
                Visitor visitor = (window, state) => {
                    try
                    {
                        if (++visited > ChildBound) { bounded = false; return false; }
                        RequireWindow(window);
                        if (GetAncestor(window, 2) != parent) throw new InvalidOperationException("Access dialog child ancestry changed.");
                        var control = new AccessSaveDialogControl { Window = window, ProcessId = processId, ThreadId = OwnerThreadId,
                            Id = GetDlgCtrlID(window), Class = ClassName(window), Style = Style(window),
                            Visible = IsWindowVisible(window), Enabled = IsWindowEnabled(window) };
                        if (control.Class == "Button" || control.Class == "Static") control.Text = Text(window);
                        controls.Add(control); return true;
                    }
                    catch (Exception failure) { error = failure; return false; }
                };
                // Win32 documents no meaningful EnumChildWindows return value.
                EnumChildWindows(parent, visitor, IntPtr.Zero); GC.KeepAlive(visitor);
                if (error != null) throw error;
                complete = bounded; return controls.ToArray();
            }

            /// <summary>Reads bounded strings and selection state only for a standard string-backed multiselect list.</summary>
            /// <param name="list">Verified object-list control whose HWND and style were captured in the current snapshot.</param>
            private void ReadItems(AccessSaveDialogControl list)
            {
                uint style = list.Style;
                if (((style & 0x30) != 0 && (style & 0x40) == 0) || (style & 0x2000) != 0 ||
                    (style & (0x8 | 0x800)) == 0) return;
                long count = ReadMessage(list.Window, 0x18B, 0, IntPtr.Zero);
                if (count < 1 || count > 1000) return;
                var items = new List<AccessSaveDialogItem>();
                for (int index = 0; index < count; index++)
                {
                    long length = ReadMessage(list.Window, 0x18A, (ulong)index, IntPtr.Zero);
                    if (length < 0 || length > 512) return;
                    IntPtr memory = Marshal.AllocHGlobal(65536 * 2);
                    string text;
                    try
                    {
                        Marshal.WriteInt16(memory, 0);
                        long copied = ReadMessage(list.Window, 0x189, (ulong)index, memory);
                        if (copied < 0 || copied > 512 || copied != length) return;
                        text = Marshal.PtrToStringUni(memory, (int)copied);
                    }
                    finally { Marshal.FreeHGlobal(memory); }
                    long selected = ReadMessage(list.Window, 0x187, (ulong)index, IntPtr.Zero);
                    if (selected < 0) return;
                    items.Add(new AccessSaveDialogItem { Text = text, Selected = selected > 0 });
                }
                if (ReadMessage(list.Window, 0x18B, 0, IntPtr.Zero) != count) return;
                list.Items = items.ToArray(); list.ItemsComplete = true;
            }

            /// <summary>Posts one WM_COMMAND for the frozen affirmative button after rechecking its native identity.</summary>
            /// <param name="candidate">Candidate previously validated by the gate, including exact dialog and button HWNDs.</param>
            /// <returns>Whether PostMessage accepted the request; false is uncertain delivery and must not be retried.</returns>
            internal bool Enqueue(AccessSaveConfirmationCandidate candidate)
            {
                RequireWindow(candidate.Window); RequireWindow(candidate.YesButton);
                if (candidate.ProcessId != processId || candidate.ThreadId != OwnerThreadId ||
                    GetAncestor(candidate.YesButton, 2) != candidate.Window || GetDlgCtrlID(candidate.YesButton) != 1 ||
                    ClassName(candidate.YesButton) != "Button" || !IsWindowEnabled(candidate.YesButton) || !IsWindowVisible(candidate.YesButton))
                    throw new InvalidOperationException("The frozen Access affirmative button changed immediately before delivery.");
                return NativePostMessageW(candidate.Window, 0x111, new UIntPtr(1), candidate.YesButton);
            }
        }
    }
}
