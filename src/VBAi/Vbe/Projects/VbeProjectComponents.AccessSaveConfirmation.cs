using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace VBAi
{

    /// <summary>Owns the vbe project components state and operations.</summary>
    internal sealed partial class VbeProjectComponents
    {

        /// <summary>Separates the one native confirmation from saved-state verification.</summary>
        internal interface IAccessSaveConfirmation
        {

            /// <summary>Handles prepare for i access save confirmation.</summary>
            void Prepare();

            /// <summary>Requires before save for i access save confirmation.</summary>
            void RequireBeforeSave();

            /// <summary>Observes  for i access save confirmation.</summary>
            /// <returns>access save confirmation candidate produced by the operation for observe on i access save confirmation.</returns>
            AccessSaveConfirmationCandidate Observe();

            /// <summary>Handles confirm for i access save confirmation.</summary>
            /// <param name="candidate">access save confirmation candidate that supplies the candidate for this operation.</param>
            /// <param name="revalidateApprovedContext">action that supplies the revalidate approved context for this operation.</param>
            /// <param name="requireDeliveryDeadline">action that supplies the require delivery deadline for this operation.</param>
            void Confirm(AccessSaveConfirmationCandidate candidate, Action revalidateApprovedContext, Action requireDeliveryDeadline = null);

            /// <summary>Gets the confirmation attempts.</summary>
            /// <value>Current confirmation attempts exposed by i access save confirmation.</value>
            int ConfirmationAttempts { get; }

            /// <summary>Gets the confirmation queued.</summary>
            /// <value>Current confirmation queued exposed by i access save confirmation.</value>
            bool ConfirmationQueued { get; }

            /// <summary>Gets the confirmation pending.</summary>
            /// <value>Current confirmation pending exposed by i access save confirmation.</value>
            bool ConfirmationPending { get; }
        }

        /// <summary>Only explicitly approved standard modules and classes can be confirmed.</summary>
        internal sealed class AccessSaveApprovedComponent
        {

            /// <summary>Maintains the name state for access save approved component.</summary>
            internal readonly string Name;

            /// <summary>Maintains the type state for access save approved component.</summary>
            internal readonly int Type;

            /// <summary>Initializes a AccessSaveApprovedComponent instance with the supplied state.</summary>
            /// <param name="name">Text that supplies the name value. Use the format required by the calling operation.</param>
            /// <param name="type">int that supplies the type for this operation.</param>
            internal AccessSaveApprovedComponent(string name, int type) { Name = name; Type = type; }
        }

        /// <summary>Immutable native target and identity frozen by this specific save gate.</summary>
        internal sealed class AccessSaveConfirmationCandidate
        {

            /// <summary>Maintains the window and yes button state for access save confirmation candidate.</summary>
            internal readonly IntPtr Window, YesButton;

            /// <summary>Identifies the process id and thread id associated with access save confirmation candidate.</summary>
            internal readonly uint ProcessId, ThreadId;

            /// <summary>Maintains the fingerprint state for access save confirmation candidate.</summary>
            internal readonly string Fingerprint;

            /// <summary>Maintains the owner state for access save confirmation candidate.</summary>
            internal readonly object Owner;

            /// <summary>Initializes a AccessSaveConfirmationCandidate instance with the supplied state.</summary>
            /// <param name="owner">object that supplies the owner for this operation.</param>
            /// <param name="dialog">access save dialog snapshot that supplies the dialog for this operation.</param>
            /// <param name="fingerprint">Text that supplies the fingerprint value. Use the format required by the calling operation.</param>
            internal AccessSaveConfirmationCandidate(object owner, AccessSaveDialogSnapshot dialog, string fingerprint)
            {
                Owner = owner; Window = dialog.Window; ProcessId = dialog.ProcessId; ThreadId = dialog.ThreadId;
                YesButton = dialog.Controls.Single(control => control.Id == 1).Window; Fingerprint = fingerprint;
            }
        }

        /// <summary>Read-only injectable native inventory; incomplete enumeration is never an absence proof.</summary>
        internal sealed class AccessSaveDialogInventory
        {

            /// <summary>Maintains the complete state for access save dialog inventory.</summary>
            internal bool Complete;

            /// <summary>Identifies the process id and owner thread id associated with access save dialog inventory.</summary>
            internal uint ProcessId, OwnerThreadId;

            /// <summary>Maintains the dialogs state for access save dialog inventory.</summary>
            internal AccessSaveDialogSnapshot[] Dialogs;
        }

        /// <summary>Owns the access save dialog snapshot state and operations.</summary>
        internal sealed class AccessSaveDialogSnapshot
        {

            /// <summary>Maintains the window state for access save dialog snapshot.</summary>
            internal IntPtr Window;

            /// <summary>Identifies the process id and thread id and style associated with access save dialog snapshot.</summary>
            internal uint ProcessId, ThreadId, Style;

            /// <summary>Maintains the class and caption state for access save dialog snapshot.</summary>
            internal string Class, Caption;

            /// <summary>Maintains the visible and enabled and children complete state for access save dialog snapshot.</summary>
            internal bool Visible, Enabled, ChildrenComplete;

            /// <summary>Maintains the controls state for access save dialog snapshot.</summary>
            internal AccessSaveDialogControl[] Controls;
        }

        /// <summary>Owns the access save dialog control state and operations.</summary>
        internal sealed class AccessSaveDialogControl
        {

            /// <summary>Maintains the window state for access save dialog control.</summary>
            internal IntPtr Window;

            /// <summary>Identifies the process id and thread id and style associated with access save dialog control.</summary>
            internal uint ProcessId, ThreadId, Style;

            /// <summary>Identifies the id associated with access save dialog control.</summary>
            internal int Id;

            /// <summary>Maintains the class and text state for access save dialog control.</summary>
            internal string Class, Text;

            /// <summary>Maintains the visible and enabled and items complete state for access save dialog control.</summary>
            internal bool Visible, Enabled, ItemsComplete;

            /// <summary>Maintains the items state for access save dialog control.</summary>
            internal AccessSaveDialogItem[] Items;
        }

        /// <summary>Owns the access save dialog item state and operations.</summary>
        internal sealed class AccessSaveDialogItem
        {

            /// <summary>Maintains the text state for access save dialog item.</summary>
            internal string Text;

            /// <summary>Maintains the selected state for access save dialog item.</summary>
            internal bool Selected;
        }

        /// <summary>Creates a save-scoped reader on the current native VBE owner thread.</summary>
        /// <param name="vbeWindow">Native handle that supplies the vbe window for this operation.</param>
        /// <param name="processId">int that supplies the process id for this operation.</param>
        /// <param name="approvedComponents">i enumerable&lt;access save approved component&gt; that supplies the approved components for this operation.</param>
        /// <returns>i access save confirmation produced by the operation for create native access save confirmation on vbe project components.</returns>
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

            /// <summary>Maintains the read state for access save confirmation.</summary>
            private readonly Func<AccessSaveDialogInventory> read;

            /// <summary>Maintains the enqueue state for access save confirmation.</summary>
            private readonly Func<AccessSaveConfirmationCandidate, bool> enqueue;

            /// <summary>Maintains the require owner state for access save confirmation.</summary>
            private readonly Action requireOwner;

            /// <summary>Identifies the process id and owner thread id associated with access save confirmation.</summary>
            private readonly uint processId, ownerThreadId;

            /// <summary>Maintains the approved state for access save confirmation.</summary>
            private readonly HashSet<string> approved;

            /// <summary>Maintains the preparation claimed and prepared and before save claimed and confirmation claimed and faulted state for access save confirmation.</summary>
            private bool preparationClaimed, prepared, beforeSaveClaimed, confirmationClaimed, faulted;

            /// <summary>Maintains the observed state for access save confirmation.</summary>
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

            /// <summary>Initializes a AccessSaveConfirmation instance with the supplied state.</summary>
            /// <param name="read">func&lt;access save dialog inventory&gt; that supplies the read for this operation.</param>
            /// <param name="enqueue">func&lt;access save confirmation candidate, bool&gt; that supplies the enqueue for this operation.</param>
            /// <param name="requireOwner">action that supplies the require owner for this operation.</param>
            /// <param name="processId">uint that supplies the process id for this operation.</param>
            /// <param name="ownerThreadId">uint that supplies the owner thread id for this operation.</param>
            /// <param name="approvedComponents">i enumerable&lt;access save approved component&gt; that supplies the approved components for this operation.</param>
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

            /// <summary>Handles prepare for access save confirmation.</summary>
            public void Prepare()
            {
                if (preparationClaimed) throw new InvalidOperationException("Access confirmation preparation was already claimed.");
                preparationClaimed = true;
                RequireNoDialog();
                prepared = true;
            }

            /// <summary>Requires before save for access save confirmation.</summary>
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

            /// <summary>Handles confirm for access save confirmation.</summary>
            /// <param name="candidate">access save confirmation candidate that supplies the candidate for this operation.</param>
            /// <param name="revalidateApprovedContext">action that supplies the revalidate approved context for this operation.</param>
            /// <param name="requireDeliveryDeadline">action that supplies the require delivery deadline for this operation.</param>
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

            /// <summary>Requires prepared for access save confirmation.</summary>
            private void RequirePrepared()
            {
                if (!prepared || faulted) throw new InvalidOperationException("The Access confirmation gate is unprepared or refused.");
            }

            /// <summary>Reads verified inventory for access save confirmation.</summary>
            /// <returns>access save dialog inventory produced by the operation for read verified inventory on access save confirmation.</returns>
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

            /// <summary>Requires no dialog for access save confirmation.</summary>
            private void RequireNoDialog()
            {
                try
                {
                    if (ReadVerifiedInventory().Dialogs.Length != 0)
                        throw new InvalidOperationException("A preexisting owned Access dialog prevents this Save.");
                }
                catch { faulted = true; throw; }
            }

            /// <summary>Requires same candidate for access save confirmation.</summary>
            /// <param name="candidate">access save confirmation candidate that supplies the candidate for this operation.</param>
            private void RequireSameCandidate(AccessSaveConfirmationCandidate candidate)
            {
                var inventory = ReadVerifiedInventory();
                if (inventory.Dialogs.Length != 1 ||
                    Fingerprint(RequireKnownDialog(inventory)) != candidate.Fingerprint)
                    throw new InvalidOperationException("The frozen Access save confirmation candidate changed before delivery.");
            }

            /// <summary>Requires known dialog for access save confirmation.</summary>
            /// <param name="inventory">access save dialog inventory that supplies the inventory for this operation.</param>
            /// <returns>access save dialog snapshot produced by the operation for require known dialog on access save confirmation.</returns>
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

            /// <summary>Requires control for access save confirmation.</summary>
            /// <param name="dialog">access save dialog snapshot that supplies the dialog for this operation.</param>
            /// <param name="id">int that supplies the id for this operation.</param>
            /// <param name="type">Text that supplies the type value. Use the format required by the calling operation.</param>
            /// <param name="text">Text that supplies the text value. Use the format required by the calling operation.</param>
            /// <returns>access save dialog control produced by the operation for require control on access save confirmation.</returns>
            private static AccessSaveDialogControl RequireControl(AccessSaveDialogSnapshot dialog, int id, string type, string text)
            {
                var matches = dialog.Controls.Where(control => control.Id == id).ToArray();
                if (matches.Length != 1 || matches[0].Class != type || !matches[0].Visible || !matches[0].Enabled ||
                    (text != null && matches[0].Text != text))
                    throw new InvalidOperationException("The qualified Access save control " + id + " changed.");
                return matches[0];
            }

            /// <summary>Handles fingerprint for access save confirmation.</summary>
            /// <param name="dialog">access save dialog snapshot that supplies the dialog for this operation.</param>
            /// <returns>Text produced by the operation for fingerprint on access save confirmation.</returns>
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

            /// <summary>Adds fingerprint for access save confirmation.</summary>
            /// <param name="result">string builder that supplies the result for this operation.</param>
            /// <param name="values">object[] that supplies the values for this operation.</param>
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

            /// <summary>Maintains the vbe window state for native access save confirmation reader.</summary>
            private readonly IntPtr vbeWindow;

            /// <summary>Identifies the process id associated with native access save confirmation reader.</summary>
            private readonly uint processId;

            /// <summary>Identifies the owner thread id associated with native access save confirmation reader.</summary>
            internal readonly uint OwnerThreadId;

            /// <summary>Maintains the window bound and child bound and snapshot read bound milliseconds state for native access save confirmation reader.</summary>
            private const int WindowBound = 8192, ChildBound = 128, SnapshotReadBoundMilliseconds = 5000;

            /// <summary>Maintains the snapshot read watch state for native access save confirmation reader.</summary>
            private Stopwatch snapshotReadWatch;

            /// <summary>Defines the visitor callback.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <param name="state">Native handle that supplies the state for this operation.</param>
            /// <returns>Boolean indicating the result of the check for operation on native access save confirmation reader.</returns>
            private delegate bool Visitor(IntPtr window, IntPtr state);

            /// <summary>Handles native get current thread id for native access save confirmation reader.</summary>
            /// <returns>uint produced by the operation for native get current thread id on native access save confirmation reader.</returns>
            [DllImport("kernel32.dll", EntryPoint = "GetCurrentThreadId")] private static extern uint NativeGetCurrentThreadId();

            /// <summary>Handles native get current process id for native access save confirmation reader.</summary>
            /// <returns>uint produced by the operation for native get current process id on native access save confirmation reader.</returns>
            [DllImport("kernel32.dll", EntryPoint = "GetCurrentProcessId")] private static extern uint NativeGetCurrentProcessId();

            /// <summary>Handles native get window thread process id for native access save confirmation reader.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <param name="process">uint that supplies the process for this operation.</param>
            /// <returns>uint produced by the operation for native get window thread process id on native access save confirmation reader.</returns>
            [DllImport("user32.dll", EntryPoint = "GetWindowThreadProcessId")] private static extern uint NativeGetWindowThreadProcessId(IntPtr window, out uint process);

            /// <summary>Handles native get ancestor for native access save confirmation reader.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <param name="flags">uint that supplies the flags for this operation.</param>
            /// <returns>int ptr produced by the operation for native get ancestor on native access save confirmation reader.</returns>
            [DllImport("user32.dll", EntryPoint = "GetAncestor")] private static extern IntPtr NativeGetAncestor(IntPtr window, uint flags);

            /// <summary>Handles native enum windows for native access save confirmation reader.</summary>
            /// <param name="visitor">visitor that supplies the visitor for this operation.</param>
            /// <param name="state">Native handle that supplies the state for this operation.</param>
            /// <returns>Boolean indicating the result of the check for native enum windows on native access save confirmation reader.</returns>
            [DllImport("user32.dll", SetLastError = true, EntryPoint = "EnumWindows")] private static extern bool NativeEnumWindows(Visitor visitor, IntPtr state);

            /// <summary>Handles native enum child windows for native access save confirmation reader.</summary>
            /// <param name="parent">Native handle that supplies the parent for this operation.</param>
            /// <param name="visitor">visitor that supplies the visitor for this operation.</param>
            /// <param name="state">Native handle that supplies the state for this operation.</param>
            /// <returns>Boolean indicating the result of the check for native enum child windows on native access save confirmation reader.</returns>
            [DllImport("user32.dll", EntryPoint = "EnumChildWindows")] private static extern bool NativeEnumChildWindows(IntPtr parent, Visitor visitor, IntPtr state);

            /// <summary>Handles native get class name w for native access save confirmation reader.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <param name="text">string builder that supplies the text for this operation.</param>
            /// <param name="max">int that supplies the max for this operation.</param>
            /// <returns>int produced by the operation for native get class name w on native access save confirmation reader.</returns>
            [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "GetClassNameW")] private static extern int NativeGetClassNameW(IntPtr window, StringBuilder text, int max);

            /// <summary>Handles native is window visible for native access save confirmation reader.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <returns>Boolean indicating the result of the check for native is window visible on native access save confirmation reader.</returns>
            [DllImport("user32.dll", EntryPoint = "IsWindowVisible")] private static extern bool NativeIsWindowVisible(IntPtr window);

            /// <summary>Handles native is window enabled for native access save confirmation reader.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <returns>Boolean indicating the result of the check for native is window enabled on native access save confirmation reader.</returns>
            [DllImport("user32.dll", EntryPoint = "IsWindowEnabled")] private static extern bool NativeIsWindowEnabled(IntPtr window);

            /// <summary>Handles native get dlg ctrl id for native access save confirmation reader.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <returns>int produced by the operation for native get dlg ctrl id on native access save confirmation reader.</returns>
            [DllImport("user32.dll", EntryPoint = "GetDlgCtrlID")] private static extern int NativeGetDlgCtrlID(IntPtr window);

            /// <summary>Handles native get window long ptr w for native access save confirmation reader.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <param name="index">int that supplies the index for this operation.</param>
            /// <returns>int ptr produced by the operation for native get window long ptr w on native access save confirmation reader.</returns>
            [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")] private static extern IntPtr NativeGetWindowLongPtrW(IntPtr window, int index);

            /// <summary>Handles native send message w for native access save confirmation reader.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <param name="message">uint that supplies the message for this operation.</param>
            /// <param name="wParam">Native handle that supplies the w param for this operation.</param>
            /// <param name="lParam">Native handle that supplies the l param for this operation.</param>
            /// <returns>int ptr produced by the operation for native send message w on native access save confirmation reader.</returns>
            [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "SendMessageW")] private static extern IntPtr NativeSendMessageW(IntPtr window, uint message, UIntPtr wParam, IntPtr lParam);

            /// <summary>Handles native post message w for native access save confirmation reader.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <param name="message">uint that supplies the message for this operation.</param>
            /// <param name="wParam">Native handle that supplies the w param for this operation.</param>
            /// <param name="lParam">Native handle that supplies the l param for this operation.</param>
            /// <returns>Boolean indicating the result of the check for native post message w on native access save confirmation reader.</returns>
            [DllImport("user32.dll", SetLastError = true, EntryPoint = "PostMessageW")] private static extern bool NativePostMessageW(IntPtr window, uint message, UIntPtr wParam, IntPtr lParam);


            // This bounds completed snapshot work; same-thread SendMessageW is not
            // represented as a cancellable call or a guaranteed wall-clock timeout.
            /// <summary>Requires read budget for native access save confirmation reader.</summary>
            private void RequireReadBudget()
            {
                if (snapshotReadWatch != null && snapshotReadWatch.ElapsedMilliseconds >= SnapshotReadBoundMilliseconds)
                    throw new InvalidOperationException("Access confirmation snapshot exceeded its read budget; no confirmation is queued.");
            }

            /// <summary>Handles checked read for native access save confirmation reader.</summary>
            /// <typeparam name="T">The type used for t.</typeparam>
            /// <param name="getter">func&lt;t&gt; that supplies the getter for this operation.</param>
            /// <returns>t produced by the operation for checked read on native access save confirmation reader.</returns>
            private T CheckedRead<T>(Func<T> getter)
            {
                RequireReadBudget(); T result = getter(); RequireReadBudget(); return result;
            }

            /// <summary>Returns current thread id for native access save confirmation reader.</summary>
            /// <returns>uint produced by the operation for get current thread id on native access save confirmation reader.</returns>
            private uint GetCurrentThreadId() => CheckedRead(() => NativeGetCurrentThreadId());

            /// <summary>Returns current process id for native access save confirmation reader.</summary>
            /// <returns>uint produced by the operation for get current process id on native access save confirmation reader.</returns>
            private uint GetCurrentProcessId() => CheckedRead(() => NativeGetCurrentProcessId());

            /// <summary>Returns window thread process id for native access save confirmation reader.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <param name="process">uint that supplies the process for this operation.</param>
            /// <returns>uint produced by the operation for get window thread process id on native access save confirmation reader.</returns>
            private uint GetWindowThreadProcessId(IntPtr window, out uint process)
            {
                RequireReadBudget(); uint thread = NativeGetWindowThreadProcessId(window, out process); RequireReadBudget(); return thread;
            }

            /// <summary>Returns ancestor for native access save confirmation reader.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <param name="flags">uint that supplies the flags for this operation.</param>
            /// <returns>int ptr produced by the operation for get ancestor on native access save confirmation reader.</returns>
            private IntPtr GetAncestor(IntPtr window, uint flags) => CheckedRead(() => NativeGetAncestor(window, flags));

            /// <summary>Handles enum windows for native access save confirmation reader.</summary>
            /// <param name="visitor">visitor that supplies the visitor for this operation.</param>
            /// <param name="state">Native handle that supplies the state for this operation.</param>
            /// <returns>Boolean indicating the result of the check for enum windows on native access save confirmation reader.</returns>
            private bool EnumWindows(Visitor visitor, IntPtr state) => CheckedRead(() => NativeEnumWindows(visitor, state));

            /// <summary>Handles enum child windows for native access save confirmation reader.</summary>
            /// <param name="parent">Native handle that supplies the parent for this operation.</param>
            /// <param name="visitor">visitor that supplies the visitor for this operation.</param>
            /// <param name="state">Native handle that supplies the state for this operation.</param>
            /// <returns>Boolean indicating the result of the check for enum child windows on native access save confirmation reader.</returns>
            private bool EnumChildWindows(IntPtr parent, Visitor visitor, IntPtr state) => CheckedRead(() => NativeEnumChildWindows(parent, visitor, state));

            /// <summary>Returns class name w for native access save confirmation reader.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <param name="text">string builder that supplies the text for this operation.</param>
            /// <param name="max">int that supplies the max for this operation.</param>
            /// <returns>int produced by the operation for get class name w on native access save confirmation reader.</returns>
            private int GetClassNameW(IntPtr window, StringBuilder text, int max) => CheckedRead(() => NativeGetClassNameW(window, text, max));

            /// <summary>Determines whether window visible for native access save confirmation reader.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <returns>Boolean indicating the result of the check for is window visible on native access save confirmation reader.</returns>
            private bool IsWindowVisible(IntPtr window) => CheckedRead(() => NativeIsWindowVisible(window));

            /// <summary>Determines whether window enabled for native access save confirmation reader.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <returns>Boolean indicating the result of the check for is window enabled on native access save confirmation reader.</returns>
            private bool IsWindowEnabled(IntPtr window) => CheckedRead(() => NativeIsWindowEnabled(window));

            /// <summary>Returns dlg ctrl id for native access save confirmation reader.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <returns>int produced by the operation for get dlg ctrl id on native access save confirmation reader.</returns>
            private int GetDlgCtrlID(IntPtr window) => CheckedRead(() => NativeGetDlgCtrlID(window));

            /// <summary>Returns window long ptr w for native access save confirmation reader.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <param name="index">int that supplies the index for this operation.</param>
            /// <returns>int ptr produced by the operation for get window long ptr w on native access save confirmation reader.</returns>
            private IntPtr GetWindowLongPtrW(IntPtr window, int index) => CheckedRead(() => NativeGetWindowLongPtrW(window, index));

            /// <summary>Handles send message w for native access save confirmation reader.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <param name="message">uint that supplies the message for this operation.</param>
            /// <param name="wParam">Native handle that supplies the w param for this operation.</param>
            /// <param name="lParam">Native handle that supplies the l param for this operation.</param>
            /// <returns>int ptr produced by the operation for send message w on native access save confirmation reader.</returns>
            private IntPtr SendMessageW(IntPtr window, uint message, UIntPtr wParam, IntPtr lParam) =>
                CheckedRead(() => NativeSendMessageW(window, message, wParam, lParam));

            /// <summary>Initializes a NativeAccessSaveConfirmationReader instance with the supplied state.</summary>
            /// <param name="vbeWindow">Native handle that supplies the vbe window for this operation.</param>
            /// <param name="processId">int that supplies the process id for this operation.</param>
            internal NativeAccessSaveConfirmationReader(IntPtr vbeWindow, int processId)
            {
                if (processId <= 0 || vbeWindow == IntPtr.Zero) throw new ArgumentException("Access confirmation requires the native VBE HWND and host PID.");
                this.vbeWindow = vbeWindow; this.processId = (uint)processId; OwnerThreadId = GetCurrentThreadId();
                RequireOwner();
            }

            /// <summary>Requires owner for native access save confirmation reader.</summary>
            internal void RequireOwner()
            {
                uint actual;
                if (GetCurrentProcessId() != processId || GetCurrentThreadId() != OwnerThreadId ||
                    GetWindowThreadProcessId(vbeWindow, out actual) != OwnerThreadId || actual != processId)
                    throw new InvalidOperationException("Access save confirmation left the exact native VBE process/UI thread.");
            }

            /// <summary>Requires window for native access save confirmation reader.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            private void RequireWindow(IntPtr window)
            {
                RequireOwner();
                uint actual, rootPid; IntPtr root = GetAncestor(window, 2);
                if (window == IntPtr.Zero || GetWindowThreadProcessId(window, out actual) != OwnerThreadId ||
                    actual != processId || GetWindowThreadProcessId(root, out rootPid) != OwnerThreadId || rootPid != processId)
                    throw new InvalidOperationException("The Access confirmation window or root identity changed.");
            }

            /// <summary>Handles class name for native access save confirmation reader.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <returns>Text produced by the operation for class name on native access save confirmation reader.</returns>
            private string ClassName(IntPtr window)
            {
                var text = new StringBuilder(256);
                if (GetClassNameW(window, text, text.Capacity) <= 0) throw new InvalidOperationException("Access native control class is unavailable.");
                return text.ToString();
            }

            /// <summary>Handles style for native access save confirmation reader.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <returns>uint produced by the operation for style on native access save confirmation reader.</returns>
            private uint Style(IntPtr window) { RequireWindow(window); return unchecked((uint)GetWindowLongPtrW(window, -16).ToInt64()); }

            /// <summary>Reads message for native access save confirmation reader.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <param name="message">uint that supplies the message for this operation.</param>
            /// <param name="wParam">ulong that supplies the w param for this operation.</param>
            /// <param name="buffer">Native handle that supplies the buffer for this operation.</param>
            /// <returns>long produced by the operation for read message on native access save confirmation reader.</returns>
            private long ReadMessage(IntPtr window, uint message, ulong wParam, IntPtr buffer)
            {
                RequireWindow(window);
                return SendMessageW(window, message, new UIntPtr(wParam), buffer).ToInt64();
            }

            /// <summary>Handles text for native access save confirmation reader.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
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

            /// <summary>Reads children for native access save confirmation reader.</summary>
            /// <param name="parent">Native handle that supplies the parent for this operation.</param>
            /// <param name="complete">Indicates whether complete is enabled.</param>
            /// <returns>access save dialog control[] produced by the operation for read children on native access save confirmation reader.</returns>
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

            /// <summary>Reads items for native access save confirmation reader.</summary>
            /// <param name="list">access save dialog control that supplies the list for this operation.</param>
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

            /// <summary>Handles enqueue for native access save confirmation reader.</summary>
            /// <param name="candidate">access save confirmation candidate that supplies the candidate for this operation.</param>
            /// <returns>Boolean indicating the result of the check for enqueue on native access save confirmation reader.</returns>
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
