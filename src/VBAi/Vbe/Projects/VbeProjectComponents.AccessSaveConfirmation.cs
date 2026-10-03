using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace VBAi
{
    internal sealed partial class VbeProjectComponents
    {
        /// <summary>Separates the one native confirmation from saved-state verification.</summary>
        internal interface IAccessSaveConfirmation
        {
            void Prepare();
            void RequireBeforeSave();
            AccessSaveConfirmationCandidate Observe();
            void Confirm(AccessSaveConfirmationCandidate candidate, Action revalidateApprovedContext, Action requireDeliveryDeadline = null);
            int ConfirmationAttempts { get; }
            bool ConfirmationQueued { get; }
            bool ConfirmationPending { get; }
        }

        /// <summary>Only explicitly approved standard modules and classes can be confirmed.</summary>
        internal sealed class AccessSaveApprovedComponent
        {
            internal readonly string Name;
            internal readonly int Type;
            internal AccessSaveApprovedComponent(string name, int type) { Name = name; Type = type; }
        }

        /// <summary>Immutable native target and identity frozen by this specific save gate.</summary>
        internal sealed class AccessSaveConfirmationCandidate
        {
            internal readonly IntPtr Window, YesButton;
            internal readonly uint ProcessId, ThreadId;
            internal readonly string Fingerprint;
            internal readonly object Owner;
            internal AccessSaveConfirmationCandidate(object owner, AccessSaveDialogSnapshot dialog, string fingerprint)
            {
                Owner = owner; Window = dialog.Window; ProcessId = dialog.ProcessId; ThreadId = dialog.ThreadId;
                YesButton = dialog.Controls.Single(control => control.Id == 1).Window; Fingerprint = fingerprint;
            }
        }

        /// <summary>Read-only injectable native inventory; incomplete enumeration is never an absence proof.</summary>
        internal sealed class AccessSaveDialogInventory
        {
            internal bool Complete;
            internal uint ProcessId, OwnerThreadId;
            internal AccessSaveDialogSnapshot[] Dialogs;
        }

        internal sealed class AccessSaveDialogSnapshot
        {
            internal IntPtr Window;
            internal uint ProcessId, ThreadId, Style;
            internal string Class, Caption;
            internal bool Visible, Enabled, ChildrenComplete;
            internal AccessSaveDialogControl[] Controls;
        }

        internal sealed class AccessSaveDialogControl
        {
            internal IntPtr Window;
            internal uint ProcessId, ThreadId, Style;
            internal int Id;
            internal string Class, Text;
            internal bool Visible, Enabled, ItemsComplete;
            internal AccessSaveDialogItem[] Items;
        }

        internal sealed class AccessSaveDialogItem
        {
            internal string Text;
            internal bool Selected;
        }

        /// <summary>Creates a save-scoped reader on the current native VBE owner thread.</summary>
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
            private readonly Func<AccessSaveDialogInventory> read;
            private readonly Func<AccessSaveConfirmationCandidate, bool> enqueue;
            private readonly Action requireOwner;
            private readonly uint processId, ownerThreadId;
            private readonly HashSet<string> approved;
            private bool preparationClaimed, prepared, beforeSaveClaimed, confirmationClaimed, faulted;
            private AccessSaveConfirmationCandidate observed;
            public int ConfirmationAttempts { get; private set; }
            public bool ConfirmationQueued { get; private set; }
            public bool ConfirmationPending { get; private set; }

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

            public void Prepare()
            {
                if (preparationClaimed) throw new InvalidOperationException("Access confirmation preparation was already claimed.");
                preparationClaimed = true;
                RequireNoDialog();
                prepared = true;
            }

            public void RequireBeforeSave()
            {
                RequirePrepared();
                if (beforeSaveClaimed) throw new InvalidOperationException("The Access pre-Save confirmation gate was already consumed.");
                beforeSaveClaimed = true;
                RequireNoDialog(); // No preexisting dialog may be accepted as the result of this Save.
            }

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

            private void RequirePrepared()
            {
                if (!prepared || faulted) throw new InvalidOperationException("The Access confirmation gate is unprepared or refused.");
            }

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

            private void RequireNoDialog()
            {
                try
                {
                    if (ReadVerifiedInventory().Dialogs.Length != 0)
                        throw new InvalidOperationException("A preexisting owned Access dialog prevents this Save.");
                }
                catch { faulted = true; throw; }
            }

            private void RequireSameCandidate(AccessSaveConfirmationCandidate candidate)
            {
                var inventory = ReadVerifiedInventory();
                if (inventory.Dialogs.Length != 1 ||
                    Fingerprint(RequireKnownDialog(inventory)) != candidate.Fingerprint)
                    throw new InvalidOperationException("The frozen Access save confirmation candidate changed before delivery.");
            }

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

            private static AccessSaveDialogControl RequireControl(AccessSaveDialogSnapshot dialog, int id, string type, string text)
            {
                var matches = dialog.Controls.Where(control => control.Id == id).ToArray();
                if (matches.Length != 1 || matches[0].Class != type || !matches[0].Visible || !matches[0].Enabled ||
                    (text != null && matches[0].Text != text))
                    throw new InvalidOperationException("The qualified Access save control " + id + " changed.");
                return matches[0];
            }

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
            private readonly IntPtr vbeWindow;
            private readonly uint processId;
            internal readonly uint OwnerThreadId;
            private const int WindowBound = 8192, ChildBound = 128, SnapshotReadBoundMilliseconds = 5000;
            private Stopwatch snapshotReadWatch;
            private delegate bool Visitor(IntPtr window, IntPtr state);
            [DllImport("kernel32.dll", EntryPoint = "GetCurrentThreadId")] private static extern uint NativeGetCurrentThreadId();
            [DllImport("kernel32.dll", EntryPoint = "GetCurrentProcessId")] private static extern uint NativeGetCurrentProcessId();
            [DllImport("user32.dll", EntryPoint = "GetWindowThreadProcessId")] private static extern uint NativeGetWindowThreadProcessId(IntPtr window, out uint process);
            [DllImport("user32.dll", EntryPoint = "GetAncestor")] private static extern IntPtr NativeGetAncestor(IntPtr window, uint flags);
            [DllImport("user32.dll", SetLastError = true, EntryPoint = "EnumWindows")] private static extern bool NativeEnumWindows(Visitor visitor, IntPtr state);
            [DllImport("user32.dll", EntryPoint = "EnumChildWindows")] private static extern bool NativeEnumChildWindows(IntPtr parent, Visitor visitor, IntPtr state);
            [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "GetClassNameW")] private static extern int NativeGetClassNameW(IntPtr window, StringBuilder text, int max);
            [DllImport("user32.dll", EntryPoint = "IsWindowVisible")] private static extern bool NativeIsWindowVisible(IntPtr window);
            [DllImport("user32.dll", EntryPoint = "IsWindowEnabled")] private static extern bool NativeIsWindowEnabled(IntPtr window);
            [DllImport("user32.dll", EntryPoint = "GetDlgCtrlID")] private static extern int NativeGetDlgCtrlID(IntPtr window);
            [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")] private static extern IntPtr NativeGetWindowLongPtrW(IntPtr window, int index);
            [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "SendMessageW")] private static extern IntPtr NativeSendMessageW(IntPtr window, uint message, UIntPtr wParam, IntPtr lParam);
            [DllImport("user32.dll", SetLastError = true, EntryPoint = "PostMessageW")] private static extern bool NativePostMessageW(IntPtr window, uint message, UIntPtr wParam, IntPtr lParam);


            // This bounds completed snapshot work; same-thread SendMessageW is not
            // represented as a cancellable call or a guaranteed wall-clock timeout.
            private void RequireReadBudget()
            {
                if (snapshotReadWatch != null && snapshotReadWatch.ElapsedMilliseconds >= SnapshotReadBoundMilliseconds)
                    throw new InvalidOperationException("Access confirmation snapshot exceeded its read budget; no confirmation is queued.");
            }
            private T CheckedRead<T>(Func<T> getter)
            {
                RequireReadBudget(); T result = getter(); RequireReadBudget(); return result;
            }
            private uint GetCurrentThreadId() => CheckedRead(() => NativeGetCurrentThreadId());
            private uint GetCurrentProcessId() => CheckedRead(() => NativeGetCurrentProcessId());
            private uint GetWindowThreadProcessId(IntPtr window, out uint process)
            {
                RequireReadBudget(); uint thread = NativeGetWindowThreadProcessId(window, out process); RequireReadBudget(); return thread;
            }
            private IntPtr GetAncestor(IntPtr window, uint flags) => CheckedRead(() => NativeGetAncestor(window, flags));
            private bool EnumWindows(Visitor visitor, IntPtr state) => CheckedRead(() => NativeEnumWindows(visitor, state));
            private bool EnumChildWindows(IntPtr parent, Visitor visitor, IntPtr state) => CheckedRead(() => NativeEnumChildWindows(parent, visitor, state));
            private int GetClassNameW(IntPtr window, StringBuilder text, int max) => CheckedRead(() => NativeGetClassNameW(window, text, max));
            private bool IsWindowVisible(IntPtr window) => CheckedRead(() => NativeIsWindowVisible(window));
            private bool IsWindowEnabled(IntPtr window) => CheckedRead(() => NativeIsWindowEnabled(window));
            private int GetDlgCtrlID(IntPtr window) => CheckedRead(() => NativeGetDlgCtrlID(window));
            private IntPtr GetWindowLongPtrW(IntPtr window, int index) => CheckedRead(() => NativeGetWindowLongPtrW(window, index));
            private IntPtr SendMessageW(IntPtr window, uint message, UIntPtr wParam, IntPtr lParam) =>
                CheckedRead(() => NativeSendMessageW(window, message, wParam, lParam));

            internal NativeAccessSaveConfirmationReader(IntPtr vbeWindow, int processId)
            {
                if (processId <= 0 || vbeWindow == IntPtr.Zero) throw new ArgumentException("Access confirmation requires the native VBE HWND and host PID.");
                this.vbeWindow = vbeWindow; this.processId = (uint)processId; OwnerThreadId = GetCurrentThreadId();
                RequireOwner();
            }

            internal void RequireOwner()
            {
                uint actual;
                if (GetCurrentProcessId() != processId || GetCurrentThreadId() != OwnerThreadId ||
                    GetWindowThreadProcessId(vbeWindow, out actual) != OwnerThreadId || actual != processId)
                    throw new InvalidOperationException("Access save confirmation left the exact native VBE process/UI thread.");
            }

            private void RequireWindow(IntPtr window)
            {
                RequireOwner();
                uint actual, rootPid; IntPtr root = GetAncestor(window, 2);
                if (window == IntPtr.Zero || GetWindowThreadProcessId(window, out actual) != OwnerThreadId ||
                    actual != processId || GetWindowThreadProcessId(root, out rootPid) != OwnerThreadId || rootPid != processId)
                    throw new InvalidOperationException("The Access confirmation window or root identity changed.");
            }

            private string ClassName(IntPtr window)
            {
                var text = new StringBuilder(256);
                if (GetClassNameW(window, text, text.Capacity) <= 0) throw new InvalidOperationException("Access native control class is unavailable.");
                return text.ToString();
            }

            private uint Style(IntPtr window) { RequireWindow(window); return unchecked((uint)GetWindowLongPtrW(window, -16).ToInt64()); }

            private long ReadMessage(IntPtr window, uint message, ulong wParam, IntPtr buffer)
            {
                RequireWindow(window);
                return SendMessageW(window, message, new UIntPtr(wParam), buffer).ToInt64();
            }

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
