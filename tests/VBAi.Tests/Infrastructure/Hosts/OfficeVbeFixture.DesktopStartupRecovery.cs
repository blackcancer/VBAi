using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace VBAi.Tests.Integration
{
    internal sealed partial class OfficeVbeFixture
    {
        internal const string FrenchPublisherRecoveryMessage = "Publisher n'a pas pu démarrer la dernière fois. Le mode sans échec permet de résoudre le problème, mais certaines fonctionnalités risquent de ne pas être disponibles sous ce mode.\n\nVoulez-vous démarrer en mode sans échec\u00A0?";
        private PublisherStartupRecoveryGate publisherStartupRecovery;

        internal sealed class PublisherRecoverySnapshot
        {
            internal IntPtr Dialog, Button, ButtonRoot;
            internal uint DialogPid, ButtonPid, DialogThread, ButtonThread;
            internal int ButtonId;
            internal string DialogClass, ButtonClass, Caption, Message, ButtonText;
        }

        /// <summary>Only the exact observed French recovery question and the native No button qualify.</summary>
        internal static bool IsExactPublisherRecovery(PublisherRecoverySnapshot value, uint expectedPid)
        {
            return value != null && expectedPid != 0 && value.Dialog != IntPtr.Zero && value.Button != IntPtr.Zero &&
                value.ButtonRoot == value.Dialog && value.DialogPid == expectedPid && value.ButtonPid == expectedPid &&
                value.DialogThread != 0 && value.ButtonThread == value.DialogThread && value.ButtonId == 7 &&
                value.DialogClass == "#32770" && value.ButtonClass == "Button" && value.Caption == "Microsoft Publisher" &&
                value.Message != null && value.Message.Replace("\r\n", "\n") == FrenchPublisherRecoveryMessage &&
                (value.ButtonText == "Non" || value.ButtonText == "&Non");
        }

        /// <summary>One startup choice per original native process; an uncertain click permanently forbids replay.</summary>
        internal sealed class PublisherStartupRecoveryGate
        {
            private readonly uint pid;
            private readonly long originalHandle;
            internal bool Claimed { get; private set; }
            internal IntPtr Dialog { get; private set; }
            internal uint DialogThread { get; private set; }
            internal PublisherStartupRecoveryGate(uint pid, long originalHandle)
            {
                if (pid == 0 || originalHandle == 0) throw new ArgumentException("Original Publisher PID and native process handle are required.");
                this.pid = pid; this.originalHandle = originalHandle;
            }

            internal void RequestNormalStart(PublisherRecoverySnapshot initial, Action guard, Func<PublisherRecoverySnapshot> reread,
                Func<bool> clickOnce, Action<IDictionary<string, object>> record)
            {
                if (Claimed) throw new InvalidOperationException("The original Publisher startup decision was already claimed; its native click cannot be replayed.");
                if (!IsExactPublisherRecovery(initial, pid)) throw new InvalidOperationException("Unknown Publisher startup dialog; no action is permitted.");
                guard();
                var current = reread();
                if (!IsExactPublisherRecovery(current, pid) || current.Dialog != initial.Dialog || current.Button != initial.Button ||
                    current.DialogThread != initial.DialogThread || current.ButtonThread != initial.ButtonThread)
                    throw new InvalidOperationException("Publisher recovery identity or message changed before the single click; no action is permitted.");
                guard();
                Claimed = true; Dialog = current.Dialog; DialogThread = current.DialogThread;
                Action<string, bool> write = (state, uncertain) => record(new Dictionary<string, object> {
                    ["PublisherStartupRecovery"] = state, ["ProcessId"] = pid, ["OriginalHandle"] = originalHandle,
                    ["DialogHwnd"] = Dialog.ToInt64(), ["DialogThreadId"] = DialogThread, ["ButtonHwnd"] = current.Button.ToInt64(),
                    ["ButtonId"] = 7, ["Locale"] = "fr-FR", ["NativeClickAttempts"] = 1,
                    ["NormalStartRequested"] = true, ["NormalModeProven"] = false, ["OutcomeUncertain"] = uncertain,
                    ["ClickTimeoutMilliseconds"] = 2000, ["AutomaticRetry"] = false, ["Utc"] = DateTime.UtcNow.ToString("o") });
                write("CLICK_PENDING", false); // Durable intent precedes the only native mutation.
                bool returned;
                try { returned = clickOnce(); }
                catch (Exception error)
                {
                    try { write("CLICK_UNCERTAIN", true); }
                    catch (Exception evidence) { throw new AggregateException("Startup click and uncertainty evidence both failed; host retained.", error, evidence); }
                    throw new InvalidOperationException("Publisher startup click threw; its outcome is uncertain and cannot be retried.", error);
                }
                if (!returned) { write("CLICK_UNCERTAIN", true); throw new InvalidOperationException("Publisher startup click did not return; its outcome is uncertain and cannot be retried."); }
                write("CLICK_RETURNED", false);
            }
        }

        private void RecordPublisherStartup(IDictionary<string, object> row)
        { row["Desktop"] = privateDesktop; steps.Add(row); FlushAdapterEvidence(); }

        private string ReadPublisherStartupText(IntPtr window, string role)
        {
            RequirePrivateHostDesktop(true);
            IsolatedTestDesktop.RequireOfficeWindowInventory(privateDesktop, (uint)ProcessId, true, window);
            RecordPublisherStartup(new Dictionary<string, object> { ["PublisherStartupText"] = role, ["State"] = "PENDING",
                ["ProcessId"] = ProcessId, ["Hwnd"] = window.ToInt64(), ["ReadTimeoutMilliseconds"] = 500, ["ReadAttempts"] = 1 });
            var text = new StringBuilder(4096); UIntPtr returned;
            bool complete = ReadDialogText(window, 0x000D, new UIntPtr(4096), text, 0x23, 500, out returned) != IntPtr.Zero;
            RecordPublisherStartup(new Dictionary<string, object> { ["PublisherStartupText"] = role, ["State"] = complete ? "RETURNED" : "UNCERTAIN",
                ["ProcessId"] = ProcessId, ["Hwnd"] = window.ToInt64(), ["ReadTimeoutMilliseconds"] = 500,
                ["ReturnedCharacters"] = returned.ToUInt64(), ["ReadAttempts"] = 1 });
            if (!complete || returned.ToUInt64() >= 4095) throw new InvalidOperationException("Owned Publisher startup text is unavailable or truncated; no native action is permitted.");
            return text.ToString();
        }

        private string PublisherStartupClass(IntPtr window)
        { var name = new StringBuilder(256); return GetClassName(window, name, name.Capacity) > 0 ? name.ToString() : null; }

        private PublisherRecoverySnapshot CapturePublisherRecovery(IntPtr dialog)
        {
            RequirePrivateHostDesktop(true);
            IsolatedTestDesktop.RequireOfficeWindowInventory(privateDesktop, (uint)ProcessId, true, dialog);
            IntPtr button = GetDlgItem(dialog, 7);
            var value = new PublisherRecoverySnapshot { Dialog = dialog, Button = button, DialogClass = PublisherStartupClass(dialog),
                ButtonClass = PublisherStartupClass(button), ButtonId = GetDlgCtrlID(button), ButtonRoot = GetAncestor(button, 2) };
            value.DialogThread = GetWindowThreadProcessId(dialog, out value.DialogPid);
            value.ButtonThread = GetWindowThreadProcessId(button, out value.ButtonPid);
            if (value.DialogPid != (uint)ProcessId || value.ButtonPid != (uint)ProcessId || value.Button == IntPtr.Zero ||
                value.DialogClass != "#32770" || value.ButtonClass != "Button" || value.ButtonId != 7 ||
                value.ButtonRoot != dialog || value.DialogThread == 0 || value.ButtonThread != value.DialogThread || !IsWindowEnabled(button))
                throw new InvalidOperationException("Owned Publisher startup native identity differs; no action is permitted.");
            value.Caption = ReadPublisherStartupText(dialog, "DialogCaption");
            value.ButtonText = ReadPublisherStartupText(button, "NoButtonCaption");
            var statics = new List<IntPtr>(); bool complete = true;
            // Win32 documents this return value as unused; validate callbacks and exact live controls.
            EnumChildWindows(dialog, (window, parameter) => {
                uint owner; GetWindowThreadProcessId(window, out owner);
                if (owner == (uint)ProcessId && PublisherStartupClass(window) == "Static") statics.Add(window);
                if (statics.Count > 16) { complete = false; return false; }
                return true;
            }, IntPtr.Zero);
            if (!complete || statics.Count == 0) throw new InvalidOperationException("Publisher startup static control inventory is incomplete.");
            var messages = statics.Select(window => ReadPublisherStartupText(window, "StaticMessage")).Where(text => text.Length != 0).ToArray();
            if (messages.Length != 1) throw new InvalidOperationException("Publisher startup has no unique exact message; no native action is permitted.");
            value.Message = messages[0];
            return value;
        }

        /// <returns>True only while the already-clicked dialog is still visible, so ROT discovery waits without replay.</returns>
        private bool ObservePrivatePublisherStartupRecovery()
        {
            if (Kind != "Publisher") return false;
            RequirePrivateHostDesktop(false);
            var dialogs = new List<IntPtr>(); bool complete = EnumWindows((window, parameter) => {
                uint owner; GetWindowThreadProcessId(window, out owner);
                if (owner == (uint)ProcessId && IsWindowVisible(window) && PublisherStartupClass(window) == "#32770") dialogs.Add(window);
                return dialogs.Count <= 8;
            }, IntPtr.Zero);
            if (!complete || dialogs.Count > 1) throw new InvalidOperationException("Publisher has an incomplete or ambiguous owned startup dialog inventory; no action is permitted.");
            if (dialogs.Count == 0) return false;
            IntPtr dialog = dialogs[0];
            if (publisherStartupRecovery.Claimed)
            {
                uint owner; uint thread = GetWindowThreadProcessId(dialog, out owner);
                if (dialog != publisherStartupRecovery.Dialog || owner != (uint)ProcessId || thread != publisherStartupRecovery.DialogThread)
                    throw new InvalidOperationException("Another Publisher startup dialog appeared after the sole recovery choice; it is not handled.");
                return true;
            }
            var snapshot = CapturePublisherRecovery(dialog);
            publisherStartupRecovery.RequestNormalStart(snapshot, () => RequirePrivateHostDesktop(true),
                () => CapturePublisherRecovery(dialog), () => {
                    RequirePrivateHostDesktop(true);
                    IsolatedTestDesktop.RequireOfficeWindowInventory(privateDesktop, (uint)ProcessId, true, snapshot.Button);
                    uint dialogPid, buttonPid;
                    uint dialogThread = GetWindowThreadProcessId(dialog, out dialogPid), buttonThread = GetWindowThreadProcessId(snapshot.Button, out buttonPid);
                    if (GetDlgItem(dialog, 7) != snapshot.Button || GetDlgCtrlID(snapshot.Button) != 7 || !IsWindowEnabled(snapshot.Button) ||
                        dialogPid != snapshot.DialogPid || buttonPid != snapshot.ButtonPid || dialogThread != snapshot.DialogThread ||
                        buttonThread != snapshot.ButtonThread || GetAncestor(snapshot.Button, 2) != dialog ||
                        PublisherStartupClass(dialog) != "#32770" || PublisherStartupClass(snapshot.Button) != "Button")
                        throw new InvalidOperationException("Publisher No button changed before its single native click.");
                    UIntPtr result;
                    return SendDialogCommand(snapshot.Button, 0x00F5, UIntPtr.Zero, IntPtr.Zero, 0x23, 2000, out result) != IntPtr.Zero;
                }, RecordPublisherStartup);
            return true;
        }
    }
}
