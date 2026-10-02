using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    /// <summary>Discards only fixture-created Access objects during one owned, settled teardown.</summary>
    internal sealed partial class OfficeVbeFixture
    {
        private readonly HashSet<string> createdAccessDiscardObjects = new HashSet<string>(StringComparer.Ordinal);
        internal Action ValidateAccessDiscardDocument;
        internal Action<AccessDiscardLease> ScanAccessDiscard = ScanNativeAccessDiscard;
        internal Func<IntPtr, Process, string, bool> IsVisibleOwnedAccessDiscard = NativeVisibleOwnedAccessDiscard;

        private void TrackCreatedAccessDiscardObject(string command, IDictionary<string, object> request, IDictionary<string, object> reply)
        {
            if (Kind != "Access" || !Equals(reply["Ok"], true) ||
                (command != "create_module" && command != "create_class" && command != "create_form")) return;
            string key = command == "create_form" ? "Form" : "Module";
            var result = VbeBridgeClient.Object(reply["Data"]);
            if (Project != null && request.TryGetValue("Project", out object requestedProject) && Equals(requestedProject, Project) &&
                result.TryGetValue("Project", out object actualProject) && Equals(actualProject, Project) &&
                request.TryGetValue(key, out object requestedName) && requestedName is string name && !string.IsNullOrWhiteSpace(name) &&
                result.TryGetValue(key, out object actualName) && Equals(actualName, name))
                createdAccessDiscardObjects.Add(name);
        }

        private void RequireAccessDiscardDocument()
        {
            if (ValidateAccessDiscardDocument != null) ValidateAccessDiscardDocument();
            else RequireOwnedDocument();
        }

        /// <summary>Preserves the existing one-shot close; only its expected naming prompt is cancelled.</summary>
        internal void CloseAccessDatabaseDiscarding()
        {
            Assert.AreEqual(shutdownOwnerThread, Thread.CurrentThread.ManagedThreadId);
            commandContainment.RequireTerminal();
            Assert.AreEqual("Access", Kind); Assert.IsFalse(NativeExecutionUnsettled);
            Assert.IsFalse(hostTeardownRefused); Assert.IsTrue(owned);
            Assert.IsNotNull(shutdownEvidence); Assert.AreEqual(true, shutdownEvidence.Record["TeardownPrepared"]);
            RequireAccessDiscardDocument();
            StopAccessSaveDialogHandler();
            var process = ownedProcess;
            Assert.IsNotNull(process); Assert.AreEqual(ProcessId, process.Id); Assert.IsFalse(process.HasExited);
            string started = process.StartTime.ToUniversalTime().ToString("o");
            string handle = "0x" + unchecked((ulong)process.Handle.ToInt64()).ToString("X16");
            Assert.AreEqual(shutdownEvidence.Record["ProcessStartedUtc"], started);
            Assert.AreEqual(shutdownEvidence.Record["OriginalProcessHandle"], handle);
            string path = DocumentPath, project = Project;
            var worker = new OwnedDialogWorker(process, Kind);
            var lease = new AccessDiscardLease(process.Id, started, handle, path, project, shutdownOwnerThread,
                createdAccessDiscardObjects, () => !worker.StopRequested && !hostTeardownRefused && owned &&
                    !commandContainment.Pending && !commandContainment.Uncertain && !NativeExecutionUnsettled &&
                    ReferenceEquals(process, ownedProcess) && ProcessId == process.Id && DocumentPath == path && Project == project &&
                    !process.HasExited && process.StartTime.ToUniversalTime().ToString("o") == started &&
                    "0x" + unchecked((ulong)process.Handle.ToInt64()).ToString("X16") == handle,
                record => File.WriteAllText(Path.Combine(Root, "access-discard-dialogs.json"), new JavaScriptSerializer().Serialize(record)));
            lease.Stopping = () => worker.StopRequested;
            lease.VisibleOwnedDialog = window => IsVisibleOwnedAccessDiscard(window, process, started);
            worker.Discard = lease;
            dialogWorker = worker;
            worker.Thread = new Thread(() => {
                try
                {
                    lease.BindWorkerThread();
                    while (!worker.StopRequested && worker.HostAlive)
                    {
                        ScanAccessDiscard(lease);
                        if (!worker.StopRequested) Thread.Sleep(25);
                    }
                }
                catch (Exception error) { worker.Failure = error; worker.StopRequested = true; }
            }) { IsBackground = true };
            worker.Thread.SetApartmentState(ApartmentState.STA);
            lease.Flush(); worker.Thread.Start();
            try
            {
                CloseAccessObjectsWithoutSaving();
                ((dynamic)application).CloseCurrentDatabase();
                StopAccessSaveDialogHandler();
                lease.Record["CloseReturned"] = true; lease.Flush();
            }
            catch
            {
                worker.StopRequested = true;
                // Retain the generation on any failure; Close and Cancel are never retried.
                RetainUncertainOffice();
                throw;
            }
        }

        /// <summary>Immutable native readback used for both initial validation and immediate pre-click revalidation.</summary>
        internal sealed class AccessDiscardDialog
        {
            internal IntPtr Window, Edit, Cancel;
            internal int ProcessId, EditProcessId, CancelProcessId, EditId, CancelId;
            internal uint ThreadId, EditThreadId, CancelThreadId;
            internal string WindowClass, EditClass, CancelClass, Title, Name, CancelText;
            internal bool Visible, Enabled;
        }

        internal sealed class AccessDiscardLease
        {
            private readonly Func<bool> authority;
            private readonly Action<IDictionary<string, object>> flush;
            private readonly HashSet<string> names;
            private readonly HashSet<IntPtr> invoked = new HashSet<IntPtr>();
            private int workerThread;
            private bool failed;
            internal Func<bool> Stopping = () => false;
            internal Func<IntPtr, bool> VisibleOwnedDialog = window => false;
            internal readonly IDictionary<string, object> Record;
            internal readonly int ProcessId;
            internal AccessDiscardLease(int processId, string started, string handle, string path, string project, int ownerThread,
                IEnumerable<string> createdNames, Func<bool> authority, Action<IDictionary<string, object>> flush)
            {
                ProcessId = processId; this.authority = authority; this.flush = flush;
                names = new HashSet<string>(createdNames, StringComparer.Ordinal);
                Record = new Dictionary<string, object> {
                    ["ProcessId"] = processId, ["ProcessStartedUtc"] = started, ["OriginalProcessHandle"] = handle,
                    ["DocumentPath"] = path, ["Project"] = project, ["OwnerThread"] = ownerThread,
                    ["Mode"] = "DISCARD_ONLY", ["CreatedNames"] = new List<string>(names),
                    ["CancelEntries"] = 0, ["SaveEntries"] = 0, ["CloseReturned"] = false,
                    ["MacroReplay"] = false, ["QuitReplay"] = false, ["CancelReplay"] = false };
            }
            internal void BindWorkerThread()
            {
                if (workerThread != 0) throw new InvalidOperationException("The Access discard worker cannot change generation or thread.");
                workerThread = Thread.CurrentThread.ManagedThreadId; Record["WorkerThread"] = workerThread;
            }
            internal void Flush() { flush(Record); }
            internal void RequireDismissed()
            {
                foreach (var window in invoked)
                    Assert.IsFalse(VisibleOwnedDialog(window), "The owned Access naming dialog remained after Cancel; no Close/Quit replay is permitted.");
            }
            private void RequireAuthority()
            {
                if (failed || workerThread != Thread.CurrentThread.ManagedThreadId || !authority())
                    throw new InvalidOperationException("The original settled Access discard authority changed; no Cancel or retry is permitted.");
            }
            internal bool Matches(AccessDiscardDialog value)
            {
                return value != null && value.Window != IntPtr.Zero && value.Edit != IntPtr.Zero && value.Cancel != IntPtr.Zero &&
                    value.ProcessId == ProcessId && value.EditProcessId == ProcessId && value.CancelProcessId == ProcessId &&
                    value.ThreadId != 0 && value.EditThreadId == value.ThreadId && value.CancelThreadId == value.ThreadId &&
                    value.WindowClass == "#32770" && value.EditClass == "RichEdit20W" && value.CancelClass == "Button" &&
                    value.EditId == 2020 && value.CancelId == 2 && value.Visible && value.Enabled &&
                    (value.Title == "Enregistrer sous" || value.Title == "Save As") &&
                    (value.CancelText == "Annuler" || value.CancelText == "Cancel") && names.Contains(value.Name);
            }
            internal void TryCancel(IntPtr window, Func<IntPtr, AccessDiscardDialog> read, Func<IntPtr, bool> click)
            {
                if (Stopping()) return;
                RequireAuthority();
                if (invoked.Contains(window)) return;
                var before = read(window);
                if (!Matches(before) || before.Window != window) return;
                var after = read(window);
                if (Stopping()) return;
                RequireAuthority();
                if (!Matches(after) || after.Window != window || after.Edit != before.Edit || after.Cancel != before.Cancel ||
                    after.Name != before.Name || after.Title != before.Title || after.ThreadId != before.ThreadId) return;
                invoked.Add(window); // Record before delivery: an unknown BM_CLICK is never repeated.
                Record["CancelEntries"] = (int)Record["CancelEntries"] + 1;
                Record["DialogHandle"] = window.ToInt64(); Record["EditHandle"] = after.Edit.ToInt64();
                Record["CancelHandle"] = after.Cancel.ToInt64(); Record["Module"] = after.Name;
                Record["DialogThread"] = after.ThreadId; Record["EditClass"] = after.EditClass; Record["GuardState"] = "ORIGINAL_SETTLED_OWNED_TEARDOWN";
                Record["CancelOutcome"] = "ENTERED"; Record["CancelEnteredUtc"] = DateTime.UtcNow.ToString("o"); Flush();
                bool returned;
                try { RequireAuthority(); returned = click(after.Cancel); }
                catch (Exception error)
                {
                    failed = true;
                    Record["CancelOutcome"] = "CALL_FAILED_EFFECT_UNKNOWN"; Record["Error"] = error.ToString(); Flush(); throw;
                }
                Record["CancelOutcome"] = returned ? "RETURNED" : "CALL_FAILED_EFFECT_UNKNOWN"; Flush();
                if (!returned) { failed = true; throw new InvalidOperationException("The owned Access Cancel outcome is uncertain; no native cleanup replay is permitted."); }
            }
        }

        private static bool NativeVisibleOwnedAccessDiscard(IntPtr window, Process process, string started)
        {
            if (process.HasExited || process.StartTime.ToUniversalTime().ToString("o") != started) return false;
            uint pid; GetWindowThreadProcessId(window, out pid);
            return pid == (uint)process.Id && NativeAccessDiscardClass(window) == "#32770" && IsWindowVisible(window);
        }
        private static void ScanNativeAccessDiscard(AccessDiscardLease lease)
        {
            var scan = Stopwatch.StartNew();
            EnumWindows((window, parameter) => {
                if (scan.ElapsedMilliseconds >= 1000) return false;
                uint processId; GetWindowThreadProcessId(window, out processId);
                if (processId == (uint)lease.ProcessId && IsWindowVisible(window))
                    lease.TryCancel(window, ReadNativeAccessDiscardDialog, cancel => {
                        UIntPtr result;
                        return SendDialogCommand(cancel, 0x00F5, UIntPtr.Zero, IntPtr.Zero, 0x23, 750, out result) != IntPtr.Zero;
                    });
                return true;
            }, IntPtr.Zero);
        }

        private static AccessDiscardDialog ReadNativeAccessDiscardDialog(IntPtr window)
        {
            IntPtr edit = GetDlgItem(window, 2020), cancel = GetDlgItem(window, 2);
            uint pid, editPid, cancelPid;
            uint thread = GetWindowThreadProcessId(window, out pid), editThread = GetWindowThreadProcessId(edit, out editPid),
                cancelThread = GetWindowThreadProcessId(cancel, out cancelPid);
            return new AccessDiscardDialog {
                Window = window, Edit = edit, Cancel = cancel, ProcessId = (int)pid, EditProcessId = (int)editPid, CancelProcessId = (int)cancelPid,
                ThreadId = thread, EditThreadId = editThread, CancelThreadId = cancelThread,
                WindowClass = NativeAccessDiscardClass(window), EditClass = NativeAccessDiscardClass(edit), CancelClass = NativeAccessDiscardClass(cancel),
                EditId = GetDlgCtrlID(edit), CancelId = GetDlgCtrlID(cancel), Visible = IsWindowVisible(window), Enabled = IsWindowEnabled(cancel),
                Title = BoundedDialogText(window), Name = ReadNativeAccessDiscardName(edit), CancelText = BoundedDialogText(cancel) };
        }
        /// <summary>Reads the attested Access Unicode rich edit naming control with bounded WM_GETTEXT.</summary>
        internal static string ReadNativeAccessDiscardName(IntPtr window)
        {
            if (NativeAccessDiscardClass(window) != "RichEdit20W") return null;
            return BoundedDialogText(window);
        }
        private static string NativeAccessDiscardClass(IntPtr window)
        {
            var value = new StringBuilder(64); GetClassName(window, value, value.Capacity); return value.ToString();
        }
    }
}