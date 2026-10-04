using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;

namespace VBAi.Tests.Integration
{
    /// <summary>Captures the owned Access save context without activating panes, clicking dialogs or changing native state.</summary>
    internal sealed partial class OfficeVbeFixture
    {
        private const int AccessAdapterNativeSnapshotBoundMilliseconds = 5000;
        private const uint AccessAdapterTextReadBoundMilliseconds = 150;
        private readonly HashSet<string> accessAdapterDiagnosticClaims = new HashSet<string>(StringComparer.Ordinal);

        private void RecordAccessAdapterDiagnostic(string phase, IDictionary<string, object> observation)
        {
            if (Kind != "Access" || privateDesktop == null ||
                (phase != "BeforeAdapter" && phase != "ImmediatelyAfterAdapter")) return;
            string claim = ProcessId + ":" + phase;
            if (!accessAdapterDiagnosticClaims.Add(claim))
                throw new InvalidOperationException("The exact Access adapter diagnostic phase is already claimed; no read replay is permitted.");
            var diagnostic = new Dictionary<string, object> {
                ["Phase"] = phase, ["State"] = "PENDING", ["ProcessId"] = ProcessId,
                ["Desktop"] = privateDesktop, ["DocumentPath"] = DocumentPath, ["Project"] = Project,
                ["ReadOnly"] = true, ["NativeMutations"] = 0, ["ActivationAttempts"] = 0,
                ["NativeSnapshotBoundMilliseconds"] = AccessAdapterNativeSnapshotBoundMilliseconds,
                ["TextReadBoundMilliseconds"] = AccessAdapterTextReadBoundMilliseconds,
                ["BridgeResponseBoundMilliseconds"] = 5000, ["BridgeConnectionAttempts"] = 1,
                ["ReadReplayAllowed"] = false, ["OriginalAdapterResponseChanged"] = false,
                ["Utc"] = DateTime.UtcNow.ToString("o")
            };
            observation["AccessAdapterDiagnostic"] = diagnostic;
            FlushAdapterEvidence();
            try
            {
                RequireAccessAdapterDiagnosticHost();
                diagnostic["PrivateAndInputInventoriesVerified"] = true;
                diagnostic["InputDesktop"] = IsolatedTestDesktop.InputDesktopName();
                diagnostic["SentinelWindow"] = System.Environment.GetEnvironmentVariable("VBAi_TEST_DESKTOP_SENTINEL_HWND");
                diagnostic["SentinelProcessId"] = System.Environment.GetEnvironmentVariable("VBAi_TEST_DESKTOP_SENTINEL_PID");
                diagnostic["SentinelThreadId"] = System.Environment.GetEnvironmentVariable("VBAi_TEST_DESKTOP_SENTINEL_TID");
                diagnostic["ProcessStartedUtc"] = shutdownEvidence.Record["ProcessStartedUtc"];
                diagnostic["ProcessImage"] = shutdownEvidence.Record["ProcessImage"];
                diagnostic["OriginalProcessHandle"] = shutdownEvidence.Record["OriginalProcessHandle"];
                diagnostic["OriginalNativeHandle"] = privateDesktopChild.ProcessHandle.ToInt64();
                diagnostic["ExpectedAssemblyMvid"] = typeof(VbeSession).Module.ModuleVersionId.ToString("D");
                diagnostic["NativeWindows"] = ReadAccessAdapterNativeWindows(diagnostic);
                diagnostic["NativeInventoryCompleted"] = true;
                FlushAdapterEvidence();
                if (((object[])diagnostic["NativeWindows"]).Cast<IDictionary<string, object>>().Any(row =>
                    Equals(row["Class"], "#32770") && Equals(row["Visible"], true)))
                    throw new InvalidOperationException("An owned visible Access dialog is preserved in the snapshot; no bridge read or Save behind an unknown modal is permitted.");

                // Keep the fixture's containment/evidence path, but bound these read-only requests
                // separately from the existing mutation transport and never repeat an emitted read.
                var originalDispatch = Dispatch;
                try
                {
                    Dispatch = (pid, request) => VbeBridgeClient.Read("VBAi." + pid, request, 5000, 250, 1, 0);
                    RequireAccessAdapterDiagnosticHost();
                    var status = Data("status");
                    diagnostic["BridgeStatus"] = status;
                    FlushAdapterEvidence();
                    if (!Equals(status["Connected"], true) || Convert.ToInt32(status["HostProcessId"]) != ProcessId ||
                        !Equals(status["AssemblyModuleVersionId"], typeof(VbeSession).Module.ModuleVersionId.ToString("D")))
                        throw new InvalidOperationException("Access diagnostic bridge PID/connection/MVID changed; no native save is authorized.");
                    RequireAccessAdapterDiagnosticHost();
                    var persistence = Data("project_persistence_status");
                    diagnostic["Persistence"] = persistence;
                    FlushAdapterEvidence();
                    if (!Equals(persistence["Host"], "Access") || !Equals(persistence["IdentityVerified"], true) ||
                        Convert.ToInt32(persistence["OwnerProcessId"]) != ProcessId ||
                        !string.Equals(Convert.ToString(persistence["HostPath"]), DocumentPath, StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException("Access diagnostic project/path identity is unavailable; no native save is authorized.");
                    RequireAccessAdapterDiagnosticHost();
                    diagnostic["SaveCommands"] = Items("list_commands", "Query", "Enregistrer", "Limit", 200);
                    FlushAdapterEvidence();
                    RequireAccessAdapterDiagnosticHost();
                    // vbe_windows enumerates existing windows and ActiveWindow. It does not
                    // obtain CodeModule.CodePane, whose getter opens/activates a native pane.
                    diagnostic["VbeWindows"] = Data("vbe_windows");
                    FlushAdapterEvidence();
                    RequireAccessAdapterDiagnosticHost();
                }
                finally { Dispatch = originalDispatch; }
                diagnostic["State"] = "READONLY_COMPLETED";
            }
            catch (Exception error)
            {
                diagnostic["State"] = "REFUSED_OR_UNAVAILABLE";
                diagnostic["Error"] = error.ToString();
                diagnostic["HResult"] = "0x" + unchecked((uint)error.HResult).ToString("X8");
                // The original response is already durable for the post-save phase. Failure
                // never promotes that response or permits Save/Compile/Close/Quit replay.
                throw;
            }
            finally
            {
                diagnostic["FinishedUtc"] = DateTime.UtcNow.ToString("o");
                FlushAdapterEvidence();
            }
        }

        private void RequireAccessAdapterDiagnosticHost(IntPtr window = default(IntPtr))
        {
            commandContainment.RequireTerminal();
            if (Kind != "Access" || privateDesktop == null || !owned || NativeExecutionUnsettled ||
                ownedProcess == null || shutdownEvidence == null || privateDesktopChild == null ||
                privateDesktopChild.ProcessId != ProcessId ||
                Thread.CurrentThread.ManagedThreadId != shutdownOwnerThread || ownedProcess.HasExited ||
                ownedProcess.Id != ProcessId ||
                !Equals(shutdownEvidence.Record["ProcessId"], ProcessId) ||
                !Equals(shutdownEvidence.Record["ProcessStartedUtc"], ownedProcess.StartTime.ToUniversalTime().ToString("o")) ||
                !Equals(shutdownEvidence.Record["OriginalProcessHandle"], "0x" + unchecked((ulong)ownedProcess.Handle.ToInt64()).ToString("X16")) ||
                !string.Equals(Convert.ToString(shutdownEvidence.Record["ProcessImage"]), ExcelOwnedProcessImage.Read(ownedProcess.Handle), StringComparison.OrdinalIgnoreCase) ||
                !Equals(shutdownEvidence.Record["QuitEntries"], 0) || Equals(shutdownEvidence.Record["TeardownPrepared"], true))
                throw new InvalidOperationException("Access adapter diagnostics require the unchanged original owned process before cleanup.");
            RequirePrivateHostDesktop(true); // Original native handle/image, current private desktop, input inventory and exact sentinel.
            if (window != IntPtr.Zero)
                IsolatedTestDesktop.RequireOfficeWindowInventory(privateDesktop, (uint)ProcessId, true, window);
        }

        private object[] ReadAccessAdapterNativeWindows(IDictionary<string, object> diagnostic)
        {
            var clock = Stopwatch.StartNew();
            var windows = new List<IntPtr>();
            bool withinBound = true;
            // RequireCurrent above binds EnumWindows to this exact inactive desktop.
            bool complete = EnumWindows((window, state) => {
                uint pid;
                uint thread = GetWindowThreadProcessId(window, out pid);
                if (pid != (uint)ProcessId) return true; // Never read another process's class or UI text.
                if (thread == 0 || windows.Count >= 128 || clock.ElapsedMilliseconds > AccessAdapterNativeSnapshotBoundMilliseconds)
                { withinBound = false; return false; }
                windows.Add(window);
                return true;
            }, IntPtr.Zero);
            if (!complete || !withinBound)
                throw new InvalidOperationException("The exact owned Access top-level window inventory is incomplete or exceeds its fixed bound.");
            var rows = new List<object>();
            diagnostic["NativeWindows"] = rows;
            foreach (var window in windows)
            {
                RequireAccessAdapterNativeTime(clock);
                var row = ReadAccessAdapterNativeWindow(window, window, diagnostic, clock, false);
                rows.Add(row);
                FlushAdapterEvidence();
                if (!Equals(row["Class"], "#32770")) continue;
                var children = new List<IntPtr>();
                bool childBound = true;
                EnumChildWindows(window, (child, state) => {
                    uint pid;
                    uint thread = GetWindowThreadProcessId(child, out pid);
                    if (pid != (uint)ProcessId) return true;
                    if (thread == 0 || GetAncestor(child, 2) != window || children.Count >= 64 ||
                        clock.ElapsedMilliseconds > AccessAdapterNativeSnapshotBoundMilliseconds)
                    { childBound = false; return false; }
                    children.Add(child);
                    return true;
                }, IntPtr.Zero);
                if (!childBound)
                    throw new InvalidOperationException("The exact owned Access dialog-child inventory is incomplete or exceeds its fixed bound.");
                var controls = new List<object>();
                row["DialogChildren"] = controls;
                foreach (var child in children)
                {
                    RequireAccessAdapterNativeTime(clock);
                    controls.Add(ReadAccessAdapterNativeWindow(child, window, diagnostic, clock, true));
                    FlushAdapterEvidence();
                }
            }
            RequireAccessAdapterNativeTime(clock);
            RequireAccessAdapterDiagnosticHost();
            diagnostic["NativeSnapshotElapsedMilliseconds"] = clock.ElapsedMilliseconds;
            return rows.ToArray();
        }

        private static void RequireAccessAdapterNativeTime(Stopwatch clock)
        {
            if (clock.ElapsedMilliseconds > AccessAdapterNativeSnapshotBoundMilliseconds)
                throw new InvalidOperationException("The read-only Access native snapshot exceeded its fixed total bound; no read is repeated.");
        }

        private IDictionary<string, object> ReadAccessAdapterNativeWindow(IntPtr window, IntPtr expectedRoot,
            IDictionary<string, object> diagnostic, Stopwatch clock, bool child)
        {
            RequireAccessAdapterDiagnosticHost(window);
            uint pid, rootPid;
            uint thread = GetWindowThreadProcessId(window, out pid);
            uint rootThread = GetWindowThreadProcessId(expectedRoot, out rootPid);
            if (thread == 0 || rootThread == 0 || pid != (uint)ProcessId || rootPid != (uint)ProcessId ||
                GetAncestor(window, 2) != expectedRoot)
                throw new InvalidOperationException("The exact owned Access native window ancestry changed before its read.");
            var className = new StringBuilder(256);
            int length = GetClassName(window, className, className.Capacity);
            if (length < 1 || length >= className.Capacity - 1)
                throw new InvalidOperationException("The owned Access native window class is unavailable or truncated.");
            string kind = className.ToString();
            bool visible = IsWindowVisible(window);
            var row = new Dictionary<string, object> {
                ["Window"] = window.ToInt64(), ["RootWindow"] = expectedRoot.ToInt64(), ["ProcessId"] = pid,
                ["ThreadId"] = thread, ["RootThreadId"] = rootThread, ["Class"] = kind,
                ["Visible"] = visible, ["Enabled"] = IsWindowEnabled(window), ["Child"] = child,
                ["ControlId"] = child ? GetDlgCtrlID(window) : 0
            };
            // Native dialog captions, Static labels and Button labels disambiguate Save
            // prompts. Editor/input contents and hidden non-dialog windows are not read.
            bool textRequired = child ? kind == "Static" || kind == "Button" : visible || kind == "#32770";
            if (!textRequired) { row["TextRead"] = "OMITTED"; return row; }
            RequireAccessAdapterNativeTime(clock);
            row["TextRead"] = "PENDING";
            diagnostic["PendingNativeTextRead"] = row;
            FlushAdapterEvidence();
            RequireAccessAdapterDiagnosticHost(window);
            if (GetWindowThreadProcessId(window, out pid) != thread || pid != (uint)ProcessId || GetAncestor(window, 2) != expectedRoot)
                throw new InvalidOperationException("The exact Access text target changed before its sole getter.");
            var text = new StringBuilder(4096);
            UIntPtr returned;
            bool read = ReadDialogText(window, 0x000D, new UIntPtr(4096), text, 0x23, AccessAdapterTextReadBoundMilliseconds, out returned) != IntPtr.Zero;
            row["TextRead"] = read ? "RETURNED" : "UNAVAILABLE";
            row["ReturnedCharacters"] = returned.ToUInt64();
            row["TextReadAttempts"] = 1;
            if (read && returned.ToUInt64() < 4095) row["Text"] = text.ToString();
            FlushAdapterEvidence();
            if (!read || returned.ToUInt64() >= 4095)
                throw new InvalidOperationException("Owned Access native text is unavailable or truncated; no getter retry or native save is authorized.");
            RequireAccessAdapterDiagnosticHost(window);
            if (GetWindowThreadProcessId(window, out pid) != thread || pid != (uint)ProcessId || GetAncestor(window, 2) != expectedRoot)
                throw new InvalidOperationException("The Access native text target changed during its sole getter.");
            RequireAccessAdapterNativeTime(clock);
            diagnostic.Remove("PendingNativeTextRead");
            return row;
        }
    }
}
