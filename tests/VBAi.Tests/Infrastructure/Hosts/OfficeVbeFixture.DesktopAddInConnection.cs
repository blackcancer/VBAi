using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace VBAi.Tests.Integration
{
    internal sealed partial class OfficeVbeFixture
    {
        private const string FixtureAddInProgId = "VBAi.AddIn";
        private readonly int addInFixtureThread = Thread.CurrentThread.ManagedThreadId;
        private object addInConnectionGeneration;
        private PrivateAccessAddInConnectionGate addInConnectionGate;

        /// <summary>Explicitly connects only the owned private Access VBE entry; bridge PID/MVID validation still follows.</summary>
        private void ConnectPrivateAccessAddIn(object editor, object window)
        {
            if (privateDesktop == null) return;
            if (Kind != "Access" || !owned || privateDesktopChild == null || ownedProcess == null)
                throw new InvalidOperationException("Add-in connection requires the original owned private Access generation.");
            if (!ReferenceEquals(addInConnectionGeneration, privateDesktopChild))
            {
                addInConnectionGeneration = privateDesktopChild;
                addInConnectionGate = new PrivateAccessAddInConnectionGate();
            }
            var child = privateDesktopChild;
            var app = application;
            var process = ownedProcess;
            int pid = ProcessId;
            object addins = null, entry = null;
            IntPtr originalWindow = IntPtr.Zero;
            uint originalWindowThread = 0;
            string appIdentity = null, editorIdentity = null, entryIdentity = null;
            Action requireOwner = () =>
            {
                RequirePrivateAccessAddInFixtureThread(addInFixtureThread, Thread.CurrentThread.ManagedThreadId, Thread.CurrentThread.GetApartmentState());
                if (Kind != "Access" || !owned || ProcessId != pid || !ReferenceEquals(child, privateDesktopChild) ||
                    !ReferenceEquals(app, application) || !ReferenceEquals(process, ownedProcess))
                    throw new InvalidOperationException("The Access add-in connection no longer targets its original owned application/process.");
                commandContainment.RequireTerminal();
                if (NativeExecutionUnsettled) throw new InvalidOperationException("An unsettled command refuses add-in connection.");
                RequireApplicationOwner();
                RequirePrivateHostDesktop(true);
                IntPtr currentWindow = new IntPtr(Convert.ToInt64(((dynamic)window).HWnd));
                uint windowPid;
                uint windowThread = GetWindowThreadProcessId(currentWindow, out windowPid);
                RequirePrivateAccessAddInWindow(pid, originalWindow, originalWindowThread, currentWindow, windowPid, windowThread, IsWindowEnabled(currentWindow));
                IsolatedTestDesktop.RequireOfficeWindowInventory(privateDesktop, (uint)pid, true, currentWindow);
                if (entry != null) RequirePrivateAccessAddInEntry((string)((dynamic)entry).ProgId, (string)((dynamic)entry).Guid);
                if (appIdentity != null && (appIdentity != PrivateAccessConnectionIdentity(app) ||
                    editorIdentity != PrivateAccessConnectionIdentity(editor) || (entryIdentity != null && entryIdentity != PrivateAccessConnectionIdentity(entry))))
                    throw new InvalidOperationException("The canonical original Access application/VBE/AddIn identity changed.");
            };
            Action<IDictionary<string, object>> record = row =>
            {
                row["ProcessId"] = pid; row["Desktop"] = privateDesktop;
                row["OriginalHandle"] = child.ProcessHandle.ToInt64();
                row["StartedUtc"] = process.StartTime.ToUniversalTime().ToString("o");
                row["Image"] = privateExecutable;
                row["ExpectedMvid"] = typeof(VbeSession).Module.ModuleVersionId.ToString("D");
                row["FixtureManagedThreadId"] = addInFixtureThread;
                row["VbeWindow"] = originalWindow.ToInt64(); row["VbeNativeThreadId"] = originalWindowThread;
                row["ApplicationIUnknown"] = appIdentity; row["VbeIUnknown"] = editorIdentity; row["AddInIUnknown"] = entryIdentity;
                steps.Add(row); FlushAdapterEvidence();
            };
            try
            {
                requireOwner();
                originalWindow = new IntPtr(Convert.ToInt64(((dynamic)window).HWnd));
                uint owner;
                originalWindowThread = GetWindowThreadProcessId(originalWindow, out owner);
                RequirePrivateAccessAddInWindow(pid, originalWindow, originalWindowThread, originalWindow, owner, originalWindowThread);
                requireOwner();
                appIdentity = PrivateAccessConnectionIdentity(app); editorIdentity = PrivateAccessConnectionIdentity(editor);
                addins = ((dynamic)editor).AddIns;
                entry = ((dynamic)addins).Item(FixtureAddInProgId);
                entryIdentity = PrivateAccessConnectionIdentity(entry);
                var beforeWindows = ReadPrivateAccessConnectionWindows(requireOwner);
                steps.Add(new { PrivateAccessConnectionWindows = "BeforeConnect", ProcessId = pid, Windows = beforeWindows, Complete = true }); FlushAdapterEvidence();
                RequireNoPrivateAccessConnectionModal(beforeWindows, IsWindowEnabled(originalWindow));
                addInConnectionGate.Run(FixtureAddInProgId, requireOwner, () => ((dynamic)entry).Connect,
                    () => ((dynamic)entry).Connect = true, ReadPrivateAccessAddInLoadBehavior, record);
                var afterWindows = ReadPrivateAccessConnectionWindows(requireOwner);
                steps.Add(new { PrivateAccessConnectionWindows = "AfterConnectBeforeBridge", ProcessId = pid, Windows = afterWindows, Complete = true }); FlushAdapterEvidence();
                RequireNoPrivateAccessConnectionModal(afterWindows, IsWindowEnabled(originalWindow));
                requireOwner();
            }
            catch
            {
                // A failed getter, setter, readback or evidence delivery never authorizes another connection.
                RetainUncertainOffice();
                throw;
            }
            finally { Release(entry); Release(addins); }
        }

        private static string PrivateAccessConnectionIdentity(object target)
        {
            IntPtr identity = Marshal.GetIUnknownForObject(target);
            try { return "0x" + unchecked((ulong)identity.ToInt64()).ToString("X16"); }
            finally { Marshal.Release(identity); }
        }

        internal sealed class PrivateAccessConnectionWindow
        {
            public long Window { get; set; }
            public uint ThreadId { get; set; }
            public string Class { get; set; }
            public bool Visible { get; set; }
            public bool Enabled { get; set; }
        }

        private PrivateAccessConnectionWindow[] ReadPrivateAccessConnectionWindows(Action requireOwner)
        {
            requireOwner();
            var rows = new List<PrivateAccessConnectionWindow>();
            var clock = Stopwatch.StartNew();
            bool complete = EnumWindows((window, parameter) =>
            {
                if (clock.ElapsedMilliseconds > 5000) return false;
                uint owner; uint thread = GetWindowThreadProcessId(window, out owner);
                if (owner != (uint)ProcessId) return true;
                if (thread == 0 || rows.Count >= 128) return false;
                var text = new StringBuilder(256);
                int count = GetClassName(window, text, text.Capacity);
                if (count <= 0 || count >= text.Capacity - 1) return false;
                rows.Add(new PrivateAccessConnectionWindow
                {
                    Window = window.ToInt64(),
                    ThreadId = thread,
                    Class = text.ToString(),
                    Visible = IsWindowVisible(window),
                    Enabled = IsWindowEnabled(window)
                });
                return true;
            }, IntPtr.Zero);
            if (!complete || clock.ElapsedMilliseconds > 5000 || rows.Count == 0)
                throw new InvalidOperationException("The complete owned private Access connection window inventory is unavailable or exceeds its bound.");
            requireOwner(); // Fresh complete private/input/sentinel proofs also follow any UI created by OnConnection.
            return rows.ToArray();
        }

        internal static void RequireNoPrivateAccessConnectionModal(IEnumerable<PrivateAccessConnectionWindow> windows, bool vbeEnabled)
        {
            if (!vbeEnabled || windows == null) throw new InvalidOperationException("The Access VBE is disabled or its connection window inventory is incomplete.");
            foreach (var window in windows)
                if (window == null || window.Window == 0 || window.ThreadId == 0 || string.IsNullOrEmpty(window.Class) ||
                    window.Visible && (window.Class == "#32770" || window.Class.StartsWith("bosa_sdm", StringComparison.Ordinal)))
                    throw new InvalidOperationException("An unknown visible owned Access dialog-class window or incomplete identity refuses subsequent bridge/scenario operations.");
        }

        internal static void RequirePrivateAccessAddInFixtureThread(int expected, int current, ApartmentState apartment)
        {
            if (expected != current || apartment != ApartmentState.STA)
                throw new InvalidOperationException("Access add-in connection must stay on the original fixture STA.");
        }

        internal static void RequirePrivateAccessAddInEntry(string progId, string clsid)
        {
            if (!string.Equals(progId, FixtureAddInProgId, StringComparison.OrdinalIgnoreCase) ||
                !Guid.TryParse(clsid, out Guid identity) || identity != typeof(AddIn).GUID)
                throw new InvalidOperationException("The native VBE entry is not the exact VBAi ProgID/CLSID.");
        }

        internal static void RequirePrivateAccessAddInWindow(int pid, IntPtr expectedWindow, uint expectedThread,
            IntPtr actualWindow, uint actualPid, uint actualThread, bool windowEnabled = true)
        {
            if (!windowEnabled || pid <= 0 || actualWindow == IntPtr.Zero || actualPid != (uint)pid || actualThread == 0 ||
                (expectedWindow != IntPtr.Zero && (expectedWindow != actualWindow || expectedThread != actualThread)))
                throw new InvalidOperationException("The Access VBE HWND/PID/native UI thread ownership changed.");
        }

        private static string ReadPrivateAccessAddInLoadBehavior()
        {
            var rows = new List<string>();
            foreach (RegistryHive hive in new[] { RegistryHive.CurrentUser, RegistryHive.LocalMachine })
                using (var root = RegistryKey.OpenBaseKey(hive, RegistryView.Registry64))
                using (var key = root.OpenSubKey(@"Software\Microsoft\VBA\VBE\6.0\Addins64\VBAi.AddIn", false))
                {
                    object value = key?.GetValue("LoadBehavior", null, RegistryValueOptions.DoNotExpandEnvironmentNames);
                    rows.Add(hive + "|" + (key == null ? "KEY_ABSENT" : value == null ? "VALUE_ABSENT" :
                        key.GetValueKind("LoadBehavior") + "|" + Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture)));
                }
            return string.Join(";", rows);
        }

        /// <summary>Single-use fixture connection; evidence and ownership must authorize the sole mutation.</summary>
        internal sealed class PrivateAccessAddInConnectionGate
        {
            internal bool Consumed { get; private set; }
            internal int SetterEntries { get; private set; }
            internal bool DeliveryUncertain { get; private set; }

            internal void Run(string progId, Action requireOwner, Func<object> readConnect, Action connectOnce,
                Func<string> readRegistration, Action<IDictionary<string, object>> record)
            {
                if (Consumed) throw new InvalidOperationException("The original Access add-in connection observation was already consumed; no retry is permitted.");
                Consumed = true;
                if (!string.Equals(progId, FixtureAddInProgId, StringComparison.Ordinal))
                    throw new InvalidOperationException("Only the exact VBAi.AddIn VBE entry may be connected.");
                Action<string, object> receipt = (state, connect) => record(new Dictionary<string, object>
                {
                    ["PrivateAccessAddInConnection"] = true,
                    ["ProgId"] = progId,
                    ["State"] = state,
                    ["Connect"] = connect,
                    ["SetterEntries"] = SetterEntries,
                    ["DeliveryUncertain"] = DeliveryUncertain,
                    ["AutomaticRetry"] = false,
                    ["Utc"] = DateTime.UtcNow.ToString("o")
                });
                try
                {
                    receipt("OBSERVATION_PENDING", null);
                    requireOwner();
                    string registration = readRegistration();
                    if (string.IsNullOrEmpty(registration)) throw new InvalidOperationException("Add-in registration observation is incomplete.");
                    record(new Dictionary<string, object>
                    {
                        ["PrivateAccessAddInLoadBehavior"] = "BeforeConnect",
                        ["Snapshot"] = registration,
                        ["RegistryWrites"] = 0
                    });
                    object before = readConnect();
                    if (!(before is bool connected)) throw new InvalidOperationException("The native AddIn.Connect getter did not return an exact Boolean.");
                    receipt("OBSERVED", connected);
                    if (connected)
                    {
                        requireOwner();
                        string noopRegistration = readRegistration();
                        record(new Dictionary<string, object>
                        {
                            ["PrivateAccessAddInLoadBehavior"] = "AfterNoop",
                            ["Snapshot"] = noopRegistration,
                            ["RegistryWrites"] = 0
                        });
                        if (registration != noopRegistration) throw new InvalidOperationException("Add-in LoadBehavior changed during the connection observation.");
                        receipt("NOOP_ALREADY_CONNECTED", true);
                        return;
                    }
                    receipt("SETTER_CLAIMED", false); // Must be durable before entering the native setter.
                    requireOwner();
                    if (registration != readRegistration()) throw new InvalidOperationException("Add-in LoadBehavior changed before the connection setter.");
                    requireOwner(); // Final original application/window/desktop/STA/ProgID authorization.
                    SetterEntries = 1; DeliveryUncertain = true;
                    connectOnce();
                    receipt("SETTER_RETURNED", null);
                    requireOwner();
                    object after = readConnect();
                    if (!(after is bool verified) || !verified)
                        throw new InvalidOperationException("The one Access connection setter was not verified by an exact true Boolean readback.");
                    string afterRegistration = readRegistration();
                    record(new Dictionary<string, object>
                    {
                        ["PrivateAccessAddInLoadBehavior"] = "AfterConnect",
                        ["Snapshot"] = afterRegistration,
                        ["RegistryWrites"] = 0
                    });
                    if (registration != afterRegistration) throw new InvalidOperationException("Add-in LoadBehavior unexpectedly changed after connection; no registry repair is permitted.");
                    requireOwner();
                    DeliveryUncertain = false;
                    receipt("CONNECTED_VERIFIED", true);
                }
                catch (Exception original)
                {
                    DeliveryUncertain = SetterEntries != 0;
                    try { receipt("FAILED_NO_RETRY", null); }
                    catch (Exception evidence) { throw new AggregateException("Add-in connection and durable evidence failed; retain the original host without retry.", original, evidence); }
                    throw;
                }
            }
        }
    }
}
