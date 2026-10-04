using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Win32;

namespace VBAi.Tests.Integration
{
    internal sealed partial class OfficeVbeFixture
    {
        private string privateDesktop, privateExecutable;
        private IsolatedTestDesktop.NativeChild privateDesktopChild;
        private Action showPrivateAccess;

        /// <summary>Requires matching generated desktop names before selecting a host-specific private bootstrap.</summary>
        internal static string RequireOfficeDesktopPair(string required, string configured)
        {
            if (!string.IsNullOrEmpty(required) && !string.Equals(required, configured, StringComparison.Ordinal))
                throw new InvalidOperationException("The qualification and Office desktop names must match exactly.");
            if (string.IsNullOrEmpty(configured)) return null;
            IsolatedTestDesktop.RequireName(configured);
            return configured;
        }

        /// <summary>Validates the opt-in without launching, activating or changing any native desktop.</summary>
        internal static string RequirePrivateOfficeExecutable(string kind, string desktop, string executable)
        {
            IsolatedTestDesktop.RequireName(desktop);
            if (kind != "Access" && kind != "Publisher" && kind != "PowerPoint")
                throw new InvalidOperationException("The private Office bootstrap supports Access, Publisher and PowerPoint; no COM activation fallback is permitted.");
            string expectedName = kind == "Access" ? "MSACCESS.EXE" : kind == "PowerPoint" ? "POWERPNT.EXE" : "MSPUB.EXE";
            if (string.IsNullOrWhiteSpace(executable) || executable.Length < 4 || !char.IsLetter(executable[0]) || executable[1] != ':' ||
                (executable[2] != '\\' && executable[2] != '/') || !Path.IsPathRooted(executable) ||
                !string.Equals(Path.GetFileName(executable), expectedName, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("A reviewed absolute local installed Office executable is required before a private host launch.");
            return Path.GetFullPath(executable);
        }

        /// <summary>Uses an owned normal GUI process; manual embedding mode did not initialize Publisher publication operations.</summary>
        internal static string[] PrivateOfficeArguments(string kind)
        {
            // Access supports GetObject attachment to a running instance:
            // https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/automation-with-microsoft-access
            // Explicit -Embedding did not expose a ROT object in the frozen private-desktop campaign.
            if (kind == "Access") return new string[0];
            if (kind == "Publisher") return new string[0];
            if (kind == "PowerPoint") return new string[0];
            throw new InvalidOperationException("Private Office arguments require Access, Publisher or PowerPoint.");
        }

        /// <summary>Separates the observed registered server command from the selected normal-GUI launch.</summary>
        internal static string RegisteredPublisherServerArguments(string executable, string commandLine)
        {
            if (string.IsNullOrWhiteSpace(executable) || string.IsNullOrWhiteSpace(commandLine) || commandLine.Length > 8192 ||
                commandLine.IndexOf('\0') >= 0 || commandLine.IndexOf('\r') >= 0 || commandLine.IndexOf('\n') >= 0)
                throw new InvalidOperationException("Publisher registered server command is absent or malformed; no guessed arguments are recorded.");
            string observed = commandLine.Trim();
            foreach (string prefix in new[] { IsolatedTestDesktop.Quote(executable), executable })
                if (observed.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) &&
                    (observed.Length == prefix.Length || char.IsWhiteSpace(observed[prefix.Length])))
                    return observed.Substring(prefix.Length).Trim();
            throw new InvalidOperationException("Publisher registered server image differs from the reviewed installed image; no arguments are inferred.");
        }

        private void RecordPrivatePublisherLaunchMode(string executable, string[] selectedArguments)
        {
            if (Kind != "Publisher") return;
            if (selectedArguments == null || selectedArguments.Length != 0)
                throw new InvalidOperationException("This frozen Publisher hypothesis requires exactly zero command-line arguments.");
            using (var classes = RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Registry64))
            using (var clsidKey = classes.OpenSubKey(@"Publisher.Application\CLSID", false))
            {
                string clsid = clsidKey?.GetValue(null, null, RegistryValueOptions.DoNotExpandEnvironmentNames) as string;
                Guid id;
                if (!Guid.TryParse(clsid, out id) || id != new Guid("0002123d-0000-0000-c000-000000000046"))
                    throw new InvalidOperationException("The installed Publisher.Application class identity is unverified; no launch is permitted.");
                using (var server = classes.OpenSubKey(@"CLSID\" + id.ToString("B") + @"\LocalServer32", false))
                {
                    string command = server?.GetValue(null, null, RegistryValueOptions.DoNotExpandEnvironmentNames) as string;
                    string registeredArguments = RegisteredPublisherServerArguments(executable, command);
                    steps.Add(new { PublisherPrivateLaunchMode = "NormalGUIExplicitCreateProcess",
                        ObservedRegisteredServerCommand = command, ObservedRegisteredServerArguments = registeredArguments,
                        RegistryView = "Registry64", SelectedArguments = selectedArguments,
                        HistoricalManualEmbeddingArguments = new[] { "/Automation", "-Embedding" },
                        ComActivationCalled = false, ComActivationFallbackAllowed = false, RegisteredServerCommandExecuted = false });
                    FlushAdapterEvidence();
                }
            }
        }
        /// <summary>Records one read-only getter without retrying discovery or a native mutation after failure.</summary>
        internal static bool ObservePrivateAccessAutomationGetter(string property, Func<object> read,
            Action<IDictionary<string, object>> record)
        {
            if (property != "UserControl" && property != "Visible")
                throw new ArgumentException("Only Access automation-state getters are permitted.", nameof(property));
            if (read == null || record == null) throw new ArgumentNullException(read == null ? nameof(read) : nameof(record));
            record(new Dictionary<string, object> { ["PrivateAccessAutomationGetter"] = property,
                ["State"] = "PENDING", ["MaximumReadEntries"] = 1, ["MutationInvoked"] = false,
                ["Utc"] = DateTime.UtcNow.ToString("o") });
            object value;
            try
            {
                value = read();
                if (!(value is bool)) throw new InvalidOperationException("Access automation-state getter did not return an exact Boolean.");

            }
            catch (Exception error)
            {
                try
                {
                    record(new Dictionary<string, object> { ["PrivateAccessAutomationGetter"] = property,
                        ["State"] = "FAILED", ["Error"] = error.ToString(),
                        ["HResult"] = "0x" + unchecked((uint)error.HResult).ToString("X8"),
                        ["AutomaticRetry"] = false, ["MutationInvoked"] = false, ["Utc"] = DateTime.UtcNow.ToString("o") });
                }
                catch (Exception evidence)
                {
                    throw new AggregateException("Access automation getter and durable failure evidence both failed.", error, evidence);
                }
                throw;
            }
            record(new Dictionary<string, object> { ["PrivateAccessAutomationGetter"] = property,
                ["State"] = "RETURNED", ["Value"] = value, ["ReadEntries"] = 1, ["MutationInvoked"] = false,
                ["Utc"] = DateTime.UtcNow.ToString("o") });
            return (bool)value;
        }

        /// <summary>Consumes one visibility decision; an already-visible owned host requires no setter.</summary>
        internal sealed class PrivateAccessVisibilityDecision
        {
            private readonly bool userControl, visible;
            private bool consumed;

            internal PrivateAccessVisibilityDecision(bool userControl, bool visible)
            {
                // UserControl is read-only for a manually launched application.
                // https://learn.microsoft.com/en-us/office/vba/api/access.application.visible
                if (userControl && !visible)
                    throw new InvalidOperationException("A manually controlled but hidden Access instance is unsupported; no setter or database operation is permitted.");
                this.userControl = userControl;
                this.visible = visible;
            }

            internal void Apply(Action requireOwner, Action setVisible, Action<IDictionary<string, object>> record)
            {
                if (requireOwner == null || setVisible == null || record == null)
                    throw new ArgumentNullException(requireOwner == null ? nameof(requireOwner) : setVisible == null ? nameof(setVisible) : nameof(record));
                if (consumed) throw new InvalidOperationException("The private Access visibility decision is already consumed; no replay is permitted.");
                consumed = true;
                requireOwner();
                record(new Dictionary<string, object> {
                    ["PrivateAccessVisibility"] = visible ? "ExistingOwnedVisibleState" : "SetVisibleOnce",
                    ["State"] = "PENDING", ["UserControl"] = userControl, ["ObservedVisible"] = visible,
                    ["SetterRequired"] = !visible, ["MaximumSetterEntries"] = visible ? 0 : 1,
                    ["AutomaticRetry"] = false, ["Utc"] = DateTime.UtcNow.ToString("o")
                });
                if (visible)
                {
                    record(new Dictionary<string, object> { ["PrivateAccessVisibility"] = "ExistingOwnedVisibleState",
                        ["State"] = "NO_OP", ["SetterEntries"] = 0, ["MutationInvoked"] = false,
                        ["Utc"] = DateTime.UtcNow.ToString("o") });
                    return;
                }
                try { setVisible(); }
                catch (Exception error)
                {
                    try
                    {
                        record(new Dictionary<string, object> { ["PrivateAccessVisibility"] = "SetVisibleOnce",
                            ["State"] = "FAILED", ["Error"] = error.ToString(),
                            ["HResult"] = "0x" + unchecked((uint)error.HResult).ToString("X8"),
                            ["SetterEntries"] = 1, ["MutationInvoked"] = true, ["SetterOutcomeUncertain"] = true,
                            ["AutomaticRetry"] = false, ["Utc"] = DateTime.UtcNow.ToString("o") });
                    }
                    catch (Exception evidence)
                    {
                        throw new AggregateException("Access visibility setter and durable failure evidence both failed.", error, evidence);
                    }
                    throw;
                }
                record(new Dictionary<string, object> { ["PrivateAccessVisibility"] = "SetVisibleOnce",
                    ["State"] = "RETURNED", ["SetterEntries"] = 1, ["MutationInvoked"] = true,
                    ["Utc"] = DateTime.UtcNow.ToString("o") });
            }
        }

        private void RecordPrivateAccessAutomationState()
        {
            Action<IDictionary<string, object>> record = row => {
                row["ProcessId"] = ProcessId; row["Desktop"] = privateDesktop;
                steps.Add(row); FlushAdapterEvidence();
            };
            bool userControl = ObservePrivateAccessAutomationGetter("UserControl", () => ((dynamic)application).UserControl, record);
            bool visible = ObservePrivateAccessAutomationGetter("Visible", () => ((dynamic)application).Visible, record);
            RequireApplicationOwner();
            RequirePrivateHostDesktop(true);
            steps.Add(new { PrivateAccessAutomationContract = true, ProcessId, UserControl = userControl,
                Visible = visible, VisibilityAction = visible ? "ExistingOwnedVisibleState" : userControl ? "RefuseUnsupportedState" : "SetVisibleOnce",
                NativeMutations = 0 });
            FlushAdapterEvidence();
            var decision = new PrivateAccessVisibilityDecision(userControl, visible);
            var observedChild = privateDesktopChild;
            var observedApplication = application;
            int observedPid = ProcessId;
            Action requireOwner = () => {
                if (!ReferenceEquals(privateDesktopChild, observedChild) || !ReferenceEquals(application, observedApplication) || ProcessId != observedPid)
                    throw new InvalidOperationException("The Access visibility observation belongs to another original host; no native operation is permitted.");
                RequireApplicationOwner();
                RequirePrivateHostDesktop(true);
            };
            showPrivateAccess = () => {
                decision.Apply(requireOwner, () => ((dynamic)observedApplication).Visible = true, record);
                // Revalidate the same original host immediately before the caller creates or opens its database.
                requireOwner();
            };
        }

        private void ShowAccessForFixture()
        {
            if (privateDesktop == null) { ((dynamic)application).Visible = true; return; }
            if (Kind != "Access" || showPrivateAccess == null)
                throw new InvalidOperationException("Private Access visibility has no complete owned-host observation; no database operation is permitted.");
            showPrivateAccess();
        }

        /// <summary>Explicitly launches once on the exact desktop, then discovers an existing ROT object without activation.</summary>
        private void StartPrivateOfficeHost(string desktop)
        {
            IsolatedTestDesktop.RequireCurrent(desktop);
            if (publisherOwnership != null || publisherOwnershipUnknown != IntPtr.Zero || publisherBootstrap != null || publisherBootstrapUnknown != IntPtr.Zero)
                throw new InvalidOperationException("A retained Publisher ownership probe refuses a new native launch before observed original exit.");
            showPrivateAccess = null;
            string executable = RequirePrivateOfficeExecutable(Kind, desktop,
                privateExecutable ?? Environment.GetEnvironmentVariable("VBAi_TEST_" + Kind.ToUpperInvariant() + "_EXE"));
            if (!File.Exists(executable)) throw new FileNotFoundException("The reviewed installed Office executable does not exist.", executable);
            string expectedProcess = Kind == "Access" ? "MSACCESS" : Kind == "PowerPoint" ? "POWERPNT" : "MSPUB";
            if (Kind == "Publisher")
            {
                var publisherBefore = ReadCompletePublisherProcessIds();
                if (publisherBefore.Length != 0) throw new InvalidOperationException("A complete empty Publisher process inventory is required before the sole private launch.");
                publisherBootstrapEmptyBefore = true;
                steps.Add(new { PublisherBootstrapPrelaunchProcessIds = publisherBefore, Complete = true });
                FlushAdapterEvidence();
            }
            var existing = Process.GetProcessesByName(expectedProcess);
            try { Assert.AreEqual(0, existing.Length, "Existing Office processes refuse private qualification before launch or activation."); }
            finally { foreach (var process in existing) process.Dispose(); }
            privateDesktop = desktop; privateExecutable = executable;
            string executableHash;
            using (var sha = SHA256.Create())
            using (var file = File.OpenRead(executable)) executableHash = BitConverter.ToString(sha.ComputeHash(file)).Replace("-", "");
            string[] arguments = PrivateOfficeArguments(Kind);
            RecordPrivatePublisherLaunchMode(executable, arguments);
            steps.Add(new { PrivateDesktopLaunchIntent = true, Desktop = desktop, Executable = executable,
                ExecutableSha256 = executableHash, Arguments = arguments, StartAttempts = 1, ComActivationAllowed = false });
            FlushAdapterEvidence();
            privateDesktopChild = IsolatedTestDesktop.Launch(executable, arguments, Root, desktop);
            ProcessId = privateDesktopChild.ProcessId;
            publisherStartupRecovery = Kind == "Publisher" ? new PublisherStartupRecoveryGate((uint)ProcessId,
                privateDesktopChild.ProcessHandle.ToInt64()) : null;
            owned = true; CaptureOwnedProcess();
            // Shutdown observations use the original CreateProcess handle, never a PID reacquisition.
            WaitForOwnedExit = (process, timeout) => {
                Assert.AreSame(ownedProcess, process);
                return privateDesktopChild.Wait(timeout);
            };
            ReadOwnedExitCode = process => { Assert.AreSame(ownedProcess, process); return unchecked((int)privateDesktopChild.ExitCode()); };
            steps.Add(new { PrivateDesktopLaunchObserved = true, ProcessId, ThreadId = privateDesktopChild.ThreadId,
                OriginalHandle = privateDesktopChild.ProcessHandle.ToInt64(), Desktop = desktop });
            FlushAdapterEvidence();
            RequirePrivateHostDesktop(false);
            var watch = Stopwatch.StartNew();
            while (watch.ElapsedMilliseconds < 20000)
            {
                RequirePrivateHostDesktop(false);
                if (ObservePrivatePublisherStartupRecovery()) { Thread.Sleep(50); continue; }
                foreach (string progId in new[] { hostProgId, Kind + ".Application" }.Distinct(StringComparer.OrdinalIgnoreCase))
                {
                    object candidate = null;
                    bool attached = false;
                    try
                    {
                        candidate = Marshal.GetActiveObject(progId); // ROT read only: never CoCreateInstance.
                        application = candidate;
                        if (Kind == "Publisher") BindPrivatePublisherBootstrap();
                        else RequireApplicationOwner();
                        RequirePrivateHostDesktop(true);
                        steps.Add(new { PrivateDesktopApplicationAttached = true, ProcessId, Desktop = desktop,
                            Attachment = "ExistingROT", ProgId = progId, ElapsedMilliseconds = watch.ElapsedMilliseconds });
                        FlushAdapterEvidence();
                        attached = true;
                    }
                    catch (COMException error) when ((Kind != "Publisher" || (candidate == null && publisherOwnership == null && publisherBootstrap == null)) &&
                        (error.ErrorCode == unchecked((int)0x800401E3) || error.ErrorCode == unchecked((int)0x80010001)))
                    {
                        application = null;
                        if (candidate != null) Marshal.ReleaseComObject(candidate);
                    }
                    if (attached)
                    {
                        // Outside ROT-discovery retries: a failed state getter preserves this exact host.
                        if (Kind == "Access") RecordPrivateAccessAutomationState();
                        return;
                    }
                }
                Thread.Sleep(50); // Read-only discovery; the native launch is never replayed.
            }
            throw new InvalidOperationException("BLOCKED: no existing ROT application could be attached to the exact private Office PID; the launched host and original handle are retained without activation, Quit or relaunch.");
        }

        /// <summary>Rejects a host on another desktop before document, visibility, bridge or teardown operations.</summary>
        private void RequirePrivateHostDesktop(bool requireWindow)
        {
            if (privateDesktop == null) return;
            IsolatedTestDesktop.RequireCurrent(privateDesktop);
            if (privateDesktopChild == null || privateDesktopChild.Wait(0) || ownedProcess == null || ownedProcess.HasExited ||
                !string.Equals(ExcelOwnedProcessImage.Read(privateDesktopChild.ProcessHandle), privateExecutable, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The originally launched private Office process/image is unavailable; no native action or relaunch is allowed.");
            // Foreign GetThreadDesktop returned stable access denied in the native v1 campaign;
            // non-GUI threads can also have no USER desktop handle. Explicit HWND membership is the proof.
            IsolatedTestDesktop.RequireOfficeWindowInventory(privateDesktop, (uint)ProcessId, requireWindow, IntPtr.Zero);
        }

        private void ReleasePrivateExitedProcess()
        {
            if (privateDesktopChild == null) return;
            if (!privateDesktopChild.Wait(0)) throw new InvalidOperationException("The original private host handle cannot be released before observed exit.");
            if (publisherOwnershipUnknown != IntPtr.Zero) { Marshal.Release(publisherOwnershipUnknown); publisherOwnershipUnknown = IntPtr.Zero; }
            ReleaseExitedPrivatePublisherBootstrap();
            publisherOwnership = null;
            publisherOwnershipChild = null;
            privateDesktopChild.Dispose(); privateDesktopChild = null;
        }
    }
}
