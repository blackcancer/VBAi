using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;

namespace VBAi.Tests.Integration
{
    internal sealed partial class OfficeVbeFixture
    {
        private bool publisherBootstrapEmptyBefore;
        private PublisherBootstrapGate publisherBootstrap;
        private PublisherOpenSecurityGate publisherOpenSecurity;
        private IntPtr publisherBootstrapUnknown;
        private IsolatedTestDesktop.NativeChild publisherBootstrapChild;
        private object publisherBootstrapApplication;

        internal sealed class PublisherBootstrapSnapshot
        {
            internal int ProcessId, SessionId, OwnerThread;
            internal long Handle;
            internal string StartedUtc, Image;
            internal bool Alive, PrivateWindowsVerified, MainWindowsVerified, MainPublisher, NoVisibleModal, Sta;
            internal int[] PublisherProcessIds;
            internal void Require()
            {
                if (ProcessId <= 0 || Handle == 0 || string.IsNullOrWhiteSpace(StartedUtc) || string.IsNullOrWhiteSpace(Image) ||
                    OwnerThread <= 0 || SessionId < 0 || !Alive ||
                    (MainPublisher ? !MainWindowsVerified || PrivateWindowsVerified : !PrivateWindowsVerified || MainWindowsVerified) ||
                    !NoVisibleModal || !Sta ||
                    PublisherProcessIds == null || PublisherProcessIds.Length != 1 || PublisherProcessIds[0] != ProcessId)
                    throw new InvalidOperationException("Publisher bootstrap requires one unchanged original private process; no publication mutation is permitted.");
            }
            internal bool Same(PublisherBootstrapSnapshot other) => other != null && ProcessId == other.ProcessId &&
                Handle == other.Handle && StartedUtc == other.StartedUtc && SessionId == other.SessionId && OwnerThread == other.OwnerThread &&
                string.Equals(Image, other.Image, StringComparison.OrdinalIgnoreCase) && MainPublisher == other.MainPublisher;
        }

        /// <summary>One original-process preparation strengthens programmatic Open security; no restoration or replay.</summary>
        internal sealed class PublisherOpenSecurityGate
        {
            private bool consumed, verified;
            internal bool Consumed => consumed;
            internal bool Verified => verified;
            internal void Prepare(Action guard, Func<int> readSecurity, Action<int> writeSecurity,
                Action<IDictionary<string, object>> record)
            {
                if (consumed) throw new InvalidOperationException("Publisher Open security preparation is already consumed; no setter or getter retry.");
                consumed = true;
                record(Row("PreparationClaimed", false));
                guard();
                record(Row("InitialReadPending", false));
                guard();
                int initial = readSecurity();
                guard();
                if (initial < 1 || initial > 3) throw new InvalidOperationException("Unknown Publisher AutomationSecurity value; Open is refused.");
                record(Row("InitialReadReturned", false, initial));
                if (initial != 3)
                {
                    record(Row("ForceDisableClaimed", false, 3));
                    guard();
                    record(Row("ForceDisablePending", false, 3));
                    guard();
                    try { writeSecurity(3); }
                    catch (Exception error)
                    {
                        var failure = Row("ForceDisableUncertain", true, 3); failure["Error"] = error.ToString();
                        try { record(failure); } catch (Exception evidence) { throw new AggregateException(error, evidence); }
                        throw;
                    }
                    record(Row("ForceDisableReturned", true, 3));
                    guard();
                }
                record(Row("FinalReadPending", false));
                guard();
                int final = readSecurity();
                guard();
                record(Row("FinalReadReturned", false, final));
                if (final != 3) throw new InvalidOperationException("Publisher AutomationSecurity ForceDisable was not verified; no Open or security-setting retry.");
                record(Row("ForceDisableVerified", false, final));
                guard();
                verified = true;
            }
            private static IDictionary<string, object> Row(string state, bool mutation, int? value = null) =>
                new Dictionary<string, object>
                {
                    ["PublisherOpenSecurity"] = state,
                    ["AutomationSecurity"] = value,
                    ["MutationInvoked"] = mutation,
                    ["Scope"] = "ExactOriginalOwnedPublisherProcess",
                    ["PropertyDispId"] = 76,
                    ["AutomaticRetry"] = false,
                    ["RestoreDefault"] = false,
                    ["GlobalTrustChanges"] = false,
                    ["Utc"] = DateTime.UtcNow.ToString("o")
                };
        }

        /// <summary>Provisional sole-process inference permits one empty publication/open only; final native proof is mandatory.</summary>
        internal sealed class PublisherBootstrapGate
        {
            private readonly PublisherBootstrapSnapshot original;
            private long identity;
            private bool bindClaimed, bound, mutationClaimed, returned, verified, closed;
            internal bool Closed => closed;
            internal PublisherBootstrapGate(bool emptyBefore, PublisherBootstrapSnapshot snapshot)
            {
                if (!emptyBefore) throw new InvalidOperationException("A complete empty Publisher prelaunch inventory is required.");
                snapshot.Require();
                original = new PublisherBootstrapSnapshot
                {
                    ProcessId = snapshot.ProcessId,
                    Handle = snapshot.Handle,
                    StartedUtc = snapshot.StartedUtc,
                    Image = snapshot.Image,
                    SessionId = snapshot.SessionId,
                    OwnerThread = snapshot.OwnerThread,
                    MainPublisher = snapshot.MainPublisher
                };
            }
            private void RequireSnapshot(Func<PublisherBootstrapSnapshot> read)
            {
                var current = read(); current.Require();
                if (!original.Same(current)) throw new InvalidOperationException("The original Publisher process/owner thread changed; no rebind or mutation is permitted.");
            }
            internal void Bind(Func<PublisherBootstrapSnapshot> read, Func<long> canonicalIdentity,
                Func<bool> queryInstalledApplication, Action<IDictionary<string, object>> record)
            {
                if (bindClaimed || closed) throw new InvalidOperationException("Publisher bootstrap application binding is already consumed.");
                bindClaimed = true;
                RequireSnapshot(read);
                identity = canonicalIdentity();
                if (identity == 0) throw new InvalidOperationException("Publisher application has no canonical COM identity.");
                record(Row("ApplicationInterfacePending", false));
                RequireSnapshot(read);
                if (!queryInstalledApplication()) throw new InvalidOperationException("The exact installed Publisher _Application interface was not returned.");
                RequireSnapshot(read);
                if (canonicalIdentity() != identity) throw new InvalidOperationException("Publisher application identity changed during its sole interface query.");
                bound = true;
                record(Row("ProvisionalBindingReturned", false));
            }
            internal void Recheck(Func<PublisherBootstrapSnapshot> read, Func<long> canonicalIdentity)
            {
                if (!bound || closed) throw new InvalidOperationException("No live provisional Publisher bootstrap binding exists.");
                RequireSnapshot(read);
                if (canonicalIdentity() != identity) throw new InvalidOperationException("The retained Publisher ROT application changed; no reacquisition is permitted.");
                RequireSnapshot(read);
            }
            internal T Invoke<T>(string operation, Func<PublisherBootstrapSnapshot> read, Func<long> canonicalIdentity,
                Func<T> mutation, Action<IDictionary<string, object>> record)
            {
                RequireAvailablePublication(operation);
                Recheck(read, canonicalIdentity);
                mutationClaimed = true;
                var claim = Row("PublicationMutationClaimed", false); claim["Operation"] = operation; record(claim);
                Recheck(read, canonicalIdentity);
                var pending = Row("PublicationMutationPending", true); pending["Operation"] = operation; record(pending);
                T result;
                try { result = mutation(); }
                catch (Exception error)
                {
                    var failed = Row("PublicationMutationUncertain", true); failed["Operation"] = operation; failed["Error"] = error.ToString();
                    try { record(failed); } catch (Exception evidence) { throw new AggregateException(error, evidence); }
                    throw;
                }
                returned = true;
                var terminal = Row("PublicationMutationReturned", true); terminal["Operation"] = operation; record(terminal);
                Recheck(read, canonicalIdentity);
                return result;
            }
            internal void RequireAvailablePublication(string operation)
            {
                if (operation != "NewDocument" && operation != "Open") throw new ArgumentException("Only the frozen publication bootstrap is permitted.");
                if (mutationClaimed || closed || !bound) throw new InvalidOperationException("Publisher publication bootstrap is unavailable or already consumed; no preparation or publication replay.");
            }
            internal void ConfirmNativePublication(Action exactNativeProof)
            {
                if (closed || !bound || !mutationClaimed || !returned)
                    throw new InvalidOperationException("A known returned Publisher publication bootstrap is required before final native proof.");
                exactNativeProof(); verified = true;
            }
            internal void RequireNativePublication()
            { if (closed || !verified) throw new InvalidOperationException("Native publication ownership is unverified; Save and VBE UI operations are refused."); }
            internal void Close(Action release)
            {
                if (closed) throw new InvalidOperationException("Publisher bootstrap references are already closed; no release replay.");
                closed = true; bound = verified = false; release();
            }
            private IDictionary<string, object> Row(string state, bool invoked) => new Dictionary<string, object>
            {
                ["PublisherBootstrap"] = state,
                ["ProcessId"] = original.ProcessId,
                ["OriginalHandle"] = original.Handle,
                ["Association"] = "ProvisionalEmptyInventoryAndExactOriginalProcess",
                ["NativeWindowAssociationVerified"] = false,
                ["ApplicationIUnknown"] = identity,
                ["MutationInvoked"] = invoked,
                ["AutomaticRetry"] = false,
                ["Utc"] = DateTime.UtcNow.ToString("o")
            };
        }

        // Complete bounded system process snapshot; no PID adoption and no content/metadata reads of other applications.
        private static int[] ReadCompletePublisherProcessIds()
        {
            IntPtr snapshot = CreateToolhelp32Snapshot(2, 0);
            if (snapshot == new IntPtr(-1)) throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
            var ids = new List<int>(); var watch = Stopwatch.StartNew(); int visited = 0;
            try
            {
                var entry = new PublisherProcessEntry { Size = (uint)Marshal.SizeOf(typeof(PublisherProcessEntry)) };
                bool next = ReadPublisherProcessFirst(snapshot, ref entry);
                while (next)
                {
                    if (++visited > 16384 || watch.ElapsedMilliseconds > 5000) throw new InvalidOperationException("Publisher process inventory exceeded its complete-snapshot bound.");
                    if (string.Equals(entry.ImageName, "MSPUB.EXE", StringComparison.OrdinalIgnoreCase)) ids.Add(checked((int)entry.ProcessId));
                    next = ReadPublisherProcessNext(snapshot, ref entry);
                }
                int error = Marshal.GetLastWin32Error();
                if (error != 18 || watch.ElapsedMilliseconds > 5000) throw new InvalidOperationException("Publisher process inventory did not reach known end-of-list: " + error);
                return ids.OrderBy(id => id).ToArray();
            }
            finally { CloseHandle(snapshot); }
        }

        private void RequireNoVisiblePublisherBootstrapModal()
        {
            bool complete = true; int visited = 0; var watch = Stopwatch.StartNew(); var dialogs = new List<object>(); Exception failure = null;
            bool enumerated = EnumWindows((window, state) =>
            {
                try
                {
                    if (++visited > 8192 || watch.ElapsedMilliseconds > 5000) { complete = false; return false; }
                    uint pid; uint tid = GetWindowThreadProcessId(window, out pid);
                    if (pid == (uint)ProcessId && IsWindowVisible(window))
                    {
                        string name = PublisherStartupClass(window);
                        if (name == "#32770" || name == "NUIDialog" || name.StartsWith("bosa_sdm", StringComparison.OrdinalIgnoreCase))
                            dialogs.Add(new { Hwnd = window.ToInt64(), ProcessId = pid, NativeThreadId = tid, Class = name, Enabled = IsWindowEnabled(window) });
                    }
                    return true;
                }
                catch (Exception error) { failure = error; complete = false; return false; }
            }, IntPtr.Zero);
            if (!enumerated || !complete || watch.ElapsedMilliseconds > 5000 || dialogs.Count != 0)
            {
                RecordPublisherStartup(new Dictionary<string, object>
                {
                    ["PublisherBootstrapModalInventory"] = "REFUSED",
                    ["Dialogs"] = dialogs.ToArray(),
                    ["Complete"] = enumerated && complete,
                    ["VisitedWindows"] = visited
                });
                throw new InvalidOperationException("An unknown visible Publisher dialog or incomplete inventory refuses bootstrap; no native mutation is permitted.", failure);
            }
        }

        private PublisherBootstrapSnapshot ReadPrivatePublisherBootstrapSnapshot()
        {
            RequirePrivatePublisherOwnershipHost();
            if (publisherBootstrapChild != null && !ReferenceEquals(publisherBootstrapChild, privateDesktopChild))
                throw new InvalidOperationException("The retained Publisher bootstrap native child changed.");
            RequireNoVisiblePublisherBootstrapModal();
            int session;
            using (var current = Process.GetCurrentProcess()) session = current.SessionId;
            if (ownedProcess.SessionId != session) throw new InvalidOperationException("Publisher bootstrap child belongs to another user session.");
            return new PublisherBootstrapSnapshot
            {
                ProcessId = ProcessId,
                Handle = privateDesktopChild.ProcessHandle.ToInt64(),
                StartedUtc = ownedProcess.StartTime.ToUniversalTime().ToString("o"),
                Image = ExcelOwnedProcessImage.Read(privateDesktopChild.ProcessHandle),
                SessionId = session,
                OwnerThread = Thread.CurrentThread.ManagedThreadId,
                Sta = Thread.CurrentThread.GetApartmentState() == ApartmentState.STA,
                Alive = !privateDesktopChild.Wait(0),
                PrivateWindowsVerified = !mainPublisherDesktop,
                MainWindowsVerified = mainPublisherDesktop,
                MainPublisher = mainPublisherDesktop,
                NoVisibleModal = true,
                PublisherProcessIds = ReadCompletePublisherProcessIds()
            };
        }

        private long ReadPrivatePublisherBootstrapIdentity()
        {
            if (application == null || !ReferenceEquals(application, publisherBootstrapApplication) || publisherBootstrapUnknown == IntPtr.Zero)
                throw new InvalidOperationException("The original Publisher ROT candidate is unavailable; no fallback or reacquisition.");
            IntPtr current = Marshal.GetIUnknownForObject(application);
            try { if (current != publisherBootstrapUnknown) throw new InvalidOperationException("Publisher canonical application changed."); return current.ToInt64(); }
            finally { Marshal.Release(current); }
        }

        private void BindPrivatePublisherBootstrap()
        {
            if (publisherBootstrap != null || publisherBootstrapUnknown != IntPtr.Zero || publisherOwnership != null)
                throw new InvalidOperationException("Publisher bootstrap binding cannot be retried or combined with another ownership discovery.");
            publisherBootstrapChild = privateDesktopChild; publisherBootstrapApplication = application;
            publisherBootstrap = new PublisherBootstrapGate(publisherBootstrapEmptyBefore, ReadPrivatePublisherBootstrapSnapshot());
            publisherOpenSecurity = new PublisherOpenSecurityGate();
            publisherBootstrapUnknown = Marshal.GetIUnknownForObject(application);
            publisherBootstrap.Bind(ReadPrivatePublisherBootstrapSnapshot, ReadPrivatePublisherBootstrapIdentity, () =>
            {
                // Installed Publisher _Application IID; no IOleWindow/CommandBar capability assumptions.
                Guid iid = new Guid("0002123e-0000-0000-c000-000000000046"); IntPtr typed = IntPtr.Zero, canonical = IntPtr.Zero;
                try
                {
                    int hr = Marshal.QueryInterface(publisherBootstrapUnknown, ref iid, out typed);
                    RecordPublisherStartup(new Dictionary<string, object>
                    {
                        ["PublisherBootstrapApplicationQI"] = "RETURNED",
                        ["InterfaceId"] = iid.ToString("D"),
                        ["HResult"] = "0x" + unchecked((uint)hr).ToString("X8"),
                        ["Pointer"] = typed.ToInt64()
                    });
                    if (hr != 0 || typed == IntPtr.Zero) throw new COMException("Publisher _Application interface query did not return an exact supported pointer.", hr == 0 ? unchecked((int)0x80004005) : hr);
                    Guid unknown = new Guid("00000000-0000-0000-C000-000000000046");
                    int identityResult = Marshal.QueryInterface(typed, ref unknown, out canonical);
                    if (identityResult != 0 || canonical != publisherBootstrapUnknown)
                        throw new InvalidOperationException("Publisher typed _Application interface has another canonical COM identity.");
                    return true;
                }
                finally { try { if (canonical != IntPtr.Zero) Marshal.Release(canonical); } finally { if (typed != IntPtr.Zero) Marshal.Release(typed); } }
            }, RecordPublisherStartup);
        }

        private void RecheckPrivatePublisherBootstrap() => publisherBootstrap.Recheck(ReadPrivatePublisherBootstrapSnapshot, ReadPrivatePublisherBootstrapIdentity);
        private object InvokePrivatePublisherBootstrap(string operation, Func<object> mutation)
        {
            if (privateDesktop == null) return mutation();
            if (Kind != "Publisher" || publisherBootstrap == null) throw new InvalidOperationException("Private Publisher has no frozen provisional bootstrap gate.");
            if (!SamePublicationPath(DocumentPath, Path.Combine(Root, "Disposable.pub")) || (operation == "Open" && !File.Exists(DocumentPath)))
                throw new InvalidOperationException("Publisher bootstrap permits only the exact original disposable publication path.");
            publisherBootstrap.RequireAvailablePublication(operation);
            if (operation == "Open")
            {
                if (publisherOpenSecurity == null) throw new InvalidOperationException("Original Publisher generation has no Open security gate.");
                var originalBootstrap = publisherBootstrap;
                var originalSecurity = publisherOpenSecurity;
                Action requireOriginalGeneration = () =>
                {
                    if (!ReferenceEquals(originalBootstrap, publisherBootstrap) || !ReferenceEquals(originalSecurity, publisherOpenSecurity))
                        throw new InvalidOperationException("The original Publisher Open security generation changed; no adoption or retry.");
                    originalBootstrap.Recheck(ReadPrivatePublisherBootstrapSnapshot, ReadPrivatePublisherBootstrapIdentity);
                };
                // Official Publisher _Application AutomationSecurity DISPID76, MsoAutomationSecurity I4: ForceDisable=3.
                // Keep this stronger mode until this disposable process ends; never enable macros after an uncertain Open.
                originalSecurity.Prepare(requireOriginalGeneration,
                    () => Convert.ToInt32(((dynamic)application).AutomationSecurity),
                    value => ((dynamic)application).AutomationSecurity = value, RecordPublisherStartup);
            }
            return publisherBootstrap.Invoke(operation, ReadPrivatePublisherBootstrapSnapshot, ReadPrivatePublisherBootstrapIdentity, mutation, RecordPublisherStartup);
        }
        private void RequirePrivatePublisherNativePublication()
        {
            if (Kind != "Publisher" || privateDesktop == null || publisherBootstrap == null) return;
            publisherBootstrap.RequireNativePublication(); RecheckPrivatePublisherBootstrap();
        }
        internal static void RequirePublisherDocumentIdentities(string retained, string active, string sole)
        {
            if (string.IsNullOrWhiteSpace(retained) || retained != active || retained != sole)
                throw new InvalidOperationException("The retained, active and sole Publisher documents do not have the same canonical COM identity.");
        }

        private void ClosePrivatePublisherBootstrapAfterQuit(bool externalReferencesReleased)
        {
            if (publisherBootstrap == null) return;
            commandContainment.RequireTerminal();
            if (!externalReferencesReleased || NativeExecutionUnsettled || application != null || document != null ||
                Thread.CurrentThread.ManagedThreadId != shutdownOwnerThread || Thread.CurrentThread.GetApartmentState() != ApartmentState.STA ||
                !ReferenceEquals(publisherBootstrapChild, privateDesktopChild) || shutdownEvidence == null ||
                !Equals(shutdownEvidence.Record["ProcessId"], ProcessId) || !Equals(shutdownEvidence.Record["QuitEntries"], 1) ||
                !Equals(shutdownEvidence.Record["QuitOutcome"], "RETURNED"))
                throw new InvalidOperationException("Publisher bootstrap reference release requires settled original Quit and COM scopes on the owner STA.");
            ReleasePrivatePublisherBootstrapReference();
        }
        private void ReleasePrivatePublisherBootstrapReference()
        {
            RecordPublisherStartup(new Dictionary<string, object> { ["PublisherBootstrapReferenceRelease"] = "PENDING", ["ReleaseAttempts"] = 1 });
            publisherBootstrap.Close(() =>
            {
                IntPtr held = publisherBootstrapUnknown; publisherBootstrapUnknown = IntPtr.Zero;
                publisherBootstrapApplication = null; if (held != IntPtr.Zero) Marshal.Release(held);
            });
            RecordPublisherStartup(new Dictionary<string, object> { ["PublisherBootstrapReferenceRelease"] = "RETURNED", ["ReleaseAttempts"] = 1 });
        }
        private void ReleaseExitedPrivatePublisherBootstrap()
        {
            if (publisherBootstrap == null) return;
            if (!ReferenceEquals(publisherBootstrapChild, privateDesktopChild) || !privateDesktopChild.Wait(0) ||
                Thread.CurrentThread.ManagedThreadId != shutdownOwnerThread || Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                throw new InvalidOperationException("Only observed original-process exit permits disposing a refused Publisher bootstrap reference.");
            if (!publisherBootstrap.Closed) ReleasePrivatePublisherBootstrapReference();
            publisherBootstrap = null; publisherOpenSecurity = null; publisherBootstrapChild = null; publisherBootstrapEmptyBefore = false;
        }
    }
}
