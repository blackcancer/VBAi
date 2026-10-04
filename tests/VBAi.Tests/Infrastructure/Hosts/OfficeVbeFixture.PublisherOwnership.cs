using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;

namespace VBAi.Tests.Integration
{
    internal sealed partial class OfficeVbeFixture
    {
        private PublisherOwnershipProbe publisherOwnership;
        private IntPtr publisherOwnershipUnknown;
        private IsolatedTestDesktop.NativeChild publisherOwnershipChild;

        // This is a capability probe, not a claim that Publisher.Application implements IOleWindow.
        // https://learn.microsoft.com/en-us/windows/win32/api/oleidl/nn-oleidl-iolewindow
        [ComImport, Guid("00000114-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface PublisherOleWindow
        {
            [PreserveSig] int GetWindow(out IntPtr window);
            [PreserveSig] int ContextSensitiveHelp([MarshalAs(UnmanagedType.Bool)] bool enter);
        }
        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int PublisherGetWindow(IntPtr self, out IntPtr window);
        // Contract confirmed against the installed Office 15 PIA metadata. Native support is
        // still observed separately for each exact owned MsoCommandBar; no default dispatch assumption.
        [ComImport, Guid("000C0304-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
        private interface PublisherNativeCommandBar
        {
            [DispId(0x60020000)]
            object Application { [return: MarshalAs(UnmanagedType.IDispatch)] get; }
        }
        [DllImport("oleacc.dll", PreserveSig = true)]
        private static extern int AccessibleObjectFromWindow(IntPtr window, uint objectId, ref Guid iid,
            out IntPtr result);

        internal sealed class PublisherNativeOmResult
        {
            internal int HResult;
            internal IntPtr Pointer;
        }

        // All injected operations are synchronous on the acquiring STA. The native pointer,
        // typed RCW and getter RCW each own one reference, balanced even when later evidence fails.
        internal sealed class PublisherCommandBarAcquisition
        {
            internal static readonly Guid InterfaceId = new Guid("000C0304-0000-0000-C000-000000000046");
            private bool claimed;

            internal long Read(PublisherOwnershipProbe probe, string operation, Action guard,
                Action<PublisherNativeOmResult> acquire, Func<IntPtr, object> wrap, Func<object, bool> typed,
                Func<object, object> readApplication, Func<object, long> identity,
                Action<IntPtr> releaseRaw, Action<object> releaseObject, Action<IDictionary<string, object>> record)
            {
                if (claimed) throw new InvalidOperationException("This exact command-bar NativeOM attempt cannot be repeated.");
                claimed = true;
                var native = new PublisherNativeOmResult();
                object bar = null, container = null;
                Exception failure = null;
                long result = 0;
                try
                {
                    guard();
                    probe.Observe(operation + ".NativeOM", () => {
                        acquire(native);
                        return new Dictionary<string, object> { ["HResult"] = "0x" + unchecked((uint)native.HResult).ToString("X8"),
                            ["Pointer"] = native.Pointer.ToInt64(), ["RequestedInterfaceId"] = InterfaceId.ToString("D") };
                    }, record);
                    if (native.HResult != 0 || native.Pointer == IntPtr.Zero)
                        throw new COMException("Exact CommandBar NativeOM interface was not returned; no fallback or retry.",
                            native.HResult == 0 ? unchecked((int)0x80004005) : native.HResult);
                    guard();
                    probe.Observe(operation + ".TypedInterface", () => {
                        bar = wrap(native.Pointer);
                        bool supported = bar != null && typed(bar);
                        if (!supported) throw new InvalidOperationException("NativeOM cannot be marshaled as the exact CommandBar interface; no fallback.");
                        return new Dictionary<string, object> { ["InterfaceId"] = InterfaceId.ToString("D"),
                            ["ManagedType"] = bar.GetType().FullName, ["TypedInterfaceSupported"] = true,
                            ["ApplicationDispId"] = "0x60020000" };
                    }, record);
                    guard();
                    probe.Observe(operation + ".Application", () => { container = readApplication(bar); return container; }, record);
                    if (container == null) throw new InvalidOperationException("Typed CommandBar.Application returned no container; no fallback.");
                    guard();
                    result = probe.Observe(operation + ".ApplicationIUnknown", () => identity(container), record);
                    if (result == 0) throw new InvalidOperationException("Typed CommandBar.Application has no canonical COM identity.");
                    guard();
                }
                catch (Exception error) { failure = error; }
                var failures = new List<Exception>();
                if (failure != null) failures.Add(failure);
                // A failing release must not prevent balancing the remaining independently owned references.
                try { if (container != null) releaseObject(container); } catch (Exception error) { failures.Add(error); }
                try { if (bar != null) releaseObject(bar); } catch (Exception error) { failures.Add(error); }
                try { if (native.Pointer != IntPtr.Zero) releaseRaw(native.Pointer); } catch (Exception error) { failures.Add(error); }
                if (failures.Count == 1) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failures[0]).Throw();
                if (failures.Count > 1) throw new AggregateException("CommandBar acquisition and reference cleanup failed; no retry.", failures);
                return result;
            }
        }

        internal sealed class PublisherOwnerWindow
        {
            internal IntPtr Window, Root;
            internal uint Process, Thread, RootProcess, RootThread;
            internal string Class;
            internal bool Same(PublisherOwnerWindow other) => other != null && Window == other.Window && Root == other.Root &&
                Process == other.Process && Thread == other.Thread && RootProcess == other.RootProcess &&
                RootThread == other.RootThread && Class == other.Class;
        }

        internal sealed class PublisherOwnershipProbe
        {
            internal const int NoInterface = unchecked((int)0x80004002);
            internal const int MaximumCommandBars = 32;
            private readonly uint pid;
            private readonly long processHandle;
            private bool claimed, verified, closed;
            private long identity;
            private PublisherOwnerWindow bound;
            internal PublisherOwnershipProbe(uint pid, long processHandle)
            {
                if (pid == 0 || processHandle == 0) throw new ArgumentException("The original Publisher process handle/PID are required.");
                this.pid = pid; this.processHandle = processHandle;
            }

            internal T Observe<T>(string operation, Func<T> read, Action<IDictionary<string, object>> record)
            {
                record(new Dictionary<string, object> { ["PublisherOwnershipRead"] = operation, ["State"] = "PENDING",
                    ["ProcessId"] = pid, ["OriginalHandle"] = processHandle, ["AutomaticRetry"] = false,
                    ["MutationInvoked"] = false, ["Utc"] = DateTime.UtcNow.ToString("o") });
                T value;
                try { value = read(); }
                catch (Exception error)
                {
                    try { record(new Dictionary<string, object> { ["PublisherOwnershipRead"] = operation, ["State"] = "FAILED",
                        ["Error"] = error.ToString(), ["HResult"] = "0x" + unchecked((uint)error.HResult).ToString("X8"),
                        ["AutomaticRetry"] = false, ["MutationInvoked"] = false }); }
                    catch (Exception evidence) { throw new AggregateException("Publisher ownership read and durable failure evidence failed.", error, evidence); }
                    throw;
                }
                record(new Dictionary<string, object> { ["PublisherOwnershipRead"] = operation, ["State"] = "RETURNED",
                    ["Value"] = EvidenceValue(value), ["AutomaticRetry"] = false, ["MutationInvoked"] = false });
                return value;
            }

            internal static object EvidenceValue(object value)
            {
                var window = value as PublisherOwnerWindow;
                if (window != null) return new Dictionary<string, object> { ["Hwnd"] = window.Window.ToInt64(),
                    ["RootHwnd"] = window.Root.ToInt64(), ["ProcessId"] = window.Process, ["NativeThreadId"] = window.Thread,
                    ["RootProcessId"] = window.RootProcess, ["RootThreadId"] = window.RootThread, ["Class"] = window.Class };
                var windows = value as PublisherOwnerWindow[];
                if (windows == null) return value != null && Marshal.IsComObject(value) ? "COM_GETTER_RETURNED" : value;
                var rows = new object[windows.Length];
                for (int i = 0; i < rows.Length; i++) rows[i] = EvidenceValue(windows[i]);
                return rows;
            }

            private void RequireWindow(PublisherOwnerWindow value, bool commandBar)
            {
                if (value == null || value.Window == IntPtr.Zero || value.Root == IntPtr.Zero || value.Process != pid ||
                    value.RootProcess != pid || value.Thread == 0 || value.RootThread == 0 ||
                    (commandBar && value.Class != "MsoCommandBar"))
                    throw new InvalidOperationException("BLOCKED: Publisher ownership has no exact original-process native window; no publication operation is permitted.");
            }

            internal void Bind(Func<long> readIdentity, Action guard, Func<int> queryOleWindow,
                Func<PublisherOwnerWindow> getOleWindow, Func<PublisherOwnerWindow[]> findCommandBars,
                Func<PublisherOwnerWindow, long> readBarApplicationIdentity, Func<IntPtr, PublisherOwnerWindow> rereadWindow,
                Action<IDictionary<string, object>> record)
            {
                if (claimed || closed) throw new InvalidOperationException("Publisher ownership discovery is already claimed or closed for this original process; no replay is permitted.");
                claimed = true; // All failures, including evidence/getter failures, permanently consume discovery.
                guard();
                identity = Observe("Application.IUnknown", readIdentity, record);
                if (identity == 0) throw new InvalidOperationException("BLOCKED: no canonical Publisher application identity.");
                guard();
                int result = Observe("Application.QueryInterface.IOleWindow", queryOleWindow, record);
                PublisherOwnerWindow candidate;
                string route;
                if (result == 0)
                {
                    guard();
                    candidate = Observe("IOleWindow.GetWindow", getOleWindow, record);
                    RequireWindow(candidate, false);
                    route = "ObservedIOleWindow";
                }
                else if (result == NoInterface)
                {
                    record(new Dictionary<string, object> { ["PublisherOwnershipCapability"] = "IOleWindow",
                        ["State"] = "KNOWN_UNSUPPORTED", ["HResult"] = "0x80004002", ["AutomaticRetry"] = false });
                    guard();
                    var bars = Observe("ExistingOwnedMsoCommandBarInventory", findCommandBars, record);
                    if (bars == null || bars.Length == 0 || bars.Length > MaximumCommandBars)
                        throw new InvalidOperationException("BLOCKED: no complete bounded existing owned MsoCommandBar inventory; no publication operation is permitted.");
                    var distinct = new HashSet<IntPtr>();
                    foreach (var bar in bars)
                    {
                        RequireWindow(bar, true);
                        if (!distinct.Add(bar.Window)) throw new InvalidOperationException("BLOCKED: duplicate command-bar HWND in the frozen ownership inventory; no NativeOM call is permitted.");
                    }
                    record(new Dictionary<string, object> { ["PublisherCommandBarInventoryValidated"] = true,
                        ["DistinctWindowCount"] = bars.Length, ["MaximumNativeObjectModelWindows"] = MaximumCommandBars,
                        ["ApplicationIdentityAdoptionAllowed"] = false });
                    candidate = null;
                    foreach (var bar in bars)
                    {
                        string operation = "MsoCommandBar[" + bar.Window.ToInt64() + "]";
                        guard();
                        var before = Observe(operation + ".BeforeNativeOM", () => rereadWindow(bar.Window), record);
                        RequireWindow(before, true);
                        if (!bar.Same(before)) throw new InvalidOperationException("Publisher command-bar identity changed before native object-model retrieval.");
                        if (Observe(operation + ".Application.IUnknown.Before", readIdentity, record) != identity)
                            throw new InvalidOperationException("Publisher application identity changed before native object-model retrieval.");
                        guard();
                        long container = Observe(operation + ".NativeOM.Application.IUnknown", () => readBarApplicationIdentity(bar), record);
                        if (container == 0) throw new InvalidOperationException("BLOCKED: command-bar container identity is missing; no further NativeOM calls are permitted.");
                        guard();
                        var afterBar = Observe(operation + ".AfterNativeOM", () => rereadWindow(bar.Window), record);
                        RequireWindow(afterBar, true);
                        if (!bar.Same(afterBar)) throw new InvalidOperationException("Publisher command-bar identity changed during native object-model retrieval.");
                        if (Observe(operation + ".Application.IUnknown.After", readIdentity, record) != identity)
                            throw new InvalidOperationException("Publisher application identity changed during native object-model retrieval.");
                        record(new Dictionary<string, object> { ["PublisherCommandBarAssociation"] = container == identity ? "EXACT_MATCH" : "KNOWN_MISMATCH",
                            ["Hwnd"] = bar.Window.ToInt64(), ["ContainerIUnknown"] = container, ["RetainedApplicationIUnknown"] = identity,
                            ["NativeObjectModelAttemptsForWindow"] = 1, ["AutomaticRetry"] = false });
                        if (container == identity) { candidate = bar; break; }
                    }
                    if (candidate == null) throw new InvalidOperationException("BLOCKED: every returned command-bar container differs from the retained ROT application; no publication operation is permitted.");
                    route = "OwnedMsoCommandBarNativeOM";
                }
                else throw new COMException("Publisher IOleWindow query failed or is uncertain; no fallback or retry is permitted.", result);
                guard();
                var after = Observe("BoundWindow.AfterOwnershipRead", () => rereadWindow(candidate.Window), record);
                RequireWindow(after, route == "OwnedMsoCommandBarNativeOM");
                if (!candidate.Same(after)) throw new InvalidOperationException("Publisher ownership window changed during its read-only binding.");
                if (Observe("Application.IUnknown.AfterBinding", readIdentity, record) != identity)
                    throw new InvalidOperationException("Publisher application identity changed during ownership discovery.");
                record(new Dictionary<string, object> { ["PublisherPrePublicationOwnership"] = "VERIFIED", ["Route"] = route,
                    ["ProcessId"] = pid, ["OriginalHandle"] = processHandle, ["ApplicationIUnknown"] = identity,
                    ["Hwnd"] = candidate.Window.ToInt64(), ["RootHwnd"] = candidate.Root.ToInt64(),
                    ["NativeThreadId"] = candidate.Thread, ["RootThreadId"] = candidate.RootThread,
                    ["Class"] = candidate.Class, ["MutationInvoked"] = false });
                bound = candidate; verified = true;
            }

            internal void Recheck(uint currentPid, long currentHandle, Func<long> readIdentity, Action guard,
                Func<IntPtr, PublisherOwnerWindow> rereadWindow)
            {
                if (closed || !verified || currentPid != pid || currentHandle != processHandle)
                    throw new InvalidOperationException("BLOCKED: no completed ownership binding for this original Publisher process.");
                guard();
                if (readIdentity() != identity) throw new InvalidOperationException("Publisher ROT application identity changed; no publication operation is permitted.");
                var current = rereadWindow(bound.Window); RequireWindow(current, false);
                if (!bound.Same(current)) throw new InvalidOperationException("Publisher bound native window identity changed; no rediscovery is permitted.");
                guard();
            }

            internal void Close(Action releaseOnce)
            {
                if (closed) throw new InvalidOperationException("Publisher ownership references are already closed; no release or reacquisition replay is permitted.");
                closed = true; verified = false; // Claim before release: an uncertain release must never be replayed.
                releaseOnce();
            }
        }

        private PublisherOwnerWindow ReadPrivatePublisherOwnerWindow(IntPtr window)
        {
            RequirePrivatePublisherOwnershipHost();
            IsolatedTestDesktop.RequireOfficeWindowInventory(privateDesktop, (uint)ProcessId, true, window);
            var result = new PublisherOwnerWindow { Window = window, Root = GetAncestor(window, 2), Class = PublisherStartupClass(window) };
            result.Thread = GetWindowThreadProcessId(window, out result.Process);
            result.RootThread = GetWindowThreadProcessId(result.Root, out result.RootProcess);
            return result;
        }

        private PublisherOwnerWindow[] FindPrivatePublisherCommandBars()
        {
            RequirePrivatePublisherOwnershipHost();
            var bars = new List<PublisherOwnerWindow>(); bool complete = true; int visited = 0, ownedRoots = 0; Exception failure = null;
            string refusal = null;
            var clock = Stopwatch.StartNew();
            DialogWindowCallback inspect = (window, state) => {
                if (!complete) return false;
                try
                {
                    if (++visited > 8192 || clock.ElapsedMilliseconds > 5000)
                    { refusal = visited > 8192 ? "WindowCountBoundExceeded" : "ElapsedTimeBoundExceeded"; complete = false; return false; }
                    uint pid; GetWindowThreadProcessId(window, out pid);
                    if (pid == (uint)ProcessId && PublisherStartupClass(window) == "MsoCommandBar")
                    {
                        bars.Add(ReadPrivatePublisherOwnerWindow(window));
                        if (bars.Count > PublisherOwnershipProbe.MaximumCommandBars)
                        { refusal = "CommandBarCountBoundExceeded"; complete = false; return false; }
                    }
                    return true;
                }
                catch (Exception error) { failure = error; refusal = "NativeWindowGuardFailed"; complete = false; return false; }
            };
            bool rootsComplete = EnumWindows((window, state) => {
                if (!inspect(window, state)) return false;
                uint pid; GetWindowThreadProcessId(window, out pid);
                if (pid == (uint)ProcessId) { ownedRoots++; EnumChildWindows(window, inspect, IntPtr.Zero); }
                return complete;
            }, IntPtr.Zero);
            if (clock.ElapsedMilliseconds > 5000) { refusal = "ElapsedTimeBoundExceeded"; complete = false; }
            RecordPublisherStartup(new Dictionary<string, object> { ["PublisherCommandBarNativeInventory"] = "TERMINAL",
                ["ProcessId"] = ProcessId, ["EnumWindowsReturned"] = rootsComplete, ["CallbacksCompleted"] = complete,
                ["VisitedWindowCount"] = visited, ["OwnedRootCount"] = ownedRoots, ["CommandBarCount"] = bars.Count,
                ["ElapsedMilliseconds"] = clock.ElapsedMilliseconds, ["MaximumVisitedWindows"] = 8192,
                ["MaximumCommandBars"] = PublisherOwnershipProbe.MaximumCommandBars, ["MaximumElapsedMilliseconds"] = 5000,
                ["RefusalReason"] = refusal ?? (rootsComplete ? null : "EnumWindowsFailed"),
                ["Windows"] = PublisherOwnershipProbe.EvidenceValue(bars.ToArray()) });
            if (!rootsComplete || !complete) throw new InvalidOperationException("BLOCKED: Publisher command-bar inventory is incomplete or exceeds its bound.", failure);
            RequirePrivatePublisherOwnershipHost();
            return bars.ToArray();
        }

        private void RequirePrivatePublisherOwnershipHost()
        {
            commandContainment.RequireTerminal();
            if (Kind != "Publisher" || privateDesktop == null || !owned || NativeExecutionUnsettled ||
                ownedProcess == null || shutdownEvidence == null || privateDesktopChild == null ||
                (publisherOwnershipChild != null && !ReferenceEquals(publisherOwnershipChild, privateDesktopChild)) ||
                privateDesktopChild.ProcessId != ProcessId || Thread.CurrentThread.ManagedThreadId != shutdownOwnerThread ||
                Thread.CurrentThread.GetApartmentState() != ApartmentState.STA || ownedProcess.HasExited || ownedProcess.Id != ProcessId ||
                !Equals(shutdownEvidence.Record["ProcessId"], ProcessId) ||
                !Equals(shutdownEvidence.Record["ProcessStartedUtc"], ownedProcess.StartTime.ToUniversalTime().ToString("o")) ||
                !Equals(shutdownEvidence.Record["OriginalProcessHandle"], "0x" + unchecked((ulong)ownedProcess.Handle.ToInt64()).ToString("X16")) ||
                !string.Equals(Convert.ToString(shutdownEvidence.Record["ProcessImage"]), ExcelOwnedProcessImage.Read(ownedProcess.Handle), StringComparison.OrdinalIgnoreCase) ||
                !Equals(shutdownEvidence.Record["QuitEntries"], 0) || Equals(shutdownEvidence.Record["TeardownPrepared"], true))
                throw new InvalidOperationException("Publisher ownership requires its unchanged original native process on the acquiring STA before cleanup.");
            RequirePrivateHostDesktop(true);
        }

        private long ReadPrivatePublisherBarApplicationIdentity(PublisherOwnerWindow window)
        {
            // The only documented Office native-object-model window class used here is MsoCommandBar.
            // https://learn.microsoft.com/en-us/windows/win32/api/oleacc/nf-oleacc-accessibleobjectfromwindow
            // https://learn.microsoft.com/en-us/office/vba/api/office.commandbar.application
            Action guard = () => {
                var current = ReadPrivatePublisherOwnerWindow(window.Window);
                if (current.Class != "MsoCommandBar" || !window.Same(current))
                    throw new InvalidOperationException("The exact documented command-bar window changed during its only NativeOM attempt.");
            };
            var acquisition = new PublisherCommandBarAcquisition();
            PublisherNativeCommandBar typedBar = null;
            return acquisition.Read(publisherOwnership, "NativeCommandBar[" + window.Window.ToInt64() + "]", guard,
                native => {
                    Guid commandBar = PublisherCommandBarAcquisition.InterfaceId;
                    native.HResult = AccessibleObjectFromWindow(window.Window, unchecked((uint)-16), ref commandBar, out native.Pointer);
                // GetTypedObjectForIUnknown requires an imported COM class, not this interface.
                // Retain the RCW before the explicit cast so a refused cast still releases it.
                }, pointer => Marshal.GetObjectForIUnknown(pointer),
                bar => { typedBar = (PublisherNativeCommandBar)bar; return typedBar != null; }, bar => typedBar.Application,
                container => {
                    if (!Marshal.IsComObject(container)) throw new InvalidOperationException("Typed CommandBar.Application did not return a COM container.");
                    IntPtr unknown = Marshal.GetIUnknownForObject(container);
                    try { return unknown.ToInt64(); } finally { Marshal.Release(unknown); }
                }, pointer => { Marshal.Release(pointer); }, value => {
                    if (Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
                }, RecordPublisherStartup);
        }

        private void BindPrivatePublisherOwnership()
        {
            if (publisherOwnership != null || publisherOwnershipUnknown != IntPtr.Zero)
                throw new InvalidOperationException("Publisher ownership discovery cannot be replayed for the original native child.");
            var child = privateDesktopChild; var candidate = application;
            publisherOwnershipChild = child;
            publisherOwnership = new PublisherOwnershipProbe((uint)ProcessId, child.ProcessHandle.ToInt64());
            Action guard = () => {
                if (!ReferenceEquals(child, privateDesktopChild) || !ReferenceEquals(candidate, application))
                    throw new InvalidOperationException("Publisher native child or ROT application changed during ownership discovery.");
                RequirePrivatePublisherOwnershipHost();
            };
            IntPtr ole = IntPtr.Zero;
            try
            {
                publisherOwnership.Bind(() => {
                    if (publisherOwnershipUnknown == IntPtr.Zero) publisherOwnershipUnknown = Marshal.GetIUnknownForObject(candidate);
                    IntPtr current = Marshal.GetIUnknownForObject(candidate);
                    try { if (current != publisherOwnershipUnknown) throw new InvalidOperationException("Canonical Publisher application identity changed."); return current.ToInt64(); }
                    finally { Marshal.Release(current); }
                }, guard, () => {
                    Guid iid = typeof(PublisherOleWindow).GUID;
                    int result = Marshal.QueryInterface(publisherOwnershipUnknown, ref iid, out ole);
                    if ((result == 0) != (ole != IntPtr.Zero) || (result != 0 && ole != IntPtr.Zero))
                        throw new InvalidOperationException("Publisher IOleWindow query returned an inconsistent pointer; no fallback.");
                    return result;
                }, () => {
                    // Invoke slot 3 on the exact successfully queried interface. No RCW cast can repeat QI.
                    IntPtr slot = Marshal.ReadIntPtr(Marshal.ReadIntPtr(ole), 3 * IntPtr.Size);
                    var getWindow = Marshal.GetDelegateForFunctionPointer<PublisherGetWindow>(slot);
                    IntPtr window; int result = getWindow(ole, out window);
                    if (result != 0) throw new COMException("Publisher IOleWindow.GetWindow failed; no fallback or retry.", result);
                    return ReadPrivatePublisherOwnerWindow(window);
                }, FindPrivatePublisherCommandBars, ReadPrivatePublisherBarApplicationIdentity,
                    ReadPrivatePublisherOwnerWindow, RecordPublisherStartup);
            }
            finally { if (ole != IntPtr.Zero) Marshal.Release(ole); }
        }

        private void RecheckPrivatePublisherOwnership()
        {
            if (publisherOwnership == null || publisherOwnershipUnknown == IntPtr.Zero || privateDesktopChild == null ||
                !ReferenceEquals(publisherOwnershipChild, privateDesktopChild))
                throw new InvalidOperationException("BLOCKED: Publisher has no retained pre-publication ownership proof.");
            publisherOwnership.Recheck((uint)ProcessId, privateDesktopChild.ProcessHandle.ToInt64(), () => {
                IntPtr current = Marshal.GetIUnknownForObject(application);
                try { return current.ToInt64(); } finally { Marshal.Release(current); }
            }, RequirePrivatePublisherOwnershipHost, ReadPrivatePublisherOwnerWindow);
        }

        private void ClosePrivatePublisherOwnershipAfterQuit(bool externalReferencesReleased)
        {
            if (Kind != "Publisher" || privateDesktop == null || publisherOwnership == null) return;
            commandContainment.RequireTerminal();
            if (!externalReferencesReleased || NativeExecutionUnsettled || application != null || document != null ||
                Thread.CurrentThread.ManagedThreadId != shutdownOwnerThread || Thread.CurrentThread.GetApartmentState() != ApartmentState.STA ||
                privateDesktopChild == null || !ReferenceEquals(publisherOwnershipChild, privateDesktopChild) ||
                privateDesktopChild.ProcessId != ProcessId || shutdownEvidence == null ||
                !Equals(shutdownEvidence.Record["ProcessId"], ProcessId) || !Equals(shutdownEvidence.Record["QuitEntries"], 1) ||
                !Equals(shutdownEvidence.Record["QuitOutcome"], "RETURNED"))
                throw new InvalidOperationException("Publisher canonical reference release requires settled original-host Quit and COM scopes on their owning STA.");
            RecordPublisherStartup(new Dictionary<string, object> { ["PublisherOwnershipReferenceRelease"] = "PENDING",
                ["ProcessId"] = ProcessId, ["OriginalHandle"] = privateDesktopChild.ProcessHandle.ToInt64(), ["ReleaseAttempts"] = 1 });
            publisherOwnership.Close(() => {
                IntPtr held = publisherOwnershipUnknown; publisherOwnershipUnknown = IntPtr.Zero;
                if (held != IntPtr.Zero) Marshal.Release(held);
            });
            RecordPublisherStartup(new Dictionary<string, object> { ["PublisherOwnershipReferenceRelease"] = "RETURNED",
                ["ProcessId"] = ProcessId, ["ReleaseAttempts"] = 1, ["CanonicalReferenceReacquisitionAllowed"] = false });
        }
    }
}
