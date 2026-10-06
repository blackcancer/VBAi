using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace VBAi
{

    /// <summary>Coordinates native SOLIDWORKS project creation and lifecycle verification.</summary>
    internal sealed partial class VbeProjectComponents
    {

        /// <summary>Reports progress and limits for one native macro-creation request.</summary>
        internal sealed class SolidWorksMacroCreationResult
        {

            /// <summary>Records verification, destination creation, uncertain outcome, command/save milestones, dialog closure, and terminal completion.</summary>
            public bool Verified, DestinationCreated, Uncertain, CommandEntered, OriginalCommandReturned,
                FilenameWriteEntered, SaveQueued, DialogClosed, Terminal;

            /// <summary>Counts claimed attempts; each native creation, filename write, and Save is limited to one.</summary>
            public int CommandAttempts, FilenameAttempts, SaveAttempts;

            /// <summary>Verified project name, requested host path, resulting project and collection revisions, and any failure detail.</summary>
            public string Project, HostPath, ProjectVersion, CollectionVersion, Error;

            /// <summary>Ordered durable claims recorded before each native mutation phase.</summary>
            private readonly List<MacroMutationClaim> claims = new List<MacroMutationClaim>();

            /// <summary>Gets the claims.</summary>
            /// <value>Current claims exposed by solid works macro creation result.</value>
            public MacroMutationClaim[] Claims => claims.ToArray();

            /// <summary>Appends a durable mutation claim to the result's ordered history.</summary>
            /// <param name="claim">Phase, destination path, and ordinal claim to retain.</param>
            internal void AddClaim(MacroMutationClaim claim) { claims.Add(claim); }

            /// <summary>Gets whether automatic retry is permitted after this native mutation flow.</summary>
            /// <value>Always <see langword="false"/> because a failed or uncertain native mutation is never replayed.</value>
            public bool RetryAllowed => false;

            /// <summary>Gets whether this flow removed a partially created native project.</summary>
            /// <value>Always <see langword="false"/>; partial host files and projects require explicit inspection.</value>
            public bool RollbackPerformed => false;

            /// <summary>Gets whether persisted contents were verified after closing and reopening the file.</summary>
            /// <value>Always <see langword="false"/>; current verification checks the live host project only.</value>
            public bool PersistenceReloadVerified => false;

            /// <summary>Gets whether a newly added project was matched to the requested destination path.</summary>
            /// <value>Mirrors <see cref="DestinationCreated"/>; it does not qualify reopened persistence.</value>
            public bool DestinationIdentityVerified => DestinationCreated;

            /// <summary>Gets whether the destination was proven absent after an uncertain partial operation.</summary>
            /// <value>Always <see langword="false"/>; callers must inspect the host path without cleanup or retry.</value>
            public bool DestinationAbsenceProven => false;

            /// <summary>Gets the limit.</summary>
            /// <value>Current limit exposed by solid works macro creation result.</value>
            public string Limit => "DestinationCreated reports a verified native destination identity. An unverified partial file may exist at HostPath; inspect locally without automatic cleanup or retry. Fresh reopen is required to qualify persisted contents.";
        }

        /// <summary>Durable record of one claimed native mutation phase.</summary>
        internal sealed class MacroMutationClaim
        {

            /// <summary>Mutation phase and exact canonical host destination captured for the claim.</summary>
            public readonly string Phase, HostPath;

            /// <summary>One-based sequence of claims within this creation operation.</summary>
            public readonly int Ordinal;

            /// <summary>UTC time at which the mutation phase was claimed.</summary>
            public readonly DateTime Utc;

            /// <summary>Creates a timestamped claim immediately before a native mutation phase.</summary>
            /// <param name="phase">Phase label such as BeforeNewMacro, BeforeFilename, BeforeSave, or Terminal.</param>
            /// <param name="path">Canonical absolute .swp destination associated with this operation.</param>
            /// <param name="ordinal">One-based order among the operation's claims.</param>
            internal MacroMutationClaim(string phase, string path, int ordinal)
            { Phase = phase; HostPath = path; Ordinal = ordinal; Utc = DateTime.UtcNow; }
        }

            /// <summary>Snapshot of the native New Macro dialog identity and filename field state.</summary>
        internal sealed class SolidWorksMacroDialog
        {

            /// <summary>HWNDs for the dialog, filename edit control, and Save button.</summary>
            internal IntPtr Window, Filename, SaveButton;

            /// <summary>Native UI thread that owns the dialog HWND.</summary>
            internal uint Thread;

            /// <summary>Current filename edit contents read back from the dialog.</summary>
            internal string FilenameText;

            /// <summary>Compares the dialog and actionable-control HWNDs and owner thread.</summary>
            /// <param name="other">Fresh snapshot to compare with this frozen dialog identity.</param>
            /// <returns><see langword="true"/> when all captured native handles and the owner thread match.</returns>
            internal bool Same(SolidWorksMacroDialog other) => other != null && Window == other.Window &&
                Filename == other.Filename && SaveButton == other.SaveButton && Thread == other.Thread;
        }

        /// <summary>Native operations used by the one-shot SOLIDWORKS macro creation state machine.</summary>
        internal interface ISolidWorksMacroCreationNative : IDisposable
        {

            /// <summary>Throws unless native calls run on the captured SOLIDWORKS owner thread.</summary>
            void RequireOwner();

            /// <summary>Captures and validates the native application's version and macro command before entry.</summary>
            void Prepare();

            /// <summary>Enters the native New Macro command once after the final caller-supplied checks.</summary>
            /// <param name="beforeEntry">Checks to run immediately before command delivery; may refuse the mutation.</param>
            void Create(Action beforeEntry);

            /// <summary>Finds and snapshots the uniquely qualified New Macro dialog, if it is currently present.</summary>
            /// <returns>Verified dialog snapshot, or <see langword="null"/> when no matching dialog is visible.</returns>
            SolidWorksMacroDialog Capture();

            /// <summary>Revalidates the frozen dialog and child control identities before a mutation.</summary>
            /// <param name="expected">Previously captured dialog identity that must still be current.</param>
            void RequireSame(SolidWorksMacroDialog expected);

            /// <summary>Writes the canonical destination to the frozen dialog's filename field once.</summary>
            /// <param name="expected">Dialog snapshot whose identity and edit control are rechecked.</param>
            /// <param name="path">Canonical absolute destination path ending in .swp.</param>
            /// <param name="beforeEntry">Checks immediately before beginning the filename mutation.</param>
            /// <param name="beforeDelivery">Final checks immediately before sending the native edit message.</param>
            void WriteFilename(SolidWorksMacroDialog expected, string path, Action beforeEntry, Action beforeDelivery);

            /// <summary>Queues the dialog's one Save action after rechecking its frozen identity.</summary>
            /// <param name="expected">Dialog snapshot whose Save button must still be current.</param>
            /// <param name="beforeEntry">Checks immediately before beginning the Save action.</param>
            /// <param name="beforeDelivery">Final checks immediately before posting the native button command.</param>
            void Save(SolidWorksMacroDialog expected, Action beforeEntry, Action beforeDelivery);

            /// <summary>Checks whether the frozen dialog has closed and no matching dialog replaced it.</summary>
            /// <param name="expected">Dialog identity captured before filename and Save delivery.</param>
            /// <returns><see langword="true"/> only when that dialog is gone and no replacement is detected.</returns>
            bool Closed(SolidWorksMacroDialog expected);

            /// <summary>Compares canonical project COM identity rather than project names or paths.</summary>
            /// <param name="first">Original project RCW captured before native creation.</param>
            /// <param name="second">Current project RCW resolved after a native transition.</param>
            /// <returns>Whether both references identify the same COM project instance.</returns>
            bool SameProject(object first, object second);
        }

        /// <summary>Factory for the native-dialog adapter, injectable by tests without activating a host.</summary>
        internal Func<object, Action, ISolidWorksMacroCreationNative> SolidWorksMacroCreationNativeFactory =
            (editor, context) => new NativeSolidWorksMacroCreation(editor, context);

        /// <summary>Factory for the owner-STA scheduler that posts command work and polls the modal dialog.</summary>
        internal Func<VbeProjectGeneralOperation.IScheduler> SolidWorksMacroCreationSchedulerFactory =
            () => new SolidWorksMacroCreationScheduler();

        /// <summary>Deadline for the asynchronous creation operation, measured by its scheduler in milliseconds.</summary>
        internal long SolidWorksMacroCreationTimeoutMilliseconds = 30000;

        // This is an explicit native creation route. It never calls Add101/SaveAs or retries a failed save.
        /// <summary>Creates one fresh native .swp project through the original SOLIDWORKS New Macro dialog and verifies the live project.</summary>
        /// <param name="request">Design-mode request with expected project revision and a fresh canonical destination path.</param>
        /// <param name="revalidateAuthorization">Optional current-authority check; true is used for discovery, false for delivery-sensitive checks.</param>
        /// <param name="recordClaim">Optional durable recorder invoked before each native mutation and at terminal completion.</param>
        /// <param name="requireNativeContext">Optional assertion that the caller's native project context is still current.</param>
        /// <returns>Task completing with mutation milestones and live-host verification; it never retries or rolls back uncertain work.</returns>
        internal Task<SolidWorksMacroCreationResult> CreateSolidWorksMacroAsync(Request request,
            Action<bool> revalidateAuthorization = null, Action<MacroMutationClaim> recordClaim = null, Action requireNativeContext = null)
        {
            if (request == null || request.ExpectedMode != 2)
                throw new ArgumentException("An explicit request in design mode is required.");
            if (solidWorksSavePending)
                throw new InvalidOperationException("An original SOLIDWORKS save is pending or uncertain; native creation cannot enter.");
            Action<bool> authorization = revalidateAuthorization ?? (_ => { });
            Action context = requireNativeContext ?? (() => { });
            int claimOrdinal = 0;
            Action<SolidWorksMacroCreationResult> durableClaim = result => {
                string phase = result.Terminal ? "Terminal" : result.SaveAttempts != 0 ? "BeforeSave" :
                    result.FilenameAttempts != 0 ? "BeforeFilename" : "BeforeNewMacro";
                var claim = new MacroMutationClaim(phase, result.HostPath, ++claimOrdinal);
                result.AddClaim(claim); recordClaim?.Invoke(claim);
            };
            string path = RequireFreshSolidWorksMacroPath(request.Path);
            var frozen = new Request { ExpectedMode = 2, ExpectedProjectVersion = request.ExpectedProjectVersion, Path = path };
            var scheduler = SolidWorksMacroCreationSchedulerFactory(); scheduler.RequireOwner();
            authorization(true);
            var before = RequireLifecycleCollection(frozen);
            Func<LifecycleProject, string> selector = row => string.IsNullOrWhiteSpace(row.Path) ? row.Name : row.Path;
            var originals = before.ToDictionary(row => row.Identity, row => (object)GetProject(selector(row)), StringComparer.OrdinalIgnoreCase);
            var native = SolidWorksMacroCreationNativeFactory((object)vbe, context);
            bool consumed = false;
            Action live = () => {
                scheduler.RequireOwner(); native.RequireOwner(); authorization(true);
                if (solidWorksSavePending) throw new InvalidOperationException("An original SOLIDWORKS save must settle before native creation.");
                RequireFreshSolidWorksMacroPath(path);
                var current = RequireLifecycleCollection(frozen);
                foreach (var row in current)
                    if (!originals.ContainsKey(row.Identity) || !native.SameProject(originals[row.Identity], (object)GetProject(selector(row))))
                        throw new InvalidOperationException("The original SOLIDWORKS collection identity changed.");
                authorization(false); native.RequireOwner();
            };
            Action final = () => { scheduler.RequireOwner(); authorization(false); native.RequireOwner(); if (solidWorksSavePending) throw new InvalidOperationException("An original SOLIDWORKS save is pending before delivery."); };
            Func<SolidWorksMacroCreationResult, SolidWorksMacroCreationResult> verify = result => {
                scheduler.RequireOwner(); native.RequireOwner(); authorization(true);
                var after = ReadLifecycleCollection(); RequireLifecycleDesign(after);
                var retained = after.Where(row => originals.ContainsKey(row.Identity)).ToList();
                var added = after.Where(row => !originals.ContainsKey(row.Identity)).ToList();
                if (after.Count != before.Count + 1 || added.Count != 1 || LifecycleVersion(retained) != LifecycleVersion(before))
                    throw new InvalidOperationException("Native creation did not preserve exactly the original collection.");
                foreach (var row in retained)
                    if (!native.SameProject(originals[row.Identity], (object)GetProject(selector(row))))
                        throw new InvalidOperationException("A preexisting canonical project was replaced.");
                var target = added[0];
                if (target.Type != 100 || target.Mode != 2 || target.Protection != 0 ||
                    !string.Equals(target.Path, path, StringComparison.OrdinalIgnoreCase) || !File.Exists(path) || new FileInfo(path).Length == 0)
                    throw new InvalidOperationException("The fresh native host project/path/file was not verified.");
                object canonical = (object)GetProject(path);
                dynamic project = canonical;
                if (!ReadNativeCreationComponents(canonical).Any(item => (int)((dynamic)item).Type == 100 && (string)((dynamic)item).Name == "ThisLibrary"))
                    throw new InvalidOperationException("The native SOLIDWORKS ThisLibrary document item is absent.");
                dynamic properties = ProjectProperties(path);
                if (!native.SameProject(canonical, (object)GetProject(path)) || (int)project.Mode != 2 || (int)project.Protection != 0 ||
                    !string.Equals(StandaloneAwareProjectPath(canonical), path, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Final destination identity/path/mode changed.");
                authorization(false); native.RequireOwner();
                result.Project = target.Name; result.ProjectVersion = (string)properties.Version;
                result.CollectionVersion = LifecycleVersion(after); result.DestinationCreated = true; result.Verified = true;
                return result;
            };
            try
            {
                live(); native.Prepare(); final();
                consumed = true;
                return RunSolidWorksMacroCreationAsync(path, native, scheduler, live, final, verify, durableClaim);
            }
            finally { if (!consumed) native.Dispose(); }
        }

        /// <summary>Reads the live VBComponents collection into name/type records for destination verification.</summary>
        /// <param name="nativeProject">Canonical project COM object returned by the host's project collection.</param>
        /// <returns>Snapshot records containing each component's current name and numeric type.</returns>
        private static IEnumerable<object> ReadNativeCreationComponents(object nativeProject)
        {
            dynamic project = nativeProject;
            var result = new List<object>();
            foreach (dynamic component in project.VBComponents)
                result.Add(new { Name = (string)component.Name, Type = (int)component.Type });
            return result;
        }

        /// <summary>Requires a canonical absolute .swp path whose parent exists and whose destination is absent.</summary>
        /// <param name="path">Destination path; its spelling must already equal <see cref="Path.GetFullPath(string)"/> output.</param>
        /// <returns>The unchanged path after extension, canonicality, parent-directory, and absence checks.</returns>
        internal static string RequireFreshSolidWorksMacroPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || path.IndexOf('\0') >= 0 || !Path.IsPathRooted(path) ||
                !string.Equals(Path.GetFullPath(path), path, StringComparison.Ordinal) ||
                !string.Equals(Path.GetExtension(path), ".swp", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("An explicit canonical absolute fresh .swp destination is required.");
            if (!Directory.Exists(Path.GetDirectoryName(path)) || File.Exists(path) || Directory.Exists(path))
                throw new IOException("The destination directory must exist and the destination must be absent.");
            return path;
        }

        /// <summary>Runs the one-shot command, filename write, Save, and close-verification sequence on the owner STA.</summary>
        /// <param name="path">Canonical destination carried into the result and native filename write.</param>
        /// <param name="native">Native adapter that validates and acts on the frozen Save dialog.</param>
        /// <param name="scheduler">Original owner-STA scheduler for command posting and bounded polling.</param>
        /// <param name="live">Revalidates current collection identity, authorization, and project context.</param>
        /// <param name="final">Checks authorization, owner identity, and pending-save state immediately before delivery.</param>
        /// <param name="verify">Final live-host collection and destination verifier run after the original command returns and dialog closes.</param>
        /// <param name="durableClaim">Records the current mutation milestone before each native side effect and at completion.</param>
        /// <returns>Task with terminal result; any uncertain native outcome ends the operation without retry or cleanup mutation.</returns>
        internal Task<SolidWorksMacroCreationResult> RunSolidWorksMacroCreationAsync(string path,
            ISolidWorksMacroCreationNative native, VbeProjectGeneralOperation.IScheduler scheduler,
            Action live, Action final, Func<SolidWorksMacroCreationResult, SolidWorksMacroCreationResult> verify,
            Action<SolidWorksMacroCreationResult> durableClaim)
        {
            var result = new SolidWorksMacroCreationResult { HostPath = path };
            var completion = new TaskCompletionSource<SolidWorksMacroCreationResult>();
            IDisposable polling = null; SolidWorksMacroDialog dialog = null; bool terminal = false, disposed = false, ticking = false;
            long began = scheduler.ElapsedMilliseconds;
            Action budget = () => { if (scheduler.ElapsedMilliseconds - began >= SolidWorksMacroCreationTimeoutMilliseconds)
                throw new TimeoutException("Native macro creation deadline expired; no retry or cleanup mutation."); };
            Action dispose = () => { if (!disposed && (!result.CommandEntered || result.OriginalCommandReturned)) { disposed = true; native.Dispose(); } };
            Action finish = () => {
                if (terminal) return;
                terminal = true; polling?.Dispose(); result.Terminal = true;
                try { dispose(); } catch (Exception error) { result.Error = (result.Error ?? "") + " | Native reference cleanup: " + error; result.Verified = false; result.Uncertain |= result.CommandEntered; }
                try { durableClaim(result); }
                catch (Exception error) { result.Error = "Terminal receipt: " + error; result.Verified = false; result.Uncertain |= result.CommandEntered; }
                completion.TrySetResult(result);
            };
            Action<Exception> fail = error => {
                result.Error = error.ToString(); result.Verified = false;
                result.Uncertain |= result.CommandEntered && (!result.OriginalCommandReturned || !result.DialogClosed || result.SaveQueued);
                finish();
            };
            Action tick = () => {
                if (terminal || ticking || !result.CommandEntered) return;
                ticking = true;
                try
                {
                    scheduler.RequireOwner(); native.RequireOwner(); budget();
                    if (dialog == null)
                    {
                        dialog = native.Capture();
                        if (dialog == null) { if (result.OriginalCommandReturned) throw new InvalidOperationException("Native command returned without its original Save dialog."); return; }
                        live(); native.RequireSame(dialog); final(); budget();
                        result.FilenameAttempts = 1; durableClaim(result);
                        native.WriteFilename(dialog, path, () => { live(); final(); budget(); },
                            () => { final(); budget(); result.FilenameWriteEntered = true; });
                        var after = native.Capture();
                        if (!dialog.Same(after) || after.FilenameText != path) throw new InvalidOperationException("Original filename readback/identity differs.");
                        dialog = after;
                        native.RequireSame(dialog); live(); final(); budget();
                        result.SaveAttempts = 1; durableClaim(result);
                        native.Save(dialog, () => { live(); final(); budget(); },
                            () => { final(); budget(); result.SaveQueued = true; });
                    }
                    else if (result.SaveQueued && native.Closed(dialog))
                    {
                        result.DialogClosed = true;
                        if (result.OriginalCommandReturned) { budget(); verify(result); finish(); }
                    }
                }
                catch (Exception error) { fail(error); }
                finally { ticking = false; }
            };
            try
            {
                polling = scheduler.Poll(tick); // Arm the original owner-STA timer before entering the modal command.
                scheduler.Post(() => {
                    if (terminal) return;
                    try
                    {
                        scheduler.RequireOwner(); live(); native.Prepare(); final(); budget();
                        result.CommandAttempts = 1; durableClaim(result);
                        native.Create(() => { live(); native.Prepare(); final(); budget(); result.CommandEntered = true; });
                        result.OriginalCommandReturned = true;
                        if (terminal) dispose(); else tick();
                    }
                    catch (Exception error) { if (terminal) { result.Error = (result.Error ?? "") + " | Original command: " + error; } else fail(error); }
                });
            }
            catch (Exception error) { fail(error); }
            return completion.Task;
        }

        /// <summary>Posts and polls creation work on the original SOLIDWORKS UI synchronization context.</summary>
        private sealed class SolidWorksMacroCreationScheduler : VbeProjectGeneralOperation.IScheduler
        {

            /// <summary>Synchronization context captured when creation starts.</summary>
            private readonly SynchronizationContext context = SynchronizationContext.Current;

            /// <summary>Managed thread ID required by this scheduler's owner-thread assertions.</summary>
            private readonly int owner = Thread.CurrentThread.ManagedThreadId;

            /// <summary>Monotonic elapsed-time source used for the operation deadline.</summary>
            private readonly Stopwatch watch = Stopwatch.StartNew();

            /// <summary>Gets the elapsed milliseconds.</summary>
            /// <value>Current elapsed milliseconds exposed by solid works macro creation scheduler.</value>
            public long ElapsedMilliseconds => watch.ElapsedMilliseconds;

            /// <summary>Requires the captured synchronization context, managed thread, and STA apartment.</summary>
            public void RequireOwner() { if (context == null || owner != Thread.CurrentThread.ManagedThreadId || Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                throw new InvalidOperationException("Native macro creation requires the original VBE UI STA/context."); }

            /// <summary>Posts one callback to the captured owner synchronization context.</summary>
            /// <param name="action">Callback to execute asynchronously on the owner thread.</param>
            public void Post(Action action) { RequireOwner(); context.Post(_ => action(), null); }

            /// <summary>Starts a 50 ms WinForms timer that polls the modal operation on the owner thread.</summary>
            /// <param name="action">One poll callback invoked by each timer tick.</param>
            /// <returns>Timer handle; disposing it stops further polling.</returns>
            public IDisposable Poll(Action action) { RequireOwner(); var timer = new System.Windows.Forms.Timer { Interval = 50 }; timer.Tick += (s, e) => action(); timer.Start(); return timer; }
        }

        /// <summary>Defines the i solid works macro application contract.</summary>
        [ComImport, Guid("83A33D22-27C5-11CE-BFD4-00400513BB57"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
        private interface ISolidWorksMacroApplication
        {

            /// <summary>Gets the process ID associated with this SOLIDWORKS automation object.</summary>
            /// <returns>Host process ID used to reject a ROT object from another instance.</returns>
            [DispId(166)] int GetProcessID();

            /// <summary>Gets the SOLIDWORKS revision string captured and rechecked around command entry.</summary>
            /// <returns>Host revision identifier.</returns>
            [DispId(12)] [return: MarshalAs(UnmanagedType.BStr)] string RevisionNumber();

            /// <summary>Invokes one host command by numeric command ID and title.</summary>
            /// <param name="command">SOLIDWORKS command identifier; this flow uses 573 for New Macro.</param>
            /// <param name="title">Command title parameter, empty for the New Macro route.</param>
            /// <returns>Host-reported command acceptance; false is surfaced as failure without retry.</returns>
            [DispId(245)] [return: MarshalAs(UnmanagedType.VariantBool)] bool RunCommand(int command, [MarshalAs(UnmanagedType.BStr)] string title);
        }

        /// <summary>Implements native SOLIDWORKS ROT binding and one-shot New Macro dialog mutations.</summary>
        private sealed class NativeSolidWorksMacroCreation : ISolidWorksMacroCreationNative
        {

            /// <summary>Current host process ID, used to bind the ROT automation object and dialog windows.</summary>
            private readonly int pid = Process.GetCurrentProcess().Id;

            /// <summary>Native UI thread captured when this adapter is constructed.</summary>
            private readonly uint thread = NativeThread();

            /// <summary>SOLIDWORKS main-window HWND used as the root of native ownership checks.</summary>
            private readonly IntPtr root;

            /// <summary>Caller-supplied assertion that the target project context remains authorized and current.</summary>
            private readonly Action requireContext;

            /// <summary>Unique RCW resolved from this process's exact SOLIDWORKS ROT moniker.</summary>
            private readonly object application;

            /// <summary>Typed dispatch interface used to check host identity and invoke command 573.</summary>
            private readonly ISolidWorksMacroApplication typed;

            /// <summary>Host revision captured at construction and rechecked before New Macro command entry.</summary>
            private readonly string revision;

            /// <summary>One-use guards for command, filename, Save, and COM cleanup; uncertain actions cannot be replayed.</summary>
            private bool commandConsumed, filenameConsumed, saveConsumed, disposed;

            /// <summary>Binds to the current SOLIDWORKS process, its main window, and exact in-process ROT application.</summary>
            /// <param name="editor">VBE editor whose MainWindow HWND identifies the current SOLIDWORKS host.</param>
            /// <param name="context">Callback that revalidates the native project context on each operation.</param>
            internal NativeSolidWorksMacroCreation(object editor, Action context)
            {
                requireContext = context ?? throw new ArgumentNullException(nameof(context));
                if (!string.Equals(Process.GetCurrentProcess().ProcessName, "SLDWORKS", StringComparison.OrdinalIgnoreCase))
                    throw new NotSupportedException("Native macro creation is available only in the current SOLIDWORKS process.");
                root = new IntPtr(Convert.ToInt64(((dynamic)editor).MainWindow.HWnd)); RequireOwner();
                application = ResolveOwnedSolidWorksApplication(pid);
                try { typed = (ISolidWorksMacroApplication)application; if (typed.GetProcessID() != pid) throw new InvalidOperationException("SOLIDWORKS application PID differs."); revision = typed.RevisionNumber(); }
                catch { Marshal.ReleaseComObject(application); throw; }
            }

            /// <summary>Requires the captured SOLIDWORKS process, window, UI thread, STA, and project context.</summary>
            public void RequireOwner()
            {
                uint owner; uint tid = WindowThread(root, out owner);
                if (disposed || root == IntPtr.Zero || owner != pid || tid != thread || NativeThread() != thread || Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                    throw new InvalidOperationException("Original SOLIDWORKS VBE owner/window/thread changed.");
                requireContext();
            }

            /// <summary>Compares project COM identity using the shared general-project identity rule.</summary>
            /// <param name="a">Previously captured canonical project COM object.</param>
            /// <param name="b">Current project COM object resolved after a host transition.</param>
            /// <returns>Whether both objects refer to the same underlying project.</returns>
            public bool SameProject(object a, object b) => SameGeneralProject(a, b);

            /// <summary>Checks owner identity, enabled host window, and absence of any visible host modal dialog.</summary>
            public void Prepare() { RequireOwner(); if (!Enabled(root) || VisibleDialogs().Count != 0) throw new InvalidOperationException("An original host modal is already present."); }

            /// <summary>Consumes the sole New Macro command attempt after host PID/revision and final checks pass.</summary>
            /// <param name="beforeEntry">Final caller checks run immediately before native command delivery.</param>
            public void Create(Action beforeEntry)
            {
                RequireOwner(); if (commandConsumed) throw new InvalidOperationException("Native New Macro is already consumed."); commandConsumed = true;
                if (typed.GetProcessID() != pid || typed.RevisionNumber() != revision) throw new InvalidOperationException("Original native application changed.");
                beforeEntry(); RequireOwner();
                if (!typed.RunCommand(573, "")) throw new InvalidOperationException("The single native New Macro returned false.");
            }

            /// <summary>Captures the sole visible modal only when it has the supported filename and Save controls.</summary>
            /// <returns>Verified dialog snapshot, or <see langword="null"/> when no host modal is present.</returns>
            public SolidWorksMacroDialog Capture()
            {
                RequireOwner(); var dialogs = VisibleDialogs(); if (dialogs.Count == 0) return null;
                if (dialogs.Count != 1) throw new InvalidOperationException("Owned visible modal is ambiguous.");
                IntPtr dialog = dialogs[0]; var filenames = new List<IntPtr>(); var saves = new List<IntPtr>(); bool complete = true; int count = 0; Exception callbackError = null;
                NativeCallback callback = (w, state) => {
                    try {
                    if (++count > 512) { complete = false; return false; }
                    RequireWindow(w); if (Visible(w) && Enabled(w)) {
                        string cls = Class(w); int id = ControlId(w);
                        if (cls == "Edit" && (id == 1001 || id == 1148) && Class(Parent(w)) == "ComboBox") filenames.Add(w);
                        if (cls == "Button" && id == 1) saves.Add(w);
                    } return true;
                    } catch (Exception error) { callbackError = error; complete = false; return false; }
                };
                EnumChildren(dialog, callback, IntPtr.Zero);
                if (callbackError != null) throw new InvalidOperationException("Native Save dialog inventory failed.", callbackError);
                if (!complete || filenames.Count != 1 || saves.Count != 1 || !SaveCaption(Text(saves[0])))
                    throw new InvalidOperationException("Only the exact original native Save dialog shape is supported.");
                return new SolidWorksMacroDialog { Window = dialog, Filename = filenames[0], SaveButton = saves[0], Thread = thread, FilenameText = Text(filenames[0]) };
            }

            /// <summary>Captures a fresh snapshot and rejects any change to the frozen dialog or child HWNDs.</summary>
            /// <param name="expected">Previously captured native dialog identity.</param>
            public void RequireSame(SolidWorksMacroDialog expected) { if (expected == null || !expected.Same(Capture())) throw new InvalidOperationException("Original Save dialog identity changed."); }

            /// <summary>Writes the destination path to the filename edit once using bounded SendMessageTimeout.</summary>
            /// <param name="expected">Frozen dialog snapshot whose HWNDs and control shape must still match.</param>
            /// <param name="path">Canonical destination path entered in the host's filename edit.</param>
            /// <param name="beforeEntry">Revalidation before consuming the filename-write phase.</param>
            /// <param name="beforeDelivery">Final authorization and identity check before the native text message.</param>
            public void WriteFilename(SolidWorksMacroDialog expected, string path, Action beforeEntry, Action beforeDelivery)
            {
                if (filenameConsumed) throw new InvalidOperationException("Filename delivery is already consumed."); filenameConsumed = true;
                RequireSame(expected); beforeEntry(); RequireOwner(); RequireFinalControl(expected, false); UIntPtr result;
                beforeDelivery(); RequireOwner();
                if (WriteText(expected.Filename, 12, UIntPtr.Zero, path, 0x23, 250, out result) == IntPtr.Zero || result == UIntPtr.Zero)
                    throw new InvalidOperationException("Original filename write is uncertain; no retry.");
            }

            /// <summary>Posts one BM_CLICK to the verified Save button after exact filename readback.</summary>
            /// <param name="expected">Frozen dialog snapshot with the exact requested filename.</param>
            /// <param name="beforeEntry">Revalidation before consuming the Save phase.</param>
            /// <param name="beforeDelivery">Final authorization and identity check before the button message.</param>
            public void Save(SolidWorksMacroDialog expected, Action beforeEntry, Action beforeDelivery)
            {
                if (saveConsumed) throw new InvalidOperationException("Save delivery is already consumed."); saveConsumed = true;
                RequireSame(expected); if (Text(expected.Filename) != expected.FilenameText || string.IsNullOrWhiteSpace(expected.FilenameText))
                    throw new InvalidOperationException("Filename differs from the exact requested path.");
                beforeEntry(); RequireOwner(); RequireFinalControl(expected, true);
                beforeDelivery(); RequireOwner();
                if (!Post(expected.SaveButton, 0xF5, UIntPtr.Zero, IntPtr.Zero)) throw new InvalidOperationException("Single Save enqueue failed; no retry.");
            }

            /// <summary>Checks whether the original dialog closed and the host remains enabled without a replacement modal.</summary>
            /// <param name="expected">Original dialog snapshot used to distinguish closure from replacement.</param>
            /// <returns>True only when the original HWND is gone and no visible modal dialog remains.</returns>
            public bool Closed(SolidWorksMacroDialog expected)
            {
                RequireOwner(); if (VisibleDialogs().Count != 0) { RequireSame(expected); return false; }
                return !IsWindow(expected.Window) && Enabled(root);
            }

            /// <summary>Enumerates the current desktop and returns visible modal windows owned by this host process.</summary>
            /// <returns>Complete list of owned visible dialog HWNDs; incomplete enumeration or a foreign owner thread throws.</returns>
            private List<IntPtr> VisibleDialogs()
            {
                RequireOwner(); var found = new List<IntPtr>(); bool valid = true, sawRoot = false; int count = 0; Exception callbackError = null;
                NativeCallback callback = (w, state) => {
                    try {
                    if (++count > 8192) { valid = false; return false; }
                    uint owner; uint tid = WindowThread(w, out owner); if (owner == 0 || tid == 0) { valid = false; return false; }
                    if (w == root) sawRoot = true;
                    if (owner == pid && Visible(w) && Class(w) == "#32770") { if (tid != thread) throw new InvalidOperationException("Owned modal has another UI thread."); found.Add(w); }
                    return true;
                    } catch (Exception error) { callbackError = error; valid = false; return false; }
                };
                bool complete = EnumDesktop(NativeDesktop(thread), callback, IntPtr.Zero);
                if (callbackError != null) throw new InvalidOperationException("Original desktop inventory failed.", callbackError);
                if (!complete || !valid || !sawRoot) throw new InvalidOperationException("Original desktop inventory is incomplete.");
                return found;
            }

            /// <summary>Requires a live HWND owned by the captured process and native UI thread.</summary>
            /// <param name="w">Window or control HWND to validate.</param>
            private void RequireWindow(IntPtr w) { uint p; if (WindowThread(w, out p) != thread || p != pid || !IsWindow(w)) throw new InvalidOperationException("Native dialog/control owner changed."); }

            /// <summary>Rechecks exact dialog, edit, and Save button classes, IDs, ancestry, visibility, and enabled state.</summary>
            /// <param name="expected">Frozen dialog identity and filename value.</param>
            /// <param name="save">When true, also requires the filename text to equal the frozen readback value.</param>
            private void RequireFinalControl(SolidWorksMacroDialog expected, bool save)
            {
                RequireWindow(expected.Window); RequireWindow(expected.Filename); RequireWindow(expected.SaveButton);
                if (Class(expected.Window) != "#32770" || !Visible(expected.Window) ||
                    Class(expected.Filename) != "Edit" || (ControlId(expected.Filename) != 1001 && ControlId(expected.Filename) != 1148) ||
                    Class(Parent(expected.Filename)) != "ComboBox" || Ancestor(expected.Filename, 2) != expected.Window ||
                    Class(expected.SaveButton) != "Button" || ControlId(expected.SaveButton) != 1 || Ancestor(expected.SaveButton, 2) != expected.Window ||
                    !Visible(expected.Filename) || !Enabled(expected.Filename) || !Visible(expected.SaveButton) || !Enabled(expected.SaveButton) ||
                    !SaveCaption(Text(expected.SaveButton)) || (save && Text(expected.Filename) != expected.FilenameText))
                    throw new InvalidOperationException("Original Save control shape/value changed before delivery.");
            }

            /// <summary>Reads the Unicode Win32 class name for a window.</summary>
            /// <param name="w">HWND whose class is requested.</param>
            /// <returns>Class name; a failed or empty native read throws.</returns>
            private static string Class(IntPtr w) { var text = new StringBuilder(256); if (GetClass(w, text, text.Capacity) == 0) throw new InvalidOperationException("Native class unavailable."); return text.ToString(); }

            /// <summary>Reads bounded Unicode control text using a 250 ms SendMessageTimeout call.</summary>
            /// <param name="w">Owner-verified HWND whose text is read.</param>
            /// <returns>Complete text; timeout or a full/truncated buffer throws.</returns>
            private string Text(IntPtr w) { RequireWindow(w); var text = new StringBuilder(32768); UIntPtr length; if (ReadText(w, 13, new UIntPtr((uint)text.Capacity), text, 0x23, 250, out length) == IntPtr.Zero || length.ToUInt64() >= (ulong)text.Capacity - 1) throw new InvalidOperationException("Bounded native text read incomplete."); return text.ToString(); }

            /// <summary>Recognizes the qualified English and French Save button captions.</summary>
            /// <param name="text">Captured button caption.</param>
            /// <returns>True for <c>Save</c>, <c>&amp;Save</c>, <c>Enregistrer</c>, or <c>&amp;Enregistrer</c>.</returns>
            private static bool SaveCaption(string text) => text == "&Save" || text == "Save" || text == "&Enregistrer" || text == "Enregistrer";

            /// <summary>Disposes  for native solid works macro creation.</summary>
            public void Dispose() { if (disposed) return; if (NativeThread() != thread || Thread.CurrentThread.GetApartmentState() != ApartmentState.STA) throw new InvalidOperationException("Native reference cleanup requires its original STA."); disposed = true; Marshal.ReleaseComObject(application); }

            /// <summary>Resolves the exact in-process SOLIDWORKS automation moniker without activating another host.</summary>
            /// <param name="pid">Current SOLIDWORKS process ID encoded in the expected ROT display name.</param>
            /// <returns>Unique COM object for <c>SolidWorks_PID_{pid}</c>; absence or inventory overflow throws.</returns>
            private static object ResolveOwnedSolidWorksApplication(int pid)
            {
                IRunningObjectTable rot = null; IBindCtx context = null; IEnumMoniker iterator = null;
                try {
                    Marshal.ThrowExceptionForHR(GetRot(0, out rot)); Marshal.ThrowExceptionForHR(CreateContext(0, out context)); rot.EnumRunning(out iterator);
                    var current = new IMoniker[1]; int count = 0;
                    while (iterator.Next(1, current, IntPtr.Zero) == 0) {
                        try {
                            if (++count > 4096) throw new InvalidOperationException("ROT inventory bound exceeded.");
                            string name; current[0].GetDisplayName(context, null, out name);
                            if (!string.Equals(name, "SolidWorks_PID_" + pid, StringComparison.OrdinalIgnoreCase)) continue;
                            object borrowed; rot.GetObject(current[0], out borrowed);
                            IntPtr identity = Marshal.GetIUnknownForObject(borrowed);
                            try { return Marshal.GetUniqueObjectForIUnknown(identity); } finally { Marshal.Release(identity); }
                        } finally { if (current[0] != null) Marshal.ReleaseComObject(current[0]); }
                    }
                    throw new InvalidOperationException("Exact in-process SOLIDWORKS ROT identity is unavailable; no activation attempted.");
                } finally { if (iterator != null) Marshal.ReleaseComObject(iterator); if (context != null) Marshal.ReleaseComObject(context); if (rot != null) Marshal.ReleaseComObject(rot); }
            }

            /// <summary>Callback signature used to enumerate desktop and dialog-child HWNDs.</summary>
            /// <param name="window">Current HWND supplied by the Win32 enumerator.</param>
            /// <param name="state">Opaque caller context pointer forwarded by Win32.</param>
            /// <returns>True to continue enumeration; false to stop.</returns>
            private delegate bool NativeCallback(IntPtr window, IntPtr state);

            /// <summary>Reads the calling native thread ID from Kernel32.</summary>
            /// <returns>Current Win32 thread ID.</returns>
            [DllImport("kernel32.dll", EntryPoint = "GetCurrentThreadId")] private static extern uint NativeThread();

            /// <summary>Reads the owning thread and process IDs for an HWND.</summary>
            /// <param name="window">Window whose owner identity is queried.</param>
            /// <param name="pid">Receives the owning process ID.</param>
            /// <returns>Owning native thread ID, or zero for an invalid window.</returns>
            [DllImport("user32.dll", EntryPoint = "GetWindowThreadProcessId")] private static extern uint WindowThread(IntPtr window, out uint pid);

            /// <summary>Gets the desktop opened by a native thread.</summary>
            /// <param name="thread">Native thread ID whose desktop is queried.</param>
            /// <returns>Desktop HDESK, or zero if the thread has no accessible desktop.</returns>
            [DllImport("user32.dll", EntryPoint = "GetThreadDesktop")] private static extern IntPtr NativeDesktop(uint thread);

            /// <summary>Enumerates windows on the captured host desktop.</summary>
            /// <param name="desktop">Desktop handle returned for the host's native thread.</param>
            /// <param name="callback">Callback invoked with each enumerated top-level HWND.</param>
            /// <param name="state">Opaque callback context.</param>
            /// <returns>True when enumeration completes; false on cancellation or API failure.</returns>
            [DllImport("user32.dll", EntryPoint = "EnumDesktopWindows", SetLastError = true)] private static extern bool EnumDesktop(IntPtr desktop, NativeCallback callback, IntPtr state);

            /// <summary>Enumerates descendants of one captured dialog.</summary>
            /// <param name="window">Dialog HWND whose child controls are enumerated.</param>
            /// <param name="callback">Callback invoked for each child HWND.</param>
            /// <param name="state">Opaque callback context.</param>
            /// <returns>Win32 enumeration result; caller also enforces a separate child-count bound.</returns>
            [DllImport("user32.dll", EntryPoint = "EnumChildWindows")] private static extern bool EnumChildren(IntPtr window, NativeCallback callback, IntPtr state);

            /// <summary>Copies a Unicode window-class name into a caller-provided buffer.</summary>
            /// <param name="window">HWND to inspect.</param>
            /// <param name="text">Destination buffer.</param>
            /// <param name="capacity">Buffer capacity in characters.</param>
            /// <returns>Characters copied, excluding the terminator; zero indicates failure.</returns>
            [DllImport("user32.dll", EntryPoint = "GetClassNameW", CharSet = CharSet.Unicode)] private static extern int GetClass(IntPtr window, StringBuilder text, int capacity);

            /// <summary>Tests whether an HWND has the WS_VISIBLE state.</summary>
            /// <param name="window">Window to inspect.</param>
            /// <returns>True when the window is visible.</returns>
            [DllImport("user32.dll", EntryPoint = "IsWindowVisible")] private static extern bool Visible(IntPtr window);

            /// <summary>Tests whether an HWND is enabled for interaction.</summary>
            /// <param name="window">Window to inspect.</param>
            /// <returns>True when the window is enabled.</returns>
            [DllImport("user32.dll", EntryPoint = "IsWindowEnabled")] private static extern bool Enabled(IntPtr window);

            /// <summary>Tests whether an HWND still identifies a live window.</summary>
            /// <param name="window">HWND to test.</param>
            /// <returns>True while the native window exists.</returns>
            [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr window);

            /// <summary>Reads a dialog child control's numeric ID.</summary>
            /// <param name="window">Child HWND to inspect.</param>
            /// <returns>Control ID assigned to the child window.</returns>
            [DllImport("user32.dll", EntryPoint = "GetDlgCtrlID")] private static extern int ControlId(IntPtr window);

            /// <summary>Reads the immediate parent HWND of a control.</summary>
            /// <param name="window">Child HWND whose parent is requested.</param>
            /// <returns>Parent HWND, or zero when there is no parent.</returns>
            [DllImport("user32.dll", EntryPoint = "GetParent")] private static extern IntPtr Parent(IntPtr window);

            /// <summary>Reads an ancestor HWND using the requested relationship flag.</summary>
            /// <param name="window">Starting HWND.</param>
            /// <param name="flags">GetAncestor selector; caller uses GA_ROOT (2) for dialog ancestry.</param>
            /// <returns>Ancestor HWND, or zero when none exists.</returns>
            [DllImport("user32.dll", EntryPoint = "GetAncestor")] private static extern IntPtr Ancestor(IntPtr window, uint flags);

            /// <summary>Sends a Unicode message with a bounded wait and writes any message result separately.</summary>
            /// <param name="window">Target HWND.</param>
            /// <param name="message">Message ID, such as WM_GETTEXT or LB_GETTEXT.</param>
            /// <param name="first">Message-specific WPARAM value.</param>
            /// <param name="text">Managed text buffer passed as LPARAM.</param>
            /// <param name="flags">SendMessageTimeout behavior flags.</param>
            /// <param name="timeout">Maximum wait in milliseconds; callers use 250 ms.</param>
            /// <param name="result">Receives the message-specific LRESULT.</param>
            /// <returns>Nonzero on successful delivery before timeout; zero on failure or timeout.</returns>
            [DllImport("user32.dll", EntryPoint = "SendMessageTimeoutW", CharSet = CharSet.Unicode, SetLastError = true)] private static extern IntPtr ReadText(IntPtr window, uint message, UIntPtr first, StringBuilder text, uint flags, uint timeout, out UIntPtr result);

            /// <summary>Sends a Unicode text-setting message with a bounded wait.</summary>
            /// <param name="window">Target edit-control HWND.</param>
            /// <param name="message">Message ID, WM_SETTEXT in this flow.</param>
            /// <param name="first">Message-specific WPARAM value.</param>
            /// <param name="text">Text to write to the filename control.</param>
            /// <param name="flags">SendMessageTimeout behavior flags.</param>
            /// <param name="timeout">Maximum wait in milliseconds; caller uses 250 ms.</param>
            /// <param name="result">Receives the control's message result.</param>
            /// <returns>Nonzero on successful delivery before timeout; zero on failure or timeout.</returns>
            [DllImport("user32.dll", EntryPoint = "SendMessageTimeoutW", CharSet = CharSet.Unicode, SetLastError = true)] private static extern IntPtr WriteText(IntPtr window, uint message, UIntPtr first, string text, uint flags, uint timeout, out UIntPtr result);

            /// <summary>Posts a native message asynchronously to the target window queue.</summary>
            /// <param name="window">Target HWND.</param>
            /// <param name="message">Win32 message ID, such as BM_CLICK.</param>
            /// <param name="first">Message-specific WPARAM value.</param>
            /// <param name="second">Message-specific LPARAM value.</param>
            /// <returns>True when the message was queued; false does not authorize a retry.</returns>
            [DllImport("user32.dll", EntryPoint = "PostMessageW", SetLastError = true)] private static extern bool Post(IntPtr window, uint message, UIntPtr first, IntPtr second);

            /// <summary>Gets the COM Running Object Table for exact in-process host resolution.</summary>
            /// <param name="reserved">Reserved parameter, which must be zero.</param>
            /// <param name="table">Receives the caller-owned ROT interface on success.</param>
            /// <returns>HRESULT from GetRunningObjectTable.</returns>
            [DllImport("ole32.dll", EntryPoint = "GetRunningObjectTable")] private static extern int GetRot(int reserved, out IRunningObjectTable table);

            /// <summary>Creates the COM bind context used to read ROT moniker display names.</summary>
            /// <param name="reserved">Reserved parameter, which must be zero.</param>
            /// <param name="context">Receives the caller-owned bind context on success.</param>
            /// <returns>int produced by the operation for create context on native solid works macro creation.</returns>
            [DllImport("ole32.dll", EntryPoint = "CreateBindCtx")] private static extern int CreateContext(int reserved, out IBindCtx context);
        }
    }
}
