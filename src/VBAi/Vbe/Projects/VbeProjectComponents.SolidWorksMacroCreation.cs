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

    /// <summary>Owns the vbe project components state and operations.</summary>
    internal sealed partial class VbeProjectComponents
    {

        /// <summary>Owns the solid works macro creation result state and operations.</summary>
        internal sealed class SolidWorksMacroCreationResult
        {

            /// <summary>Maintains the verified and destination created and uncertain and command entered and original command returned and filename write entered and save queued and dialog closed and terminal state for solid works macro creation result.</summary>
            public bool Verified, DestinationCreated, Uncertain, CommandEntered, OriginalCommandReturned,
                FilenameWriteEntered, SaveQueued, DialogClosed, Terminal;

            /// <summary>Maintains the command attempts and filename attempts and save attempts state for solid works macro creation result.</summary>
            public int CommandAttempts, FilenameAttempts, SaveAttempts;

            /// <summary>Keeps the project and host path and project version and collection version and error path available to solid works macro creation result.</summary>
            public string Project, HostPath, ProjectVersion, CollectionVersion, Error;

            /// <summary>Maintains the claims state for solid works macro creation result.</summary>
            private readonly List<MacroMutationClaim> claims = new List<MacroMutationClaim>();

            /// <summary>Gets the claims.</summary>
            /// <value>Current claims exposed by solid works macro creation result.</value>
            public MacroMutationClaim[] Claims => claims.ToArray();

            /// <summary>Adds claim for solid works macro creation result.</summary>
            /// <param name="claim">macro mutation claim that supplies the claim for this operation.</param>
            internal void AddClaim(MacroMutationClaim claim) { claims.Add(claim); }

            /// <summary>Gets the retry allowed.</summary>
            /// <value>Current retry allowed exposed by solid works macro creation result.</value>
            public bool RetryAllowed => false;

            /// <summary>Gets the rollback performed.</summary>
            /// <value>Current rollback performed exposed by solid works macro creation result.</value>
            public bool RollbackPerformed => false;

            /// <summary>Gets the persistence reload verified.</summary>
            /// <value>Current persistence reload verified exposed by solid works macro creation result.</value>
            public bool PersistenceReloadVerified => false;

            /// <summary>Gets the destination identity verified.</summary>
            /// <value>Current destination identity verified exposed by solid works macro creation result.</value>
            public bool DestinationIdentityVerified => DestinationCreated;

            /// <summary>Gets the destination absence proven.</summary>
            /// <value>Current destination absence proven exposed by solid works macro creation result.</value>
            public bool DestinationAbsenceProven => false;

            /// <summary>Gets the limit.</summary>
            /// <value>Current limit exposed by solid works macro creation result.</value>
            public string Limit => "DestinationCreated reports a verified native destination identity. An unverified partial file may exist at HostPath; inspect locally without automatic cleanup or retry. Fresh reopen is required to qualify persisted contents.";
        }

        /// <summary>Owns the macro mutation claim state and operations.</summary>
        internal sealed class MacroMutationClaim
        {

            /// <summary>Keeps the phase and host path path available to macro mutation claim.</summary>
            public readonly string Phase, HostPath;

            /// <summary>Maintains the ordinal state for macro mutation claim.</summary>
            public readonly int Ordinal;

            /// <summary>Maintains the utc state for macro mutation claim.</summary>
            public readonly DateTime Utc;

            /// <summary>Initializes a MacroMutationClaim instance with the supplied state.</summary>
            /// <param name="phase">Text that supplies the phase value. Use the format required by the calling operation.</param>
            /// <param name="path">Path used for the path being processed.</param>
            /// <param name="ordinal">int that supplies the ordinal for this operation.</param>
            internal MacroMutationClaim(string phase, string path, int ordinal)
            { Phase = phase; HostPath = path; Ordinal = ordinal; Utc = DateTime.UtcNow; }
        }

        /// <summary>Owns the solid works macro dialog state and operations.</summary>
        internal sealed class SolidWorksMacroDialog
        {

            /// <summary>Maintains the window and filename and save button state for solid works macro dialog.</summary>
            internal IntPtr Window, Filename, SaveButton;

            /// <summary>Maintains the thread state for solid works macro dialog.</summary>
            internal uint Thread;

            /// <summary>Maintains the filename text state for solid works macro dialog.</summary>
            internal string FilenameText;

            /// <summary>Compares  for solid works macro dialog.</summary>
            /// <param name="other">solid works macro dialog that supplies the other for this operation.</param>
            /// <returns>Boolean indicating the result of the check for same on solid works macro dialog.</returns>
            internal bool Same(SolidWorksMacroDialog other) => other != null && Window == other.Window &&
                Filename == other.Filename && SaveButton == other.SaveButton && Thread == other.Thread;
        }

        /// <summary>Defines the i solid works macro creation native contract.</summary>
        internal interface ISolidWorksMacroCreationNative : IDisposable
        {

            /// <summary>Requires owner for i solid works macro creation native.</summary>
            void RequireOwner();

            /// <summary>Handles prepare for i solid works macro creation native.</summary>
            void Prepare();

            /// <summary>Creates  for i solid works macro creation native.</summary>
            /// <param name="beforeEntry">action that supplies the before entry for this operation.</param>
            void Create(Action beforeEntry);

            /// <summary>Captures  for i solid works macro creation native.</summary>
            /// <returns>solid works macro dialog produced by the operation for capture on i solid works macro creation native.</returns>
            SolidWorksMacroDialog Capture();

            /// <summary>Requires same for i solid works macro creation native.</summary>
            /// <param name="expected">solid works macro dialog that supplies the expected for this operation.</param>
            void RequireSame(SolidWorksMacroDialog expected);

            /// <summary>Writes filename for i solid works macro creation native.</summary>
            /// <param name="expected">solid works macro dialog that supplies the expected for this operation.</param>
            /// <param name="path">Path used for the path being processed.</param>
            /// <param name="beforeEntry">action that supplies the before entry for this operation.</param>
            /// <param name="beforeDelivery">action that supplies the before delivery for this operation.</param>
            void WriteFilename(SolidWorksMacroDialog expected, string path, Action beforeEntry, Action beforeDelivery);

            /// <summary>Saves  for i solid works macro creation native.</summary>
            /// <param name="expected">solid works macro dialog that supplies the expected for this operation.</param>
            /// <param name="beforeEntry">action that supplies the before entry for this operation.</param>
            /// <param name="beforeDelivery">action that supplies the before delivery for this operation.</param>
            void Save(SolidWorksMacroDialog expected, Action beforeEntry, Action beforeDelivery);

            /// <summary>Closes d for i solid works macro creation native.</summary>
            /// <param name="expected">solid works macro dialog that supplies the expected for this operation.</param>
            /// <returns>Boolean indicating the result of the check for closed on i solid works macro creation native.</returns>
            bool Closed(SolidWorksMacroDialog expected);

            /// <summary>Compares project for i solid works macro creation native.</summary>
            /// <param name="first">object that supplies the first for this operation.</param>
            /// <param name="second">object that supplies the second for this operation.</param>
            /// <returns>Boolean indicating the result of the check for same project on i solid works macro creation native.</returns>
            bool SameProject(object first, object second);
        }

        /// <summary>Maintains the solid works macro creation native factory state for vbe project components.</summary>
        internal Func<object, Action, ISolidWorksMacroCreationNative> SolidWorksMacroCreationNativeFactory =
            (editor, context) => new NativeSolidWorksMacroCreation(editor, context);

        /// <summary>Maintains the solid works macro creation scheduler factory state for vbe project components.</summary>
        internal Func<VbeProjectGeneralOperation.IScheduler> SolidWorksMacroCreationSchedulerFactory =
            () => new SolidWorksMacroCreationScheduler();

        /// <summary>Maintains the solid works macro creation timeout milliseconds state for vbe project components.</summary>
        internal long SolidWorksMacroCreationTimeoutMilliseconds = 30000;

        // This is an explicit native creation route. It never calls Add101/SaveAs or retries a failed save.
        /// <summary>Creates solid works macro async for vbe project components.</summary>
        /// <param name="request">request that supplies the request for this operation.</param>
        /// <param name="revalidateAuthorization">action&lt;bool&gt; that supplies the revalidate authorization for this operation.</param>
        /// <param name="recordClaim">action&lt;macro mutation claim&gt; that supplies the record claim for this operation.</param>
        /// <param name="requireNativeContext">action that supplies the require native context for this operation.</param>
        /// <returns>task&lt;solid works macro creation result&gt; produced by the operation for create solid works macro async on vbe project components.</returns>
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

        /// <summary>Reads native creation components for vbe project components.</summary>
        /// <param name="nativeProject">object that supplies the native project for this operation.</param>
        /// <returns>i enumerable&lt;object&gt; produced by the operation for read native creation components on vbe project components.</returns>
        private static IEnumerable<object> ReadNativeCreationComponents(object nativeProject)
        {
            dynamic project = nativeProject;
            var result = new List<object>();
            foreach (dynamic component in project.VBComponents)
                result.Add(new { Name = (string)component.Name, Type = (int)component.Type });
            return result;
        }

        /// <summary>Requires fresh solid works macro path for vbe project components.</summary>
        /// <param name="path">Path used for the path being processed.</param>
        /// <returns>Text produced by the operation for require fresh solid works macro path on vbe project components.</returns>
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

        /// <summary>Runs solid works macro creation async for vbe project components.</summary>
        /// <param name="path">Path used for the path being processed.</param>
        /// <param name="native">i solid works macro creation native that supplies the native for this operation.</param>
        /// <param name="scheduler">i scheduler that supplies the scheduler for this operation.</param>
        /// <param name="live">action that supplies the live for this operation.</param>
        /// <param name="final">action that supplies the final for this operation.</param>
        /// <param name="verify">func&lt;solid works macro creation result, solid works macro creation result&gt; that supplies the verify for this operation.</param>
        /// <param name="durableClaim">action&lt;solid works macro creation result&gt; that supplies the durable claim for this operation.</param>
        /// <returns>task&lt;solid works macro creation result&gt; produced by the operation for run solid works macro creation async on vbe project components.</returns>
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

        /// <summary>Owns the solid works macro creation scheduler state and operations.</summary>
        private sealed class SolidWorksMacroCreationScheduler : VbeProjectGeneralOperation.IScheduler
        {

            /// <summary>Maintains the context state for solid works macro creation scheduler.</summary>
            private readonly SynchronizationContext context = SynchronizationContext.Current;

            /// <summary>Maintains the owner state for solid works macro creation scheduler.</summary>
            private readonly int owner = Thread.CurrentThread.ManagedThreadId;

            /// <summary>Maintains the watch state for solid works macro creation scheduler.</summary>
            private readonly Stopwatch watch = Stopwatch.StartNew();

            /// <summary>Gets the elapsed milliseconds.</summary>
            /// <value>Current elapsed milliseconds exposed by solid works macro creation scheduler.</value>
            public long ElapsedMilliseconds => watch.ElapsedMilliseconds;

            /// <summary>Requires owner for solid works macro creation scheduler.</summary>
            public void RequireOwner() { if (context == null || owner != Thread.CurrentThread.ManagedThreadId || Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                throw new InvalidOperationException("Native macro creation requires the original VBE UI STA/context."); }

            /// <summary>Handles post for solid works macro creation scheduler.</summary>
            /// <param name="action">action that supplies the action for this operation.</param>
            public void Post(Action action) { RequireOwner(); context.Post(_ => action(), null); }

            /// <summary>Handles poll for solid works macro creation scheduler.</summary>
            /// <param name="action">action that supplies the action for this operation.</param>
            /// <returns>i disposable produced by the operation for poll on solid works macro creation scheduler.</returns>
            public IDisposable Poll(Action action) { RequireOwner(); var timer = new System.Windows.Forms.Timer { Interval = 50 }; timer.Tick += (s, e) => action(); timer.Start(); return timer; }
        }

        /// <summary>Defines the i solid works macro application contract.</summary>
        [ComImport, Guid("83A33D22-27C5-11CE-BFD4-00400513BB57"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
        private interface ISolidWorksMacroApplication
        {

            /// <summary>Returns process id for i solid works macro application.</summary>
            /// <returns>int produced by the operation for get process id on i solid works macro application.</returns>
            [DispId(166)] int GetProcessID();

            /// <summary>Handles revision number for i solid works macro application.</summary>
            /// <returns>Text produced by the operation for revision number on i solid works macro application.</returns>
            [DispId(12)] [return: MarshalAs(UnmanagedType.BStr)] string RevisionNumber();

            /// <summary>Runs command for i solid works macro application.</summary>
            /// <param name="command">int that supplies the command for this operation.</param>
            /// <param name="title">Text that supplies the title value. Use the format required by the calling operation.</param>
            /// <returns>Boolean indicating the result of the check for run command on i solid works macro application.</returns>
            [DispId(245)] [return: MarshalAs(UnmanagedType.VariantBool)] bool RunCommand(int command, [MarshalAs(UnmanagedType.BStr)] string title);
        }

        /// <summary>Owns the native solid works macro creation state and operations.</summary>
        private sealed class NativeSolidWorksMacroCreation : ISolidWorksMacroCreationNative
        {

            /// <summary>Identifies the pid associated with native solid works macro creation.</summary>
            private readonly int pid = Process.GetCurrentProcess().Id;

            /// <summary>Maintains the thread state for native solid works macro creation.</summary>
            private readonly uint thread = NativeThread();

            /// <summary>Maintains the root state for native solid works macro creation.</summary>
            private readonly IntPtr root;

            /// <summary>Maintains the require context state for native solid works macro creation.</summary>
            private readonly Action requireContext;

            /// <summary>Maintains the application state for native solid works macro creation.</summary>
            private readonly object application;

            /// <summary>Maintains the typed state for native solid works macro creation.</summary>
            private readonly ISolidWorksMacroApplication typed;

            /// <summary>Maintains the revision state for native solid works macro creation.</summary>
            private readonly string revision;

            /// <summary>Maintains the command consumed and filename consumed and save consumed and disposed state for native solid works macro creation.</summary>
            private bool commandConsumed, filenameConsumed, saveConsumed, disposed;

            /// <summary>Initializes a NativeSolidWorksMacroCreation instance with the supplied state.</summary>
            /// <param name="editor">object that supplies the editor for this operation.</param>
            /// <param name="context">action that supplies the context for this operation.</param>
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

            /// <summary>Requires owner for native solid works macro creation.</summary>
            public void RequireOwner()
            {
                uint owner; uint tid = WindowThread(root, out owner);
                if (disposed || root == IntPtr.Zero || owner != pid || tid != thread || NativeThread() != thread || Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                    throw new InvalidOperationException("Original SOLIDWORKS VBE owner/window/thread changed.");
                requireContext();
            }

            /// <summary>Compares project for native solid works macro creation.</summary>
            /// <param name="a">object that supplies the a for this operation.</param>
            /// <param name="b">object that supplies the b for this operation.</param>
            /// <returns>Boolean indicating the result of the check for same project on native solid works macro creation.</returns>
            public bool SameProject(object a, object b) => SameGeneralProject(a, b);

            /// <summary>Handles prepare for native solid works macro creation.</summary>
            public void Prepare() { RequireOwner(); if (!Enabled(root) || VisibleDialogs().Count != 0) throw new InvalidOperationException("An original host modal is already present."); }

            /// <summary>Creates  for native solid works macro creation.</summary>
            /// <param name="beforeEntry">action that supplies the before entry for this operation.</param>
            public void Create(Action beforeEntry)
            {
                RequireOwner(); if (commandConsumed) throw new InvalidOperationException("Native New Macro is already consumed."); commandConsumed = true;
                if (typed.GetProcessID() != pid || typed.RevisionNumber() != revision) throw new InvalidOperationException("Original native application changed.");
                beforeEntry(); RequireOwner();
                if (!typed.RunCommand(573, "")) throw new InvalidOperationException("The single native New Macro returned false.");
            }

            /// <summary>Captures  for native solid works macro creation.</summary>
            /// <returns>solid works macro dialog produced by the operation for capture on native solid works macro creation.</returns>
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

            /// <summary>Requires same for native solid works macro creation.</summary>
            /// <param name="expected">solid works macro dialog that supplies the expected for this operation.</param>
            public void RequireSame(SolidWorksMacroDialog expected) { if (expected == null || !expected.Same(Capture())) throw new InvalidOperationException("Original Save dialog identity changed."); }

            /// <summary>Writes filename for native solid works macro creation.</summary>
            /// <param name="expected">solid works macro dialog that supplies the expected for this operation.</param>
            /// <param name="path">Path used for the path being processed.</param>
            /// <param name="beforeEntry">action that supplies the before entry for this operation.</param>
            /// <param name="beforeDelivery">action that supplies the before delivery for this operation.</param>
            public void WriteFilename(SolidWorksMacroDialog expected, string path, Action beforeEntry, Action beforeDelivery)
            {
                if (filenameConsumed) throw new InvalidOperationException("Filename delivery is already consumed."); filenameConsumed = true;
                RequireSame(expected); beforeEntry(); RequireOwner(); RequireFinalControl(expected, false); UIntPtr result;
                beforeDelivery(); RequireOwner();
                if (WriteText(expected.Filename, 12, UIntPtr.Zero, path, 0x23, 250, out result) == IntPtr.Zero || result == UIntPtr.Zero)
                    throw new InvalidOperationException("Original filename write is uncertain; no retry.");
            }

            /// <summary>Saves  for native solid works macro creation.</summary>
            /// <param name="expected">solid works macro dialog that supplies the expected for this operation.</param>
            /// <param name="beforeEntry">action that supplies the before entry for this operation.</param>
            /// <param name="beforeDelivery">action that supplies the before delivery for this operation.</param>
            public void Save(SolidWorksMacroDialog expected, Action beforeEntry, Action beforeDelivery)
            {
                if (saveConsumed) throw new InvalidOperationException("Save delivery is already consumed."); saveConsumed = true;
                RequireSame(expected); if (Text(expected.Filename) != expected.FilenameText || string.IsNullOrWhiteSpace(expected.FilenameText))
                    throw new InvalidOperationException("Filename differs from the exact requested path.");
                beforeEntry(); RequireOwner(); RequireFinalControl(expected, true);
                beforeDelivery(); RequireOwner();
                if (!Post(expected.SaveButton, 0xF5, UIntPtr.Zero, IntPtr.Zero)) throw new InvalidOperationException("Single Save enqueue failed; no retry.");
            }

            /// <summary>Closes d for native solid works macro creation.</summary>
            /// <param name="expected">solid works macro dialog that supplies the expected for this operation.</param>
            /// <returns>Boolean indicating the result of the check for closed on native solid works macro creation.</returns>
            public bool Closed(SolidWorksMacroDialog expected)
            {
                RequireOwner(); if (VisibleDialogs().Count != 0) { RequireSame(expected); return false; }
                return !IsWindow(expected.Window) && Enabled(root);
            }

            /// <summary>Handles visible dialogs for native solid works macro creation.</summary>
            /// <returns>list&lt;int ptr&gt; produced by the operation for visible dialogs on native solid works macro creation.</returns>
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

            /// <summary>Requires window for native solid works macro creation.</summary>
            /// <param name="w">Native handle that supplies the w for this operation.</param>
            private void RequireWindow(IntPtr w) { uint p; if (WindowThread(w, out p) != thread || p != pid || !IsWindow(w)) throw new InvalidOperationException("Native dialog/control owner changed."); }

            /// <summary>Requires final control for native solid works macro creation.</summary>
            /// <param name="expected">solid works macro dialog that supplies the expected for this operation.</param>
            /// <param name="save">Indicates whether save is enabled.</param>
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

            /// <summary>Handles class for native solid works macro creation.</summary>
            /// <param name="w">Native handle that supplies the w for this operation.</param>
            /// <returns>Text produced by the operation for class on native solid works macro creation.</returns>
            private static string Class(IntPtr w) { var text = new StringBuilder(256); if (GetClass(w, text, text.Capacity) == 0) throw new InvalidOperationException("Native class unavailable."); return text.ToString(); }

            /// <summary>Handles text for native solid works macro creation.</summary>
            /// <param name="w">Native handle that supplies the w for this operation.</param>
            /// <returns>Text produced by the operation for text on native solid works macro creation.</returns>
            private string Text(IntPtr w) { RequireWindow(w); var text = new StringBuilder(32768); UIntPtr length; if (ReadText(w, 13, new UIntPtr((uint)text.Capacity), text, 0x23, 250, out length) == IntPtr.Zero || length.ToUInt64() >= (ulong)text.Capacity - 1) throw new InvalidOperationException("Bounded native text read incomplete."); return text.ToString(); }

            /// <summary>Saves caption for native solid works macro creation.</summary>
            /// <param name="text">Text that supplies the text value. Use the format required by the calling operation.</param>
            /// <returns>Boolean indicating the result of the check for save caption on native solid works macro creation.</returns>
            private static bool SaveCaption(string text) => text == "&Save" || text == "Save" || text == "&Enregistrer" || text == "Enregistrer";

            /// <summary>Disposes  for native solid works macro creation.</summary>
            public void Dispose() { if (disposed) return; if (NativeThread() != thread || Thread.CurrentThread.GetApartmentState() != ApartmentState.STA) throw new InvalidOperationException("Native reference cleanup requires its original STA."); disposed = true; Marshal.ReleaseComObject(application); }

            /// <summary>Resolves owned solid works application for native solid works macro creation.</summary>
            /// <param name="pid">int that supplies the pid for this operation.</param>
            /// <returns>object produced by the operation for resolve owned solid works application on native solid works macro creation.</returns>
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

            /// <summary>Defines the native callback callback.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <param name="state">Native handle that supplies the state for this operation.</param>
            /// <returns>Boolean indicating the result of the check for operation on native solid works macro creation.</returns>
            private delegate bool NativeCallback(IntPtr window, IntPtr state);

            /// <summary>Handles native thread for native solid works macro creation.</summary>
            /// <returns>uint produced by the operation for native thread on native solid works macro creation.</returns>
            [DllImport("kernel32.dll", EntryPoint = "GetCurrentThreadId")] private static extern uint NativeThread();

            /// <summary>Handles window thread for native solid works macro creation.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <param name="pid">uint that supplies the pid for this operation.</param>
            /// <returns>uint produced by the operation for window thread on native solid works macro creation.</returns>
            [DllImport("user32.dll", EntryPoint = "GetWindowThreadProcessId")] private static extern uint WindowThread(IntPtr window, out uint pid);

            /// <summary>Handles native desktop for native solid works macro creation.</summary>
            /// <param name="thread">uint that supplies the thread for this operation.</param>
            /// <returns>int ptr produced by the operation for native desktop on native solid works macro creation.</returns>
            [DllImport("user32.dll", EntryPoint = "GetThreadDesktop")] private static extern IntPtr NativeDesktop(uint thread);

            /// <summary>Handles enum desktop for native solid works macro creation.</summary>
            /// <param name="desktop">Native handle that supplies the desktop for this operation.</param>
            /// <param name="callback">native callback that supplies the callback for this operation.</param>
            /// <param name="state">Native handle that supplies the state for this operation.</param>
            /// <returns>Boolean indicating the result of the check for enum desktop on native solid works macro creation.</returns>
            [DllImport("user32.dll", EntryPoint = "EnumDesktopWindows", SetLastError = true)] private static extern bool EnumDesktop(IntPtr desktop, NativeCallback callback, IntPtr state);

            /// <summary>Handles enum children for native solid works macro creation.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <param name="callback">native callback that supplies the callback for this operation.</param>
            /// <param name="state">Native handle that supplies the state for this operation.</param>
            /// <returns>Boolean indicating the result of the check for enum children on native solid works macro creation.</returns>
            [DllImport("user32.dll", EntryPoint = "EnumChildWindows")] private static extern bool EnumChildren(IntPtr window, NativeCallback callback, IntPtr state);

            /// <summary>Returns class for native solid works macro creation.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <param name="text">string builder that supplies the text for this operation.</param>
            /// <param name="capacity">int that supplies the capacity for this operation.</param>
            /// <returns>int produced by the operation for get class on native solid works macro creation.</returns>
            [DllImport("user32.dll", EntryPoint = "GetClassNameW", CharSet = CharSet.Unicode)] private static extern int GetClass(IntPtr window, StringBuilder text, int capacity);

            /// <summary>Handles visible for native solid works macro creation.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <returns>Boolean indicating the result of the check for visible on native solid works macro creation.</returns>
            [DllImport("user32.dll", EntryPoint = "IsWindowVisible")] private static extern bool Visible(IntPtr window);

            /// <summary>Handles enabled for native solid works macro creation.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <returns>Boolean indicating the result of the check for enabled on native solid works macro creation.</returns>
            [DllImport("user32.dll", EntryPoint = "IsWindowEnabled")] private static extern bool Enabled(IntPtr window);

            /// <summary>Determines whether window for native solid works macro creation.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <returns>Boolean indicating the result of the check for is window on native solid works macro creation.</returns>
            [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr window);

            /// <summary>Handles control id for native solid works macro creation.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <returns>int produced by the operation for control id on native solid works macro creation.</returns>
            [DllImport("user32.dll", EntryPoint = "GetDlgCtrlID")] private static extern int ControlId(IntPtr window);

            /// <summary>Handles parent for native solid works macro creation.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <returns>int ptr produced by the operation for parent on native solid works macro creation.</returns>
            [DllImport("user32.dll", EntryPoint = "GetParent")] private static extern IntPtr Parent(IntPtr window);

            /// <summary>Handles ancestor for native solid works macro creation.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <param name="flags">uint that supplies the flags for this operation.</param>
            /// <returns>int ptr produced by the operation for ancestor on native solid works macro creation.</returns>
            [DllImport("user32.dll", EntryPoint = "GetAncestor")] private static extern IntPtr Ancestor(IntPtr window, uint flags);

            /// <summary>Reads text for native solid works macro creation.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <param name="message">uint that supplies the message for this operation.</param>
            /// <param name="first">Native handle that supplies the first for this operation.</param>
            /// <param name="text">string builder that supplies the text for this operation.</param>
            /// <param name="flags">uint that supplies the flags for this operation.</param>
            /// <param name="timeout">uint that supplies the timeout for this operation.</param>
            /// <param name="result">Native handle that supplies the result for this operation.</param>
            /// <returns>int ptr produced by the operation for read text on native solid works macro creation.</returns>
            [DllImport("user32.dll", EntryPoint = "SendMessageTimeoutW", CharSet = CharSet.Unicode, SetLastError = true)] private static extern IntPtr ReadText(IntPtr window, uint message, UIntPtr first, StringBuilder text, uint flags, uint timeout, out UIntPtr result);

            /// <summary>Writes text for native solid works macro creation.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <param name="message">uint that supplies the message for this operation.</param>
            /// <param name="first">Native handle that supplies the first for this operation.</param>
            /// <param name="text">Text that supplies the text value. Use the format required by the calling operation.</param>
            /// <param name="flags">uint that supplies the flags for this operation.</param>
            /// <param name="timeout">uint that supplies the timeout for this operation.</param>
            /// <param name="result">Native handle that supplies the result for this operation.</param>
            /// <returns>int ptr produced by the operation for write text on native solid works macro creation.</returns>
            [DllImport("user32.dll", EntryPoint = "SendMessageTimeoutW", CharSet = CharSet.Unicode, SetLastError = true)] private static extern IntPtr WriteText(IntPtr window, uint message, UIntPtr first, string text, uint flags, uint timeout, out UIntPtr result);

            /// <summary>Handles post for native solid works macro creation.</summary>
            /// <param name="window">Native handle that supplies the window for this operation.</param>
            /// <param name="message">uint that supplies the message for this operation.</param>
            /// <param name="first">Native handle that supplies the first for this operation.</param>
            /// <param name="second">Native handle that supplies the second for this operation.</param>
            /// <returns>Boolean indicating the result of the check for post on native solid works macro creation.</returns>
            [DllImport("user32.dll", EntryPoint = "PostMessageW", SetLastError = true)] private static extern bool Post(IntPtr window, uint message, UIntPtr first, IntPtr second);

            /// <summary>Returns rot for native solid works macro creation.</summary>
            /// <param name="reserved">int that supplies the reserved for this operation.</param>
            /// <param name="table">i running object table that supplies the table for this operation.</param>
            /// <returns>int produced by the operation for get rot on native solid works macro creation.</returns>
            [DllImport("ole32.dll", EntryPoint = "GetRunningObjectTable")] private static extern int GetRot(int reserved, out IRunningObjectTable table);

            /// <summary>Creates context for native solid works macro creation.</summary>
            /// <param name="reserved">int that supplies the reserved for this operation.</param>
            /// <param name="context">i bind ctx that supplies the context for this operation.</param>
            /// <returns>int produced by the operation for create context on native solid works macro creation.</returns>
            [DllImport("ole32.dll", EntryPoint = "CreateBindCtx")] private static extern int CreateContext(int reserved, out IBindCtx context);
        }
    }
}
