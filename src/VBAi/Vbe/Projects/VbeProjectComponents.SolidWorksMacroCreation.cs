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
    internal sealed partial class VbeProjectComponents
    {
        internal sealed class SolidWorksMacroCreationResult
        {
            public bool Verified, DestinationCreated, Uncertain, CommandEntered, OriginalCommandReturned,
                FilenameWriteEntered, SaveQueued, DialogClosed, Terminal;
            public int CommandAttempts, FilenameAttempts, SaveAttempts;
            public string Project, HostPath, ProjectVersion, CollectionVersion, Error;
            private readonly List<MacroMutationClaim> claims = new List<MacroMutationClaim>();
            public MacroMutationClaim[] Claims => claims.ToArray();
            internal void AddClaim(MacroMutationClaim claim) { claims.Add(claim); }
            public bool RetryAllowed => false;
            public bool RollbackPerformed => false;
            public bool PersistenceReloadVerified => false;
            public bool DestinationIdentityVerified => DestinationCreated;
            public bool DestinationAbsenceProven => false;
            public string Limit => "DestinationCreated reports a verified native destination identity. An unverified partial file may exist at HostPath; inspect locally without automatic cleanup or retry. Fresh reopen is required to qualify persisted contents.";
        }

        internal sealed class MacroMutationClaim
        {
            public readonly string Phase, HostPath;
            public readonly int Ordinal;
            public readonly DateTime Utc;
            internal MacroMutationClaim(string phase, string path, int ordinal)
            { Phase = phase; HostPath = path; Ordinal = ordinal; Utc = DateTime.UtcNow; }
        }

        internal sealed class SolidWorksMacroDialog
        {
            internal IntPtr Window, Filename, SaveButton;
            internal uint Thread;
            internal string FilenameText;
            internal bool Same(SolidWorksMacroDialog other) => other != null && Window == other.Window &&
                Filename == other.Filename && SaveButton == other.SaveButton && Thread == other.Thread;
        }

        internal interface ISolidWorksMacroCreationNative : IDisposable
        {
            void RequireOwner();
            void Prepare();
            void Create(Action beforeEntry);
            SolidWorksMacroDialog Capture();
            void RequireSame(SolidWorksMacroDialog expected);
            void WriteFilename(SolidWorksMacroDialog expected, string path, Action beforeEntry, Action beforeDelivery);
            void Save(SolidWorksMacroDialog expected, Action beforeEntry, Action beforeDelivery);
            bool Closed(SolidWorksMacroDialog expected);
            bool SameProject(object first, object second);
        }

        internal Func<object, Action, ISolidWorksMacroCreationNative> SolidWorksMacroCreationNativeFactory =
            (editor, context) => new NativeSolidWorksMacroCreation(editor, context);
        internal Func<VbeProjectGeneralOperation.IScheduler> SolidWorksMacroCreationSchedulerFactory =
            () => new SolidWorksMacroCreationScheduler();
        internal long SolidWorksMacroCreationTimeoutMilliseconds = 30000;

        // This is an explicit native creation route. It never calls Add101/SaveAs or retries a failed save.
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

        private static IEnumerable<object> ReadNativeCreationComponents(object nativeProject)
        {
            dynamic project = nativeProject;
            var result = new List<object>();
            foreach (dynamic component in project.VBComponents)
                result.Add(new { Name = (string)component.Name, Type = (int)component.Type });
            return result;
        }

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

        private sealed class SolidWorksMacroCreationScheduler : VbeProjectGeneralOperation.IScheduler
        {
            private readonly SynchronizationContext context = SynchronizationContext.Current;
            private readonly int owner = Thread.CurrentThread.ManagedThreadId;
            private readonly Stopwatch watch = Stopwatch.StartNew();
            public long ElapsedMilliseconds => watch.ElapsedMilliseconds;
            public void RequireOwner() { if (context == null || owner != Thread.CurrentThread.ManagedThreadId || Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                throw new InvalidOperationException("Native macro creation requires the original VBE UI STA/context."); }
            public void Post(Action action) { RequireOwner(); context.Post(_ => action(), null); }
            public IDisposable Poll(Action action) { RequireOwner(); var timer = new System.Windows.Forms.Timer { Interval = 50 }; timer.Tick += (s, e) => action(); timer.Start(); return timer; }
        }

        [ComImport, Guid("83A33D22-27C5-11CE-BFD4-00400513BB57"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
        private interface ISolidWorksMacroApplication
        {
            [DispId(166)] int GetProcessID();
            [DispId(12)] [return: MarshalAs(UnmanagedType.BStr)] string RevisionNumber();
            [DispId(245)] [return: MarshalAs(UnmanagedType.VariantBool)] bool RunCommand(int command, [MarshalAs(UnmanagedType.BStr)] string title);
        }

        private sealed class NativeSolidWorksMacroCreation : ISolidWorksMacroCreationNative
        {
            private readonly int pid = Process.GetCurrentProcess().Id;
            private readonly uint thread = NativeThread();
            private readonly IntPtr root;
            private readonly Action requireContext;
            private readonly object application;
            private readonly ISolidWorksMacroApplication typed;
            private readonly string revision;
            private bool commandConsumed, filenameConsumed, saveConsumed, disposed;
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
            public void RequireOwner()
            {
                uint owner; uint tid = WindowThread(root, out owner);
                if (disposed || root == IntPtr.Zero || owner != pid || tid != thread || NativeThread() != thread || Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                    throw new InvalidOperationException("Original SOLIDWORKS VBE owner/window/thread changed.");
                requireContext();
            }
            public bool SameProject(object a, object b) => SameGeneralProject(a, b);
            public void Prepare() { RequireOwner(); if (!Enabled(root) || VisibleDialogs().Count != 0) throw new InvalidOperationException("An original host modal is already present."); }
            public void Create(Action beforeEntry)
            {
                RequireOwner(); if (commandConsumed) throw new InvalidOperationException("Native New Macro is already consumed."); commandConsumed = true;
                if (typed.GetProcessID() != pid || typed.RevisionNumber() != revision) throw new InvalidOperationException("Original native application changed.");
                beforeEntry(); RequireOwner();
                if (!typed.RunCommand(573, "")) throw new InvalidOperationException("The single native New Macro returned false.");
            }
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
            public void RequireSame(SolidWorksMacroDialog expected) { if (expected == null || !expected.Same(Capture())) throw new InvalidOperationException("Original Save dialog identity changed."); }
            public void WriteFilename(SolidWorksMacroDialog expected, string path, Action beforeEntry, Action beforeDelivery)
            {
                if (filenameConsumed) throw new InvalidOperationException("Filename delivery is already consumed."); filenameConsumed = true;
                RequireSame(expected); beforeEntry(); RequireOwner(); RequireFinalControl(expected, false); UIntPtr result;
                beforeDelivery(); RequireOwner();
                if (WriteText(expected.Filename, 12, UIntPtr.Zero, path, 0x23, 250, out result) == IntPtr.Zero || result == UIntPtr.Zero)
                    throw new InvalidOperationException("Original filename write is uncertain; no retry.");
            }
            public void Save(SolidWorksMacroDialog expected, Action beforeEntry, Action beforeDelivery)
            {
                if (saveConsumed) throw new InvalidOperationException("Save delivery is already consumed."); saveConsumed = true;
                RequireSame(expected); if (Text(expected.Filename) != expected.FilenameText || string.IsNullOrWhiteSpace(expected.FilenameText))
                    throw new InvalidOperationException("Filename differs from the exact requested path.");
                beforeEntry(); RequireOwner(); RequireFinalControl(expected, true);
                beforeDelivery(); RequireOwner();
                if (!Post(expected.SaveButton, 0xF5, UIntPtr.Zero, IntPtr.Zero)) throw new InvalidOperationException("Single Save enqueue failed; no retry.");
            }
            public bool Closed(SolidWorksMacroDialog expected)
            {
                RequireOwner(); if (VisibleDialogs().Count != 0) { RequireSame(expected); return false; }
                return !IsWindow(expected.Window) && Enabled(root);
            }
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
            private void RequireWindow(IntPtr w) { uint p; if (WindowThread(w, out p) != thread || p != pid || !IsWindow(w)) throw new InvalidOperationException("Native dialog/control owner changed."); }
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
            private static string Class(IntPtr w) { var text = new StringBuilder(256); if (GetClass(w, text, text.Capacity) == 0) throw new InvalidOperationException("Native class unavailable."); return text.ToString(); }
            private string Text(IntPtr w) { RequireWindow(w); var text = new StringBuilder(32768); UIntPtr length; if (ReadText(w, 13, new UIntPtr((uint)text.Capacity), text, 0x23, 250, out length) == IntPtr.Zero || length.ToUInt64() >= (ulong)text.Capacity - 1) throw new InvalidOperationException("Bounded native text read incomplete."); return text.ToString(); }
            private static bool SaveCaption(string text) => text == "&Save" || text == "Save" || text == "&Enregistrer" || text == "Enregistrer";
            public void Dispose() { if (disposed) return; if (NativeThread() != thread || Thread.CurrentThread.GetApartmentState() != ApartmentState.STA) throw new InvalidOperationException("Native reference cleanup requires its original STA."); disposed = true; Marshal.ReleaseComObject(application); }

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
            private delegate bool NativeCallback(IntPtr window, IntPtr state);
            [DllImport("kernel32.dll", EntryPoint = "GetCurrentThreadId")] private static extern uint NativeThread();
            [DllImport("user32.dll", EntryPoint = "GetWindowThreadProcessId")] private static extern uint WindowThread(IntPtr window, out uint pid);
            [DllImport("user32.dll", EntryPoint = "GetThreadDesktop")] private static extern IntPtr NativeDesktop(uint thread);
            [DllImport("user32.dll", EntryPoint = "EnumDesktopWindows", SetLastError = true)] private static extern bool EnumDesktop(IntPtr desktop, NativeCallback callback, IntPtr state);
            [DllImport("user32.dll", EntryPoint = "EnumChildWindows")] private static extern bool EnumChildren(IntPtr window, NativeCallback callback, IntPtr state);
            [DllImport("user32.dll", EntryPoint = "GetClassNameW", CharSet = CharSet.Unicode)] private static extern int GetClass(IntPtr window, StringBuilder text, int capacity);
            [DllImport("user32.dll", EntryPoint = "IsWindowVisible")] private static extern bool Visible(IntPtr window);
            [DllImport("user32.dll", EntryPoint = "IsWindowEnabled")] private static extern bool Enabled(IntPtr window);
            [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr window);
            [DllImport("user32.dll", EntryPoint = "GetDlgCtrlID")] private static extern int ControlId(IntPtr window);
            [DllImport("user32.dll", EntryPoint = "GetParent")] private static extern IntPtr Parent(IntPtr window);
            [DllImport("user32.dll", EntryPoint = "GetAncestor")] private static extern IntPtr Ancestor(IntPtr window, uint flags);
            [DllImport("user32.dll", EntryPoint = "SendMessageTimeoutW", CharSet = CharSet.Unicode, SetLastError = true)] private static extern IntPtr ReadText(IntPtr window, uint message, UIntPtr first, StringBuilder text, uint flags, uint timeout, out UIntPtr result);
            [DllImport("user32.dll", EntryPoint = "SendMessageTimeoutW", CharSet = CharSet.Unicode, SetLastError = true)] private static extern IntPtr WriteText(IntPtr window, uint message, UIntPtr first, string text, uint flags, uint timeout, out UIntPtr result);
            [DllImport("user32.dll", EntryPoint = "PostMessageW", SetLastError = true)] private static extern bool Post(IntPtr window, uint message, UIntPtr first, IntPtr second);
            [DllImport("ole32.dll", EntryPoint = "GetRunningObjectTable")] private static extern int GetRot(int reserved, out IRunningObjectTable table);
            [DllImport("ole32.dll", EntryPoint = "CreateBindCtx")] private static extern int CreateContext(int reserved, out IBindCtx context);
        }
    }
}
