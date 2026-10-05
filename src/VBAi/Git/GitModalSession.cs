using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VBAi
{
    /// <summary>One-shot import admission after the actual modal call returns on its original owner STA.</summary>
    internal sealed partial class GitModalSession
    {
        internal sealed class Request
        {
            internal readonly string Id = Guid.NewGuid().ToString("N");
            internal readonly string Action, Name, Text, Choice, Path, Revision;
            internal readonly bool References;
            private readonly string[] modules;
            internal readonly Func<Task> Revalidate;
            private readonly TaskCompletionSource<bool> admission = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            private readonly TaskCompletionSource<bool> completion = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            internal string Phase { get; private set; } = "Prepared";
            internal Exception Error { get; private set; }
            internal Task Admission => admission.Task;
            internal Task Completion => completion.Task;
            internal string[] Modules => (string[])modules.Clone();
            internal Request(string action, string name, string text, string choice, string path, string[] selected, bool references, string revision, Func<Task> revalidate = null)
            {
                if (!RequiresHandoff(action) || string.IsNullOrWhiteSpace(revision)) throw new ArgumentException("An import request and exact revision are required.");
                Action = action; Name = name; Text = text; Choice = choice; Path = path;
                modules = selected == null ? new string[0] : (string[])selected.Clone(); References = references; Revision = revision;
                Revalidate = revalidate ?? (() => Task.CompletedTask);
            }
            internal void Queue()
            {
                if (Phase != "Prepared") throw new InvalidOperationException("An import request can be queued only once.");
                Phase = "AwaitingModalReturn";
            }
            internal void Release(Action validate)
            {
                if (Phase != "AwaitingModalReturn") throw new InvalidOperationException("An import request can be admitted only once.");
                try { validate(); Phase = "Executing"; admission.SetResult(true); }
                catch (Exception error) { Reject(error); }
            }
            internal void Reject(Exception error)
            {
                if (Phase != "AwaitingModalReturn") throw new InvalidOperationException("Cannot reject an import already admitted.");
                Error = error; Phase = "Refused"; admission.SetException(error);
            }
            internal void Complete(Exception error, Action publish = null)
            {
                if (Phase != "Executing" && Phase != "Refused") throw new InvalidOperationException("No admitted or refused request can complete twice.");
                Error = error ?? Error; Phase = Error == null ? "Succeeded" : "Failed";
                try { publish?.Invoke(); completion.SetResult(true); }
                catch (Exception publication)
                {
                    Error = Error == null ? publication : new AggregateException("Git operation and terminal presentation failed.", Error, publication);
                    Phase = "Failed"; completion.SetException(Error);
                }
            }
        }

        private readonly Action show, validateOwner;
        private Request pending;
        private Request active;
        private bool insideModal, started, finished;
        internal readonly string Id = Guid.NewGuid().ToString("N");
        internal GitModalSession(Action show, Action validateOwner)
        { this.show = show ?? throw new ArgumentNullException(nameof(show)); this.validateOwner = validateOwner ?? throw new ArgumentNullException(nameof(validateOwner)); }
        internal static bool RequiresHandoff(string action) => action == "checkpoint_restore" || action == "branch_switch" ||
            action == "module_restore" || action == "pull" || action == "merge_complete" || action == "rollback";

        internal Task Queue(Request request, Action leaveModal)
        {
            if (!insideModal || finished || pending != null || request == null) throw new InvalidOperationException("No unique live modal can accept this request.");
            request.Queue(); pending = request;
            try { leaveModal(); }
            catch (Exception error) { request.Reject(error); }
            return request.Admission;
        }

        /// <summary>Only this stack frame can attest return from show; posted callbacks cannot release admission.</summary>
        internal async Task RunAsync()
        {
            if (started) throw new InvalidOperationException("One modal session only.");
            started = true;
            try
            {
                while (true)
                {
                    Exception showError = null;
                    insideModal = true;
                    try { show(); }
                    catch (Exception error) { showError = error; }
                    finally { insideModal = false; }
                    Request request = pending; pending = null;
                    if (request == null) { if (showError != null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(showError).Throw(); return; }
                    active = request;
                    bool refusedBeforeReturn = request.Phase == "Refused" || request.Phase == "Failed";
                    if (!refusedBeforeReturn)
                    {
                        if (showError != null) request.Reject(showError);
                        else
                        {
                            try { validateOwner(); await request.Revalidate(); request.Release(validateOwner); }
                            catch (Exception error) { request.Reject(error); }
                        }
                    }
                    try { await request.Completion; }
                    catch (Exception completion)
                    {
                        if (showError != null && refusedBeforeReturn)
                            throw new AggregateException("Git modal return and refused operation completion both failed.", showError, completion);
                        throw;
                    }
                    active = null;
                    if (showError != null)
                    {
                        if (request.Error != null && !ReferenceEquals(request.Error, showError))
                            throw new AggregateException("Git modal return and its pending operation both failed.", showError, request.Error);
                        System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(showError).Throw();
                    }
                    if (refusedBeforeReturn) return;
                    // A completed failure is presented, never retried. An unfinished task keeps this session alive.
                }
            }
            finally { finished = true; }
        }

        internal void RequireImportOwner(Request request)
        {
            if (insideModal || finished || !ReferenceEquals(active, request) || request == null || request.Phase != "Executing")
                throw new InvalidOperationException("Native import requires the exact admitted Git operation outside its modal loop.");
            validateOwner();
        }

        internal static async Task ShowAsync(GitWindow window, IWin32Window owner, Func<Form, IWin32Window, DialogResult> show)
        {
            using (var lease = TryAcquire(owner))
            {
                if (lease == null) throw new InvalidOperationException("The exact Git owner already has a live session.");
                await lease.ShowAsync(window, show);
            }
        }

        private static Action CaptureOwner(IWin32Window owner)
        {
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                throw new InvalidOperationException("Git modal work requires its original owning STA.");
            if (owner == null) throw new ArgumentNullException(nameof(owner));
            IntPtr handle = owner.Handle;
            uint pid; uint tid = GetWindowThreadProcessId(handle, out pid);
            uint current = GetCurrentThreadId();
            int processId; long processBirth;
            using (var process = System.Diagnostics.Process.GetCurrentProcess())
            { processId = process.Id; processBirth = process.StartTime.ToUniversalTime().Ticks; }
            Action validate = () => {
                    uint actualPid; uint actualTid = GetWindowThreadProcessId(handle, out actualPid);
                    using (var process = System.Diagnostics.Process.GetCurrentProcess())
                        if (pid != processId || process.Id != processId || process.StartTime.ToUniversalTime().Ticks != processBirth)
                            throw new InvalidOperationException("The original Git owner process generation changed.");
                    RequireOwner(pid, tid, current, handle, actualPid, actualTid, GetCurrentThreadId(), IsWindow(handle), IsWindowEnabled(handle));
                };
            return validate;
        }

        private static Task ShowOwnedAsync(GitWindow window, IWin32Window owner, Func<Form, IWin32Window, DialogResult> show, Action validate)
        {
            if (window == null || show == null) throw new ArgumentNullException("Modal session dependencies");
            var session = new GitModalSession(() => { validate(); window.DialogResult = DialogResult.None; show(window, owner); }, validate);
            window.AttachModalSession(session);
            try { return VbeUiTask.Run(async () => { try { await session.RunAsync(); return true; } finally { window.DetachModalSession(session); } }); }
            catch (Exception original)
            {
                try { window.DetachModalSession(session); }
                catch (Exception cleanup) { throw new AggregateException(original, cleanup); }
                throw;
            }
        }

        internal static void RequireOwner(uint pid, uint tid, uint current, IntPtr handle, uint actualPid, uint actualTid,
            uint currentNow, bool exists, bool enabled)
        {
            if (handle == IntPtr.Zero || pid == 0 || tid == 0 || tid != current || currentNow != current ||
                actualPid != pid || actualTid != tid || !exists || !enabled)
                throw new InvalidOperationException("The exact original Git owner must be enabled on its owning STA after modal return.");
        }
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr hwnd);
        [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr hwnd);
    }
}
