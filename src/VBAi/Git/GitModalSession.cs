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

        /// <summary>Owns the request state and operations.</summary>
        internal sealed class Request
        {

            /// <summary>Identifies the id associated with request.</summary>
            internal readonly string Id = Guid.NewGuid().ToString("N");

            /// <summary>Keeps the action and name and text and choice and path and revision path available to request.</summary>
            internal readonly string Action, Name, Text, Choice, Path, Revision;

            /// <summary>Maintains the references state for request.</summary>
            internal readonly bool References;

            /// <summary>Maintains the modules state for request.</summary>
            private readonly string[] modules;

            /// <summary>Maintains the revalidate state for request.</summary>
            internal readonly Func<Task> Revalidate;

            /// <summary>Maintains the admission state for request.</summary>
            private readonly TaskCompletionSource<bool> admission = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);

            /// <summary>Maintains the completion state for request.</summary>
            private readonly TaskCompletionSource<bool> completion = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);

            /// <summary>Gets or sets the phase.</summary>
            /// <value>Current phase exposed by request.</value>
            internal string Phase { get; private set; } = "Prepared";

            /// <summary>Gets or sets the error.</summary>
            /// <value>Current error exposed by request.</value>
            internal Exception Error { get; private set; }

            /// <summary>Gets the admission.</summary>
            /// <value>Current admission exposed by request.</value>
            internal Task Admission => admission.Task;

            /// <summary>Gets the completion.</summary>
            /// <value>Current completion exposed by request.</value>
            internal Task Completion => completion.Task;

            /// <summary>Gets the modules.</summary>
            /// <value>Current modules exposed by request.</value>
            internal string[] Modules => (string[])modules.Clone();

            /// <summary>Initializes a Request instance with the supplied state.</summary>
            /// <param name="action">Text that supplies the action value. Use the format required by the calling operation.</param>
            /// <param name="name">Text that supplies the name value. Use the format required by the calling operation.</param>
            /// <param name="text">Text that supplies the text value. Use the format required by the calling operation.</param>
            /// <param name="choice">Text that supplies the choice value. Use the format required by the calling operation.</param>
            /// <param name="path">Path used for the path being processed.</param>
            /// <param name="selected">string[] that supplies the selected for this operation.</param>
            /// <param name="references">Indicates whether references is enabled.</param>
            /// <param name="revision">Text that supplies the revision value. Use the format required by the calling operation.</param>
            /// <param name="revalidate">func&lt;task&gt; that supplies the revalidate for this operation.</param>
            internal Request(string action, string name, string text, string choice, string path, string[] selected, bool references, string revision, Func<Task> revalidate = null)
            {
                if (!RequiresHandoff(action) || string.IsNullOrWhiteSpace(revision)) throw new ArgumentException("An import request and exact revision are required.");
                Action = action; Name = name; Text = text; Choice = choice; Path = path;
                modules = selected == null ? new string[0] : (string[])selected.Clone(); References = references; Revision = revision;
                Revalidate = revalidate ?? (() => Task.CompletedTask);
            }

            /// <summary>Handles queue for request.</summary>
            internal void Queue()
            {
                if (Phase != "Prepared") throw new InvalidOperationException("An import request can be queued only once.");
                Phase = "AwaitingModalReturn";
            }

            /// <summary>Releases  for request.</summary>
            /// <param name="validate">action that supplies the validate for this operation.</param>
            internal void Release(Action validate)
            {
                if (Phase != "AwaitingModalReturn") throw new InvalidOperationException("An import request can be admitted only once.");
                try { validate(); Phase = "Executing"; admission.SetResult(true); }
                catch (Exception error) { Reject(error); }
            }

            /// <summary>Handles reject for request.</summary>
            /// <param name="error">Exception describing the error failure.</param>
            internal void Reject(Exception error)
            {
                if (Phase != "AwaitingModalReturn") throw new InvalidOperationException("Cannot reject an import already admitted.");
                Error = error; Phase = "Refused"; admission.SetException(error);
            }

            /// <summary>Handles complete for request.</summary>
            /// <param name="error">Exception describing the error failure.</param>
            /// <param name="publish">action that supplies the publish for this operation.</param>
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

        /// <summary>Maintains the show and validate owner state for git modal session.</summary>
        private readonly Action show, validateOwner;

        /// <summary>Maintains the pending state for git modal session.</summary>
        private Request pending;

        /// <summary>Maintains the active state for git modal session.</summary>
        private Request active;

        /// <summary>Maintains the inside modal and started and finished state for git modal session.</summary>
        private bool insideModal, started, finished;

        /// <summary>Identifies the id associated with git modal session.</summary>
        internal readonly string Id = Guid.NewGuid().ToString("N");

        /// <summary>Initializes a GitModalSession instance with the supplied state.</summary>
        /// <param name="show">action that supplies the show for this operation.</param>
        /// <param name="validateOwner">action that supplies the validate owner for this operation.</param>
        internal GitModalSession(Action show, Action validateOwner)
        { this.show = show ?? throw new ArgumentNullException(nameof(show)); this.validateOwner = validateOwner ?? throw new ArgumentNullException(nameof(validateOwner)); }

        /// <summary>Requires s handoff for git modal session.</summary>
        /// <param name="action">Text that supplies the action value. Use the format required by the calling operation.</param>
        /// <returns>Boolean indicating the result of the check for requires handoff on git modal session.</returns>
        internal static bool RequiresHandoff(string action) => action == "checkpoint_restore" || action == "branch_switch" ||
            action == "module_restore" || action == "pull" || action == "merge_complete" || action == "rollback";

        /// <summary>Handles queue for git modal session.</summary>
        /// <param name="request">request that supplies the request for this operation.</param>
        /// <param name="leaveModal">action that supplies the leave modal for this operation.</param>
        /// <returns>task produced by the operation for queue on git modal session.</returns>
        internal Task Queue(Request request, Action leaveModal)
        {
            if (!insideModal || finished || pending != null || request == null) throw new InvalidOperationException("No unique live modal can accept this request.");
            request.Queue(); pending = request;
            try { leaveModal(); }
            catch (Exception error) { request.Reject(error); }
            return request.Admission;
        }

        /// <summary>Only this stack frame can attest return from show; posted callbacks cannot release admission.</summary>
        /// <returns>task produced by the operation for run async on git modal session.</returns>
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

        /// <summary>Requires import owner for git modal session.</summary>
        /// <param name="request">request that supplies the request for this operation.</param>
        internal void RequireImportOwner(Request request)
        {
            if (insideModal || finished || !ReferenceEquals(active, request) || request == null || request.Phase != "Executing")
                throw new InvalidOperationException("Native import requires the exact admitted Git operation outside its modal loop.");
            validateOwner();
        }

        /// <summary>Handles show async for git modal session.</summary>
        /// <param name="window">git window that supplies the window for this operation.</param>
        /// <param name="owner">i win32 window that supplies the owner for this operation.</param>
        /// <param name="show">func&lt;form, i win32 window, dialog result&gt; that supplies the show for this operation.</param>
        /// <returns>task produced by the operation for show async on git modal session.</returns>
        internal static async Task ShowAsync(GitWindow window, IWin32Window owner, Func<Form, IWin32Window, DialogResult> show)
        {
            using (var lease = TryAcquire(owner))
            {
                if (lease == null) throw new InvalidOperationException("The exact Git owner already has a live session.");
                await lease.ShowAsync(window, show);
            }
        }

        /// <summary>Captures owner for git modal session.</summary>
        /// <param name="owner">i win32 window that supplies the owner for this operation.</param>
        /// <returns>action produced by the operation for capture owner on git modal session.</returns>
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

        /// <summary>Handles show owned async for git modal session.</summary>
        /// <param name="window">git window that supplies the window for this operation.</param>
        /// <param name="owner">i win32 window that supplies the owner for this operation.</param>
        /// <param name="show">func&lt;form, i win32 window, dialog result&gt; that supplies the show for this operation.</param>
        /// <param name="validate">action that supplies the validate for this operation.</param>
        /// <returns>task produced by the operation for show owned async on git modal session.</returns>
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

        /// <summary>Requires owner for git modal session.</summary>
        /// <param name="pid">uint that supplies the pid for this operation.</param>
        /// <param name="tid">uint that supplies the tid for this operation.</param>
        /// <param name="current">uint that supplies the current for this operation.</param>
        /// <param name="handle">Native handle that supplies the handle for this operation.</param>
        /// <param name="actualPid">uint that supplies the actual pid for this operation.</param>
        /// <param name="actualTid">uint that supplies the actual tid for this operation.</param>
        /// <param name="currentNow">uint that supplies the current now for this operation.</param>
        /// <param name="exists">Indicates whether exists is enabled.</param>
        /// <param name="enabled">Indicates whether enabled is enabled.</param>
        internal static void RequireOwner(uint pid, uint tid, uint current, IntPtr handle, uint actualPid, uint actualTid,
            uint currentNow, bool exists, bool enabled)
        {
            if (handle == IntPtr.Zero || pid == 0 || tid == 0 || tid != current || currentNow != current ||
                actualPid != pid || actualTid != tid || !exists || !enabled)
                throw new InvalidOperationException("The exact original Git owner must be enabled on its owning STA after modal return.");
        }

        /// <summary>Returns window thread process id for git modal session.</summary>
        /// <param name="hwnd">Native handle that supplies the hwnd for this operation.</param>
        /// <param name="pid">uint that supplies the pid for this operation.</param>
        /// <returns>uint produced by the operation for get window thread process id on git modal session.</returns>
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);

        /// <summary>Returns current thread id for git modal session.</summary>
        /// <returns>uint produced by the operation for get current thread id on git modal session.</returns>
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();

        /// <summary>Determines whether window for git modal session.</summary>
        /// <param name="hwnd">Native handle that supplies the hwnd for this operation.</param>
        /// <returns>Boolean indicating the result of the check for is window on git modal session.</returns>
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr hwnd);

        /// <summary>Determines whether window enabled for git modal session.</summary>
        /// <param name="hwnd">Native handle that supplies the hwnd for this operation.</param>
        /// <returns>Boolean indicating the result of the check for is window enabled on git modal session.</returns>
        [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr hwnd);
    }
}
