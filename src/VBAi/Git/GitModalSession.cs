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

        /// <summary>Immutable import request plus one-shot admission and completion signals.</summary>
        internal sealed class Request
        {

            /// <summary>Unique correlation ID for this modal handoff.</summary>
            internal readonly string Id = Guid.NewGuid().ToString("N");

            /// <summary>Operation payload captured when the user requests a Git import action, including its exact repository revision.</summary>
            internal readonly string Action, Name, Text, Choice, Path, Revision;

            /// <summary>Whether the admitted operation may also update project references.</summary>
            internal readonly bool References;

            /// <summary>Defensive copy of selected module names; callers receive a clone.</summary>
            private readonly string[] modules;

            /// <summary>Async repository revalidation that runs after the real modal loop returns and before admission.</summary>
            internal readonly Func<Task> Revalidate;

            /// <summary>Completes once when the original modal stack validates or refuses this import.</summary>
            private readonly TaskCompletionSource<bool> admission = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);

            /// <summary>Completes after import execution and its terminal UI publication have both finished.</summary>
            private readonly TaskCompletionSource<bool> completion = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);

            /// <summary>One-way request phase: Prepared, AwaitingModalReturn, Executing, Refused, Succeeded, or Failed.</summary>
            /// <value>Current phase; admission and completion methods enforce legal transitions.</value>
            internal string Phase { get; private set; } = "Prepared";

            /// <summary>Admission, execution, or terminal-publication failure, if any.</summary>
            /// <value>The first failure, or an aggregate when both execution and terminal presentation fail.</value>
            internal Exception Error { get; private set; }

            /// <summary>Gets the one-shot task that signals import admission or refusal.</summary>
            /// <value>Task faults with the refusal reason; it does not signal successful execution.</value>
            internal Task Admission => admission.Task;

            /// <summary>Gets the task that signals terminal execution and presentation.</summary>
            /// <value>Task faults when execution or terminal presentation fails.</value>
            internal Task Completion => completion.Task;

            /// <summary>Gets a defensive copy of selected module names.</summary>
            /// <value>A new array that cannot mutate the request's captured selection.</value>
            internal string[] Modules => (string[])modules.Clone();

            /// <summary>Captures the payload for a revision-bound native import handoff.</summary>
            /// <param name="action">Import action accepted by <see cref="RequiresHandoff"/>.</param>
            /// <param name="name">Optional target branch, module, or checkpoint name used by the action.</param>
            /// <param name="text">Optional user-provided action text such as a commit message.</param>
            /// <param name="choice">Action-specific selected choice.</param>
            /// <param name="path">Repository or target path captured by the Git window.</param>
            /// <param name="selected">Selected module names; copied to prevent later caller mutation.</param>
            /// <param name="references"><see langword="true"/> when the import should include reference updates.</param>
            /// <param name="revision">Exact repository revision against which the operation was prepared; required and non-empty.</param>
            /// <param name="revalidate">Optional asynchronous check run after modal return, before the final owner validation.</param>
            internal Request(string action, string name, string text, string choice, string path, string[] selected, bool references, string revision, Func<Task> revalidate = null)
            {
                if (!RequiresHandoff(action) || string.IsNullOrWhiteSpace(revision)) throw new ArgumentException("An import request and exact revision are required.");
                Action = action; Name = name; Text = text; Choice = choice; Path = path;
                modules = selected == null ? new string[0] : (string[])selected.Clone(); References = references; Revision = revision;
                Revalidate = revalidate ?? (() => Task.CompletedTask);
            }

            /// <summary>Moves a prepared request to the modal-return wait state.</summary>
            internal void Queue()
            {
                if (Phase != "Prepared") throw new InvalidOperationException("An import request can be queued only once.");
                Phase = "AwaitingModalReturn";
            }

            /// <summary>Admits the import only after its caller-provided owner/revision validation succeeds.</summary>
            /// <param name="validate">Synchronous final check executed before the request enters Executing.</param>
            internal void Release(Action validate)
            {
                if (Phase != "AwaitingModalReturn") throw new InvalidOperationException("An import request can be admitted only once.");
                try { validate(); Phase = "Executing"; admission.SetResult(true); }
                catch (Exception error) { Reject(error); }
            }

            /// <summary>Refuses a request that has not yet crossed its admission boundary.</summary>
            /// <param name="error">Reason the modal return, owner, revision, or revalidation could not be accepted.</param>
            internal void Reject(Exception error)
            {
                if (Phase != "AwaitingModalReturn") throw new InvalidOperationException("Cannot reject an import already admitted.");
                Error = error; Phase = "Refused"; admission.SetException(error);
            }

            /// <summary>Publishes one terminal result and resolves completion; a publication failure is retained with the operation error.</summary>
            /// <param name="error">Execution error, or null after successful execution.</param>
            /// <param name="publish">Optional UI terminal-state publisher invoked exactly once.</param>
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

        /// <summary>Actual modal display and original-owner validator supplied by the session creator.</summary>
        private readonly Action show, validateOwner;

        /// <summary>Single import request queued while the actual modal call is active.</summary>
        private Request pending;

        /// <summary>Request admitted by the modal return stack and currently executing outside that loop.</summary>
        private Request active;

        /// <summary>Guards against duplicate starts and distinguishes the modal stack from post-return import execution.</summary>
        private bool insideModal, started, finished;

        /// <summary>Unique identity for this exact owner-bound modal session.</summary>
        internal readonly string Id = Guid.NewGuid().ToString("N");

        /// <summary>Creates a one-shot modal runner for one Git window and its captured owner.</summary>
        /// <param name="show">Synchronous modal call; only its returning stack can attest modal completion.</param>
        /// <param name="validateOwner">Check that the original HWND, process generation, thread, and enabled state remain valid.</param>
        internal GitModalSession(Action show, Action validateOwner)
        { this.show = show ?? throw new ArgumentNullException(nameof(show)); this.validateOwner = validateOwner ?? throw new ArgumentNullException(nameof(validateOwner)); }

        /// <summary>Identifies native-changing Git actions that require the post-modal admission handoff.</summary>
        /// <param name="action">Git operation identifier.</param>
        /// <returns><see langword="true"/> for checkpoint restore, branch switch, module restore, pull, merge completion, and rollback.</returns>
        internal static bool RequiresHandoff(string action) => action == "checkpoint_restore" || action == "branch_switch" ||
            action == "module_restore" || action == "pull" || action == "merge_complete" || action == "rollback";

        /// <summary>Queues exactly one import request from the live modal, then asks its callback to return the dialog.</summary>
        /// <param name="request">Prepared revision-bound import request.</param><param name="leaveModal">Dialog-close action invoked after the request is queued.</param>
        /// <returns>The task that completes when the modal-return stack admits or refuses the request.</returns>
        internal Task Queue(Request request, Action leaveModal)
        {
            if (!insideModal || finished || pending != null || request == null) throw new InvalidOperationException("No unique live modal can accept this request.");
            request.Queue(); pending = request;
            try { leaveModal(); }
            catch (Exception error) { request.Reject(error); }
            return request.Admission;
        }

        /// <summary>Only this stack frame can attest return from show; posted callbacks cannot release admission.</summary>
        /// <returns>A task for the modal loop and any admitted operation; failures are propagated without retrying the import.</returns>
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
        /// <param name="request">Exact request currently admitted by this session.</param>
        internal void RequireImportOwner(Request request)
        {
            if (insideModal || finished || !ReferenceEquals(active, request) || request == null || request.Phase != "Executing")
                throw new InvalidOperationException("Native import requires the exact admitted Git operation outside its modal loop.");
            validateOwner();
        }

        /// <summary>Shows a Git window under a unique lease tied to the supplied native owner.</summary>
        /// <param name="window">Git dialog to show.</param><param name="owner">Original native owner whose HWND and thread are validated.</param>
        /// <param name="show">Modal display route, normally Form.ShowDialog.</param>
        /// <returns>A task completed after the modal session and any handoff have finished.</returns>
        internal static async Task ShowAsync(GitWindow window, IWin32Window owner, Func<Form, IWin32Window, DialogResult> show)
        {
            using (var lease = TryAcquire(owner))
            {
                if (lease == null) throw new InvalidOperationException("The exact Git owner already has a live session.");
                await lease.ShowAsync(window, show);
            }
        }

        /// <summary>Captures the owner HWND, process generation, and STA thread, returning a closure that rechecks them later.</summary>
        /// <param name="owner">Native window that owns this Git modal session.</param>
        /// <returns>Validator for the exact captured HWND and process generation.</returns>
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

        /// <summary>Attaches a fresh modal session, runs it on the owner STA, and detaches it on every completion path.</summary>
        /// <param name="window">Git form participating in the one-shot handoff.</param><param name="owner">Native owner passed to the modal display route.</param>
        /// <param name="show">Synchronous modal display function.</param><param name="validate">Captured owner validator.</param>
        /// <returns>A task completed after modal return and request processing.</returns>
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

        /// <summary>Requires the same live, enabled HWND, process, and STA thread captured before modal display.</summary>
        /// <param name="pid">Original owner process ID.</param><param name="tid">Original owner thread ID.</param>
        /// <param name="current">Calling native thread ID captured before display.</param><param name="handle">Original owner HWND.</param>
        /// <param name="actualPid">Process currently owning the HWND.</param><param name="actualTid">Thread currently owning the HWND.</param>
        /// <param name="currentNow">Calling native thread ID at revalidation.</param><param name="exists">Whether the HWND is still valid.</param>
        /// <param name="enabled">Whether the owner window accepts input.</param>
        internal static void RequireOwner(uint pid, uint tid, uint current, IntPtr handle, uint actualPid, uint actualTid,
            uint currentNow, bool exists, bool enabled)
        {
            if (handle == IntPtr.Zero || pid == 0 || tid == 0 || tid != current || currentNow != current ||
                actualPid != pid || actualTid != tid || !exists || !enabled)
                throw new InvalidOperationException("The exact original Git owner must be enabled on its owning STA after modal return.");
        }

        /// <summary>Reads the process and thread that own a native window handle.</summary>
        /// <param name="hwnd">Window whose owner is queried.</param><param name="pid">Receives the owning process ID.</param>
        /// <returns>Owning thread ID.</returns>
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);

        /// <summary>Reads the native ID of the calling Windows thread.</summary><returns>Current native thread ID.</returns>
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();

        /// <summary>Checks whether an HWND still refers to a live native window.</summary>
        /// <param name="hwnd">Window handle to validate.</param><returns><see langword="true"/> when the handle is valid.</returns>
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr hwnd);

        /// <summary>Checks whether the native owner window currently accepts user input.</summary>
        /// <param name="hwnd">Window handle to inspect.</param><returns><see langword="true"/> when the window is enabled.</returns>
        [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr hwnd);
    }
}
