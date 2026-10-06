using System;
using System.Diagnostics;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace VBAi
{
    // Proposed additive native General route. It is never invoked after a COM setter.
    /// <summary>Coordinates one bounded read or single-field mutation of the existing VBE Project Properties General page, preserving owner-thread, authorization, readback, receipt, and no-retry checks.</summary>
    internal sealed class VbeProjectGeneralOperation
    {
        // Only the pure, strict encoding round trip may produce this refusal.
        // Native identity, code-page discovery and authorization failures do not.
        /// <summary>Signals a known, exact-encoding refusal that occurred before any native field setter was entered.</summary>
        internal sealed class TextRepresentationRefusedException : InvalidOperationException
        {

            /// <summary>Creates the pre-write encoding refusal, optionally preserving the encoder or decoder failure.</summary>
            /// <param name="inner">Underlying fallback exception, when an encoding operation could not represent the text.</param>
            internal TextRepresentationRefusedException(Exception inner = null)
                : base("Native General text cannot preserve the exact requested value in its code page; no field write was entered.", inner) { }
        }

        /// <summary>Captures native dialog/control identity and the five General-page text values used for stale-state checks.</summary>
        internal sealed class Snapshot
        {

            /// <summary>HWNDs for the dialog, General page, tab, and each captured field control; these identities must remain stable across a write.</summary>
            internal IntPtr Dialog, Page, Tab, Context, NameEdit, DescriptionEdit, HelpFileEdit, CompilationEdit;

            /// <summary>Text read from the project's name, description, HelpFile, HelpContextID, and conditional-compilation controls.</summary>
            internal string Name, Description, HelpFile, ContextText, Compilation;

            /// <summary>Computes a stable digest of the captured General-page option text, excluding transient HWNDs.</summary>
            /// <value>Lowercase SHA-256 hex digest over length-prefixed UTF-8 fields.</value>
            internal string OptionsVersion
            {
                get
                {
                    // Length-prefixed fields avoid ambiguous concatenation and exclude transient handles.
                    string text = Pack(Name) + Pack(Description) + Pack(HelpFile) + Pack(ContextText) + Pack(Compilation);
                    using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(text))).Replace("-", "").ToLowerInvariant();
                }
            }

            /// <summary>Prefixes text with its invariant-culture character count so adjacent fields cannot collide when concatenated.</summary>
            /// <param name="text">Field value; null is encoded as an empty string.</param>
            /// <returns>Length-prefixed field in the form <c>length:value</c>.</returns>
            private static string Pack(string text) => (text ?? "").Length.ToString(CultureInfo.InvariantCulture) + ":" + (text ?? "");

            /// <summary>Checks that both snapshots refer to the same dialog, page, tab, and field HWNDs.</summary>
            /// <param name="other">Snapshot to compare; null never matches.</param>
            /// <returns><see langword="true"/> when all captured native identities are identical.</returns>
            internal bool SameNative(Snapshot other) => other != null && Dialog == other.Dialog && Page == other.Page && Tab == other.Tab && Context == other.Context && NameEdit == other.NameEdit && DescriptionEdit == other.DescriptionEdit && HelpFileEdit == other.HelpFileEdit && CompilationEdit == other.CompilationEdit;

            /// <summary>Checks native identity and all option text except HelpContextID.</summary>
            /// <param name="other">Post-write snapshot to compare.</param>
            /// <returns><see langword="true"/> only when identity, name, description, HelpFile, and compilation text are unchanged.</returns>
            internal bool UnchangedExceptContext(Snapshot other) => SameNative(other) && Name == other.Name && Description == other.Description && HelpFile == other.HelpFile && Compilation == other.Compilation;

            /// <summary>Checks native identity and all option text except HelpFile.</summary>
            /// <param name="other">Post-write snapshot to compare.</param>
            /// <returns><see langword="true"/> only when identity, name, description, HelpContextID, and compilation text are unchanged.</returns>
            internal bool UnchangedExceptHelpFile(Snapshot other) => SameNative(other) && Name == other.Name && Description == other.Description && ContextText == other.ContextText && Compilation == other.Compilation;
        }

        /// <summary>Reports operation evidence and terminal state; persistence and retry are deliberately never claimed by this native route.</summary>
        internal sealed class Result
        {

            /// <summary>Evidence flags for capability, native mutation entry, uncertainty, readback, requested commit, dialog closure, command execution, terminal receipt, and pre-write refusal.</summary>
            public bool Available, MutationInvoked, Uncertain, ControlValueVerified, CommittedRequested, DialogClosed, CommandEntered, OriginalExecuteReturned, Terminal, RefusedBeforeWrite;

            /// <summary>Indicates whether this route proved durable host persistence.</summary>
            /// <value>Always <see langword="false"/>; native dialog readback is not a persistence proof.</value>
            public bool PersistenceVerified => false;

            /// <summary>Indicates whether the original operation can be replayed.</summary>
            /// <value>Always <see langword="false"/> because a native mutation outcome may be uncertain.</value>
            public bool RetryAllowed => false;

            /// <summary>Returned snapshot values, digest, and diagnostic text; populated only as each stage is observed.</summary>
            public string OptionsVersion, Name, Description, HelpFile, HelpContextText, ConditionalCompilation, Error;

            /// <summary>One-shot attempt counts for opening the command, writing a field, and posting OK or Cancel.</summary>
            public int OpenAttempts, FieldAttempts, OkAttempts, CancelAttempts;
        }

        /// <summary>Defines the native VBE UI operations required by the coordinator.</summary>
        internal interface INative
        {

            /// <summary>Throws unless the caller is on the original VBE UI owner thread and native context.</summary>
            void RequireOwner();

            /// <summary>Requires an enabled original VBE root with no preexisting owned modal dialog.</summary>
            void Prepare(); // Original root enabled, no preexisting owned visible modal.

            /// <summary>Captures the visible General page without opening a dialog or changing tabs.</summary>
            /// <param name="exactProjectName">Canonical project name required to match the unique dialog and its Name field.</param>
            /// <returns>Captured native identity and values, or null only when no owned modal exists.</returns>
            Snapshot Capture(string exactProjectName); // Null only if no owned modal; unknown shapes throw.

            /// <summary>Re-captures and rejects changed dialog/control identity or, optionally, changed option values.</summary>
            /// <param name="expected">Previously captured concurrency token.</param>
            /// <param name="compareValues">When true, also requires the captured options digest to match.</param>
            void RequireSame(Snapshot expected, bool compareValues);

            /// <summary>Checks the current snapshot and verifies exact HelpFile round-tripping through the target edit encoding.</summary>
            /// <param name="expected">Snapshot whose identity and values must still match.</param>
            /// <param name="value">Requested HelpFile text; replacement, best-fit conversion, and normalization are refused.</param>
            void RequireHelpFileRepresentable(Snapshot expected, string value);

            /// <summary>Writes HelpContextID once after the final authorization callback.</summary>
            /// <param name="expected">Snapshot that identifies the current General page.</param>
            /// <param name="value">HelpContextID value to format as invariant decimal text.</param>
            /// <param name="beforeEntry">Final authorization check run immediately before the native setter.</param>
            void WriteContext(Snapshot expected, int value, Action beforeEntry);

            /// <summary>Writes HelpFile once after the final authorization callback.</summary>
            /// <param name="expected">Snapshot that identifies the current General page.</param>
            /// <param name="value">Exact HelpFile text to set.</param>
            /// <param name="beforeEntry">Final authorization check run immediately before the native setter.</param>
            void WriteHelpFile(Snapshot expected, string value, Action beforeEntry);

            /// <summary>Posts one verified OK or Cancel command to the captured dialog.</summary>
            /// <param name="expected">Snapshot identifying the original dialog and controls.</param>
            /// <param name="buttonId">Native button ID: 1 for OK or 2 for Cancel.</param>
            /// <param name="beforeEnqueue">Final authorization callback invoked immediately before posting.</param>
            void Close(Snapshot expected, int buttonId, Action beforeEnqueue);

            /// <summary>Checks whether the captured dialog is gone and the original VBE is enabled with no owned modal remaining.</summary>
            /// <param name="expected">Snapshot identifying the dialog whose closure is being checked.</param>
            /// <returns><see langword="true"/> when closure is proved; false while the original dialog remains open.</returns>
            bool Closed(Snapshot expected);
        }

        /// <summary>Defines timer and callback scheduling that remains bound to the original VBE UI STA.</summary>
        internal interface IScheduler
        {

            /// <summary>Throws unless scheduling is performed on the original VBE UI owner thread.</summary>
            void RequireOwner();

            /// <summary>Gets elapsed time from the scheduler's monotonic clock.</summary>
            /// <value>Milliseconds since scheduler construction.</value>
            long ElapsedMilliseconds { get; }

            /// <summary>Queues a callback onto the VBE UI synchronization context.</summary>
            /// <param name="callback">Work to run on the owner thread.</param>
            void Post(Action callback);

            /// <summary>Schedules recurring owner-thread polling until the returned timer is disposed.</summary>
            /// <param name="callback">Poll action invoked on each timer tick.</param>
            /// <returns>Timer handle that stops future poll callbacks when disposed.</returns>
            IDisposable Poll(Action callback);
        }

        /// <summary>Provides monotonic timing and WinForms timer callbacks on the captured VBE UI STA.</summary>
        private sealed class OwnerScheduler : IScheduler
        {

            /// <summary>Synchronization context captured from the thread that created the scheduler.</summary>
            private readonly SynchronizationContext context;

            /// <summary>Managed thread ID captured at construction and required for all scheduler operations.</summary>
            private readonly int owner;

            /// <summary>Monotonic stopwatch used for operation deadlines.</summary>
            private readonly Stopwatch clock = Stopwatch.StartNew();

            /// <summary>Captures the current STA and synchronization context, failing if no VBE UI context is installed.</summary>
            internal OwnerScheduler()
            {
                context = SynchronizationContext.Current;
                owner = Thread.CurrentThread.ManagedThreadId;
                if (context == null) throw new InvalidOperationException("Original VBE UI context required.");
                RequireOwner();
            }

            /// <summary>Throws when called from another thread or from a thread that is not an STA.</summary>
            public void RequireOwner()
            {
                if (Thread.CurrentThread.ManagedThreadId != owner || Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                    throw new InvalidOperationException("General scheduler left its original VBE UI STA.");
            }

            /// <summary>Gets elapsed time from the scheduler's monotonic stopwatch.</summary>
            /// <value>Milliseconds since scheduler construction.</value>
            public long ElapsedMilliseconds => clock.ElapsedMilliseconds;

            /// <summary>Posts work to the captured synchronization context after checking thread ownership.</summary>
            /// <param name="callback">Action dispatched asynchronously to the captured UI context.</param>
            public void Post(Action callback) { RequireOwner(); context.Post(_ => callback(), null); }

            /// <summary>Starts a 50 ms WinForms timer that invokes the callback on the owner thread.</summary>
            /// <param name="callback">Action invoked for each timer tick.</param>
            /// <returns>The started timer; disposing it stops the polling.</returns>
            public IDisposable Poll(Action callback)
            {
                RequireOwner(); var timer = new System.Windows.Forms.Timer { Interval = 50 };
                timer.Tick += (sender, args) => callback(); timer.Start(); return timer;
            }
        }

        /// <summary>Native VBE UI adapter used to inspect, mutate, and close the existing General dialog.</summary>
        private readonly INative native;

        /// <summary>Optional owner-thread scheduler; null selects the WinForms STA implementation.</summary>
        private readonly IScheduler scheduler;

        /// <summary>One-shot guard set before validation or native work so this coordinator instance can never be replayed.</summary>
        private bool consumed;

        /// <summary>Creates a single-use coordinator for an existing General dialog.</summary>
        /// <param name="native">Native adapter for owner checks, capture, the single field setter, and close.</param>
        /// <param name="scheduler">Optional owner-thread scheduler; null selects the WinForms STA implementation.</param>
        internal VbeProjectGeneralOperation(INative native, IScheduler scheduler = null)
        {
            this.native = native ?? throw new ArgumentNullException(nameof(native)); this.scheduler = scheduler;
        }

        // All callbacks run on the original VBE UI STA. Full authorization includes live COM reads;
        // pure authorization must perform no COM/host reads and is repeated after native getters.
        /// <summary>Runs one bounded read or one authorized HelpContextID/HelpFile mutation, records attempt receipts, verifies readback, and never replays uncertain native work.</summary>
        /// <param name="exactProjectName">Canonical project name used to bind the native dialog to the approved project.</param>
        /// <param name="contextValue">Requested HelpContextID; null means no context mutation is requested.</param>
        /// <param name="expectedOptionsVersion">Digest returned by the prior read; a write is refused when current options differ.</param>
        /// <param name="authorizeLiveTarget">Full authorization callback that may re-read the live COM target.</param>
        /// <param name="authorizeCachedPolicy">Cached-policy authorization repeated before each native stage and required to perform no host reads.</param>
        /// <param name="openExactCommand">Callback that resolves the original command and invokes the entry callback immediately before its original Execute.</param>
        /// <param name="durableClaim">Receipt callback that records each one-shot attempt before entering the corresponding native action.</param>
        /// <param name="helpFileValue">Requested HelpFile text; null means no HelpFile mutation is requested.</param>
        /// <returns>Task completed with observed operation evidence; this dialog route does not assert durable persistence.</returns>
        internal Task<Result> RunAsync(string exactProjectName, int? contextValue, string expectedOptionsVersion,
            Action authorizeLiveTarget, Action authorizeCachedPolicy, Action<Action> openExactCommand, Action<Result> durableClaim, string helpFileValue = null)
        {
            if (consumed) throw new InvalidOperationException("This original General operation cannot be retried.");
            consumed = true;
            if (string.IsNullOrWhiteSpace(exactProjectName) || authorizeLiveTarget == null || authorizeCachedPolicy == null || openExactCommand == null || durableClaim == null)
                throw new ArgumentException("Exact project and all authorization/claim callbacks are required.");
            bool write = contextValue.HasValue || helpFileValue != null;
            if (contextValue.HasValue && helpFileValue != null) throw new ArgumentException("Only one General property may be written per original operation.");
            if (write && string.IsNullOrWhiteSpace(expectedOptionsVersion)) throw new ArgumentException("ExpectedOptionsVersion from read_project_general is required for a write.");
            var scheduling = scheduler ?? new OwnerScheduler(); scheduling.RequireOwner();
            native.RequireOwner(); authorizeLiveTarget(); native.Prepare(); authorizeCachedPolicy(); native.RequireOwner();
            var result = new Result(); var completion = new TaskCompletionSource<Result>(TaskCreationOptions.RunContinuationsAsynchronously);
            long started = scheduling.ElapsedMilliseconds; IDisposable polling = null;
            Snapshot before = null; bool tickActive = false, commandEntered = false, commandReturned = false, closeClaimed = false, terminal = false;
            Action finish = () => {
                if (terminal) return; terminal = true; result.Terminal = true;
                try { polling?.Dispose(); }
                catch (Exception error) { result.Error = (result.Error == null ? "" : result.Error + " | ") + "Scheduler cleanup: " + error.Message; result.Uncertain |= commandEntered; }
                try { scheduling.RequireOwner(); native.RequireOwner(); durableClaim(result); }
                catch (Exception error) { result.Error = (result.Error == null ? "" : result.Error + " | ") + "Terminal receipt/owner: " + error.Message; result.Uncertain |= commandEntered; }
                finally { completion.TrySetResult(result); }
            };
            Action<Exception> fail = error => {
                if (terminal) return;
                result.Error = error.ToString();
                result.Uncertain = result.FieldAttempts != 0 || result.OkAttempts != 0 || (commandEntered && (!result.DialogClosed || !commandReturned));
                // A claimed field/OK or unsettled Execute is retained. No Cancel/second close fallback.
                finish();
            };
            Action requireDeadline = () => {
                scheduling.RequireOwner();
                long elapsed = scheduling.ElapsedMilliseconds - started;
                if (elapsed < 0 || elapsed > 20000) throw new InvalidOperationException("Original General operation deadline expired; no replay.");
            };
            Action cancelUnchanged = () => {
                if (result.FieldAttempts != 0 || result.OkAttempts != 0 || result.MutationInvoked || result.CancelAttempts != 0 || closeClaimed)
                    throw new InvalidOperationException("Only an unchanged original General inspection may be cancelled once.");
                authorizeLiveTarget(); native.RequireSame(before, true); authorizeCachedPolicy(); native.RequireOwner(); requireDeadline();
                result.CancelAttempts = 1; durableClaim(result); native.RequireSame(before, true); authorizeCachedPolicy(); native.RequireOwner(); requireDeadline();
                closeClaimed = true;
                native.Close(before, 2, () => { authorizeCachedPolicy(); native.RequireOwner(); requireDeadline(); });
            };
            Action tick = () => {
                if (terminal || tickActive) return; tickActive = true;
                try
                {
                    scheduling.RequireOwner(); native.RequireOwner();
                    requireDeadline();
                    if (!commandEntered) return; // Never adopt a dialog while command resolution/authorization is still pumping.
                    if (closeClaimed)
                    {
                        if (native.Closed(before))
                        {
                            result.DialogClosed = true;
                            if (commandReturned) finish();
                        }
                        return;
                    }
                    var observed = native.Capture(exactProjectName); if (observed == null) return;
                    before = observed;
                    result.Available = true; result.Name = before.Name; result.Description = before.Description;
                    result.HelpFile = before.HelpFile; result.HelpContextText = before.ContextText; result.ConditionalCompilation = before.Compilation;
                    result.OptionsVersion = before.OptionsVersion;
                    if (!write)
                    {
                        cancelUnchanged(); return;
                    }
                    if (!string.Equals(expectedOptionsVersion, before.OptionsVersion, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("General options changed; no mutation was entered.");
                    authorizeLiveTarget(); native.RequireSame(before, true); authorizeCachedPolicy(); native.RequireOwner();
                    if (helpFileValue != null)
                    {
                        try { native.RequireHelpFileRepresentable(before, helpFileValue); }
                        catch (TextRepresentationRefusedException error)
                        {
                            result.RefusedBeforeWrite = true; result.Available = false; result.Error = error.Message;
                            cancelUnchanged(); return;
                        }
                    }
                    requireDeadline(); result.FieldAttempts = 1; durableClaim(result); native.RequireSame(before, true); authorizeCachedPolicy(); native.RequireOwner();
                    Action entry = () => { authorizeCachedPolicy(); native.RequireOwner(); requireDeadline(); result.MutationInvoked = true; };
                    if (contextValue.HasValue) native.WriteContext(before, contextValue.Value, entry);
                    else native.WriteHelpFile(before, helpFileValue, entry);
                    var after = native.Capture(exactProjectName);
                    bool retained = contextValue.HasValue ? before.UnchangedExceptContext(after) && after.ContextText == contextValue.Value.ToString(CultureInfo.InvariantCulture)
                        : before.UnchangedExceptHelpFile(after) && after.HelpFile == helpFileValue;
                    if (!retained) throw new InvalidOperationException("General field readback differs; retain this original modal.");
                    result.ControlValueVerified = true; result.HelpContextText = after.ContextText; result.HelpFile = after.HelpFile; result.OptionsVersion = after.OptionsVersion;
                    authorizeLiveTarget(); native.RequireSame(after, true); authorizeCachedPolicy(); native.RequireOwner();
                    requireDeadline(); result.OkAttempts = 1; durableClaim(result); native.RequireSame(after, true); authorizeCachedPolicy(); native.RequireOwner();
                    closeClaimed = true; native.Close(after, 1, () => { authorizeCachedPolicy(); native.RequireOwner(); requireDeadline(); result.CommittedRequested = true; });
                }
                catch (Exception error) { fail(error); }
                finally { tickActive = false; }
            };
            try
            {
                polling = scheduling.Poll(tick);
                scheduling.Post(() => {
                if (terminal) return;
                try
                {
                    scheduling.RequireOwner(); native.RequireOwner(); authorizeLiveTarget(); native.Prepare(); authorizeCachedPolicy(); native.RequireOwner();
                    result.OpenAttempts = 1; durableClaim(result); authorizeLiveTarget(); native.Prepare(); authorizeCachedPolicy(); native.RequireOwner();
                    openExactCommand(() => {
                        native.Prepare(); authorizeCachedPolicy(); native.RequireOwner();
                        result.CommandEntered = commandEntered = true; // Only immediately before original Execute, after all COM reads.
                    });
                    result.OriginalExecuteReturned = commandReturned = true;
                    if (result.DialogClosed) finish();
                }
                catch (Exception error) { fail(error); }
                });
            }
            catch (Exception error) { fail(error); }
            return completion.Task;
        }
    }
}
