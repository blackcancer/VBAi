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
    /// <summary>Owns the vbe project general operation state and operations.</summary>
    internal sealed class VbeProjectGeneralOperation
    {
        // Only the pure, strict encoding round trip may produce this refusal.
        // Native identity, code-page discovery and authorization failures do not.
        /// <summary>Owns the text representation refused exception state and operations.</summary>
        internal sealed class TextRepresentationRefusedException : InvalidOperationException
        {

            /// <summary>Initializes a TextRepresentationRefusedException instance with the supplied state.</summary>
            /// <param name="inner">Exception describing the inner failure.</param>
            internal TextRepresentationRefusedException(Exception inner = null)
                : base("Native General text cannot preserve the exact requested value in its code page; no field write was entered.", inner) { }
        }

        /// <summary>Owns the snapshot state and operations.</summary>
        internal sealed class Snapshot
        {

            /// <summary>Maintains the dialog and page and tab and context and name edit and description edit and help file edit and compilation edit state for snapshot.</summary>
            internal IntPtr Dialog, Page, Tab, Context, NameEdit, DescriptionEdit, HelpFileEdit, CompilationEdit;

            /// <summary>Maintains the name and description and help file and context text and compilation state for snapshot.</summary>
            internal string Name, Description, HelpFile, ContextText, Compilation;

            /// <summary>Gets the options version.</summary>
            /// <value>Current options version exposed by snapshot.</value>
            internal string OptionsVersion
            {
                get
                {
                    // Length-prefixed fields avoid ambiguous concatenation and exclude transient handles.
                    string text = Pack(Name) + Pack(Description) + Pack(HelpFile) + Pack(ContextText) + Pack(Compilation);
                    using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(text))).Replace("-", "").ToLowerInvariant();
                }
            }

            /// <summary>Handles pack for snapshot.</summary>
            /// <param name="text">Text that supplies the text value. Use the format required by the calling operation.</param>
            /// <returns>Text produced by the operation for pack on snapshot.</returns>
            private static string Pack(string text) => (text ?? "").Length.ToString(CultureInfo.InvariantCulture) + ":" + (text ?? "");

            /// <summary>Compares native for snapshot.</summary>
            /// <param name="other">snapshot that supplies the other for this operation.</param>
            /// <returns>Boolean indicating the result of the check for same native on snapshot.</returns>
            internal bool SameNative(Snapshot other) => other != null && Dialog == other.Dialog && Page == other.Page && Tab == other.Tab && Context == other.Context && NameEdit == other.NameEdit && DescriptionEdit == other.DescriptionEdit && HelpFileEdit == other.HelpFileEdit && CompilationEdit == other.CompilationEdit;

            /// <summary>Handles unchanged except context for snapshot.</summary>
            /// <param name="other">snapshot that supplies the other for this operation.</param>
            /// <returns>Boolean indicating the result of the check for unchanged except context on snapshot.</returns>
            internal bool UnchangedExceptContext(Snapshot other) => SameNative(other) && Name == other.Name && Description == other.Description && HelpFile == other.HelpFile && Compilation == other.Compilation;

            /// <summary>Handles unchanged except help file for snapshot.</summary>
            /// <param name="other">snapshot that supplies the other for this operation.</param>
            /// <returns>Boolean indicating the result of the check for unchanged except help file on snapshot.</returns>
            internal bool UnchangedExceptHelpFile(Snapshot other) => SameNative(other) && Name == other.Name && Description == other.Description && ContextText == other.ContextText && Compilation == other.Compilation;
        }

        /// <summary>Owns the result state and operations.</summary>
        internal sealed class Result
        {

            /// <summary>Maintains the available and mutation invoked and uncertain and control value verified and committed requested and dialog closed and command entered and original execute returned and terminal and refused before write state for result.</summary>
            public bool Available, MutationInvoked, Uncertain, ControlValueVerified, CommittedRequested, DialogClosed, CommandEntered, OriginalExecuteReturned, Terminal, RefusedBeforeWrite;

            /// <summary>Gets the persistence verified.</summary>
            /// <value>Current persistence verified exposed by result.</value>
            public bool PersistenceVerified => false;

            /// <summary>Gets the retry allowed.</summary>
            /// <value>Current retry allowed exposed by result.</value>
            public bool RetryAllowed => false;

            /// <summary>Maintains the options version and name and description and help file and help context text and conditional compilation and error state for result.</summary>
            public string OptionsVersion, Name, Description, HelpFile, HelpContextText, ConditionalCompilation, Error;

            /// <summary>Tracks the open attempts and field attempts and ok attempts and cancel attempts state of result.</summary>
            public int OpenAttempts, FieldAttempts, OkAttempts, CancelAttempts;
        }

        /// <summary>Defines the i native contract.</summary>
        internal interface INative
        {

            /// <summary>Requires owner for i native.</summary>
            void RequireOwner();

            /// <summary>Handles prepare for i native.</summary>
            void Prepare(); // Original root enabled, no preexisting owned visible modal.

            /// <summary>Captures  for i native.</summary>
            /// <param name="exactProjectName">Text that supplies the exact project name value. Use the format required by the calling operation.</param>
            /// <returns>snapshot produced by the operation for capture on i native.</returns>
            Snapshot Capture(string exactProjectName); // Null only if no owned modal; unknown shapes throw.

            /// <summary>Requires same for i native.</summary>
            /// <param name="expected">snapshot that supplies the expected for this operation.</param>
            /// <param name="compareValues">Indicates whether compare values is enabled.</param>
            void RequireSame(Snapshot expected, bool compareValues);

            /// <summary>Requires help file representable for i native.</summary>
            /// <param name="expected">snapshot that supplies the expected for this operation.</param>
            /// <param name="value">Text that supplies the value value. Use the format required by the calling operation.</param>
            void RequireHelpFileRepresentable(Snapshot expected, string value);

            /// <summary>Writes context for i native.</summary>
            /// <param name="expected">snapshot that supplies the expected for this operation.</param>
            /// <param name="value">int that supplies the value for this operation.</param>
            /// <param name="beforeEntry">action that supplies the before entry for this operation.</param>
            void WriteContext(Snapshot expected, int value, Action beforeEntry);

            /// <summary>Writes help file for i native.</summary>
            /// <param name="expected">snapshot that supplies the expected for this operation.</param>
            /// <param name="value">Text that supplies the value value. Use the format required by the calling operation.</param>
            /// <param name="beforeEntry">action that supplies the before entry for this operation.</param>
            void WriteHelpFile(Snapshot expected, string value, Action beforeEntry);

            /// <summary>Closes  for i native.</summary>
            /// <param name="expected">snapshot that supplies the expected for this operation.</param>
            /// <param name="buttonId">int that supplies the button id for this operation.</param>
            /// <param name="beforeEnqueue">action that supplies the before enqueue for this operation.</param>
            void Close(Snapshot expected, int buttonId, Action beforeEnqueue);

            /// <summary>Closes d for i native.</summary>
            /// <param name="expected">snapshot that supplies the expected for this operation.</param>
            /// <returns>Boolean indicating the result of the check for closed on i native.</returns>
            bool Closed(Snapshot expected);
        }

        /// <summary>Defines the i scheduler contract.</summary>
        internal interface IScheduler
        {

            /// <summary>Requires owner for i scheduler.</summary>
            void RequireOwner();

            /// <summary>Gets the elapsed milliseconds.</summary>
            /// <value>Current elapsed milliseconds exposed by i scheduler.</value>
            long ElapsedMilliseconds { get; }

            /// <summary>Handles post for i scheduler.</summary>
            /// <param name="callback">action that supplies the callback for this operation.</param>
            void Post(Action callback);

            /// <summary>Handles poll for i scheduler.</summary>
            /// <param name="callback">action that supplies the callback for this operation.</param>
            /// <returns>i disposable produced by the operation for poll on i scheduler.</returns>
            IDisposable Poll(Action callback);
        }

        /// <summary>Owns the owner scheduler state and operations.</summary>
        private sealed class OwnerScheduler : IScheduler
        {

            /// <summary>Maintains the context state for owner scheduler.</summary>
            private readonly SynchronizationContext context;

            /// <summary>Maintains the owner state for owner scheduler.</summary>
            private readonly int owner;

            /// <summary>Maintains the clock state for owner scheduler.</summary>
            private readonly Stopwatch clock = Stopwatch.StartNew();

            /// <summary>Initializes a OwnerScheduler instance with the supplied state.</summary>
            internal OwnerScheduler()
            {
                context = SynchronizationContext.Current;
                owner = Thread.CurrentThread.ManagedThreadId;
                if (context == null) throw new InvalidOperationException("Original VBE UI context required.");
                RequireOwner();
            }

            /// <summary>Requires owner for owner scheduler.</summary>
            public void RequireOwner()
            {
                if (Thread.CurrentThread.ManagedThreadId != owner || Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                    throw new InvalidOperationException("General scheduler left its original VBE UI STA.");
            }

            /// <summary>Gets the elapsed milliseconds.</summary>
            /// <value>Current elapsed milliseconds exposed by owner scheduler.</value>
            public long ElapsedMilliseconds => clock.ElapsedMilliseconds;

            /// <summary>Handles post for owner scheduler.</summary>
            /// <param name="callback">action that supplies the callback for this operation.</param>
            public void Post(Action callback) { RequireOwner(); context.Post(_ => callback(), null); }

            /// <summary>Handles poll for owner scheduler.</summary>
            /// <param name="callback">action that supplies the callback for this operation.</param>
            /// <returns>i disposable produced by the operation for poll on owner scheduler.</returns>
            public IDisposable Poll(Action callback)
            {
                RequireOwner(); var timer = new System.Windows.Forms.Timer { Interval = 50 };
                timer.Tick += (sender, args) => callback(); timer.Start(); return timer;
            }
        }

        /// <summary>Maintains the native state for vbe project general operation.</summary>
        private readonly INative native;

        /// <summary>Maintains the scheduler state for vbe project general operation.</summary>
        private readonly IScheduler scheduler;

        /// <summary>Maintains the consumed state for vbe project general operation.</summary>
        private bool consumed;

        /// <summary>Initializes a VbeProjectGeneralOperation instance with the supplied state.</summary>
        /// <param name="native">i native that supplies the native for this operation.</param>
        /// <param name="scheduler">i scheduler that supplies the scheduler for this operation.</param>
        internal VbeProjectGeneralOperation(INative native, IScheduler scheduler = null)
        {
            this.native = native ?? throw new ArgumentNullException(nameof(native)); this.scheduler = scheduler;
        }

        // All callbacks run on the original VBE UI STA. Full authorization includes live COM reads;
        // pure authorization must perform no COM/host reads and is repeated after native getters.
        /// <summary>Runs async for vbe project general operation.</summary>
        /// <param name="exactProjectName">Text that supplies the exact project name value. Use the format required by the calling operation.</param>
        /// <param name="contextValue">int that supplies the context value for this operation.</param>
        /// <param name="expectedOptionsVersion">Text that supplies the expected options version value. Use the format required by the calling operation.</param>
        /// <param name="authorizeLiveTarget">action that supplies the authorize live target for this operation.</param>
        /// <param name="authorizeCachedPolicy">action that supplies the authorize cached policy for this operation.</param>
        /// <param name="openExactCommand">action&lt;action&gt; that supplies the open exact command for this operation.</param>
        /// <param name="durableClaim">action&lt;result&gt; that supplies the durable claim for this operation.</param>
        /// <param name="helpFileValue">Text that supplies the help file value value. Use the format required by the calling operation.</param>
        /// <returns>task&lt;result&gt; produced by the operation for run async on vbe project general operation.</returns>
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
