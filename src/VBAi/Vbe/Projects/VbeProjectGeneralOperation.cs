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
    internal sealed class VbeProjectGeneralOperation
    {
        // Only the pure, strict encoding round trip may produce this refusal.
        // Native identity, code-page discovery and authorization failures do not.
        internal sealed class TextRepresentationRefusedException : InvalidOperationException
        {
            internal TextRepresentationRefusedException(Exception inner = null)
                : base("Native General text cannot preserve the exact requested value in its code page; no field write was entered.", inner) { }
        }
        internal sealed class Snapshot
        {
            internal IntPtr Dialog, Page, Tab, Context, NameEdit, DescriptionEdit, HelpFileEdit, CompilationEdit;
            internal string Name, Description, HelpFile, ContextText, Compilation;
            internal string OptionsVersion
            {
                get
                {
                    // Length-prefixed fields avoid ambiguous concatenation and exclude transient handles.
                    string text = Pack(Name) + Pack(Description) + Pack(HelpFile) + Pack(ContextText) + Pack(Compilation);
                    using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(text))).Replace("-", "").ToLowerInvariant();
                }
            }
            private static string Pack(string text) => (text ?? "").Length.ToString(CultureInfo.InvariantCulture) + ":" + (text ?? "");
            internal bool SameNative(Snapshot other) => other != null && Dialog == other.Dialog && Page == other.Page && Tab == other.Tab && Context == other.Context && NameEdit == other.NameEdit && DescriptionEdit == other.DescriptionEdit && HelpFileEdit == other.HelpFileEdit && CompilationEdit == other.CompilationEdit;
            internal bool UnchangedExceptContext(Snapshot other) => SameNative(other) && Name == other.Name && Description == other.Description && HelpFile == other.HelpFile && Compilation == other.Compilation;
            internal bool UnchangedExceptHelpFile(Snapshot other) => SameNative(other) && Name == other.Name && Description == other.Description && ContextText == other.ContextText && Compilation == other.Compilation;
        }
        internal sealed class Result
        {
            public bool Available, MutationInvoked, Uncertain, ControlValueVerified, CommittedRequested, DialogClosed, CommandEntered, OriginalExecuteReturned, Terminal, RefusedBeforeWrite;
            public bool PersistenceVerified => false;
            public bool RetryAllowed => false;
            public string OptionsVersion, Name, Description, HelpFile, HelpContextText, ConditionalCompilation, Error;
            public int OpenAttempts, FieldAttempts, OkAttempts, CancelAttempts;
        }
        internal interface INative
        {
            void RequireOwner();
            void Prepare(); // Original root enabled, no preexisting owned visible modal.
            Snapshot Capture(string exactProjectName); // Null only if no owned modal; unknown shapes throw.
            void RequireSame(Snapshot expected, bool compareValues);
            void RequireHelpFileRepresentable(Snapshot expected, string value);
            void WriteContext(Snapshot expected, int value, Action beforeEntry);
            void WriteHelpFile(Snapshot expected, string value, Action beforeEntry);
            void Close(Snapshot expected, int buttonId, Action beforeEnqueue);
            bool Closed(Snapshot expected);
        }
        internal interface IScheduler
        {
            void RequireOwner();
            long ElapsedMilliseconds { get; }
            void Post(Action callback);
            IDisposable Poll(Action callback);
        }
        private sealed class OwnerScheduler : IScheduler
        {
            private readonly SynchronizationContext context;
            private readonly int owner;
            private readonly Stopwatch clock = Stopwatch.StartNew();
            internal OwnerScheduler()
            {
                context = SynchronizationContext.Current;
                owner = Thread.CurrentThread.ManagedThreadId;
                if (context == null) throw new InvalidOperationException("Original VBE UI context required.");
                RequireOwner();
            }
            public void RequireOwner()
            {
                if (Thread.CurrentThread.ManagedThreadId != owner || Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                    throw new InvalidOperationException("General scheduler left its original VBE UI STA.");
            }
            public long ElapsedMilliseconds => clock.ElapsedMilliseconds;
            public void Post(Action callback) { RequireOwner(); context.Post(_ => callback(), null); }
            public IDisposable Poll(Action callback)
            {
                RequireOwner(); var timer = new System.Windows.Forms.Timer { Interval = 50 };
                timer.Tick += (sender, args) => callback(); timer.Start(); return timer;
            }
        }
        private readonly INative native;
        private readonly IScheduler scheduler;
        private bool consumed;
        internal VbeProjectGeneralOperation(INative native, IScheduler scheduler = null)
        {
            this.native = native ?? throw new ArgumentNullException(nameof(native)); this.scheduler = scheduler;
        }

        // All callbacks run on the original VBE UI STA. Full authorization includes live COM reads;
        // pure authorization must perform no COM/host reads and is repeated after native getters.
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
