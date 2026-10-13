using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;

namespace VBAi
{
    /// <summary>Identifies native test stages without recording source, arguments or project paths.</summary>
    internal enum NativeTestTimingStage
    {
        /// <summary>Live source and reference acquisition.</summary>
        ProjectSnapshot,
        /// <summary>Initial revision and execution-policy checks.</summary>
        ValidateBeforePreparation,
        /// <summary>Native pane and command preparation.</summary>
        PrepareNativeDispatch,
        /// <summary>Authority check after native pane selection.</summary>
        ValidateAfterPreparation,
        /// <summary>Authority check after focus preparation.</summary>
        ValidateBeforeArming,
        /// <summary>Final native selection and command revalidation.</summary>
        RevalidateNativeDispatch,
        /// <summary>Correlation receipt armed before dispatch.</summary>
        ArmCallback,
        /// <summary>Temporary runtime callback exposure.</summary>
        ExposeCallback,
        /// <summary>Single native dispatch and its final selection checks.</summary>
        ExecuteNativeDispatch,
        /// <summary>Current project mode read.</summary>
        ReadNativeMode,
        /// <summary>Correlated callback result completion.</summary>
        CompleteCallback,
        /// <summary>Temporary callback exposure revocation.</summary>
        RevokeCallback,
        /// <summary>Live document association and canonical host-path read within a project snapshot.</summary>
        SnapshotHostPath,
        /// <summary>Current component names, types and complete code-module source reads.</summary>
        SnapshotModules,
        /// <summary>Current reference names, GUIDs, versions and broken-reference flags.</summary>
        SnapshotReferences
    }

    /// <summary>Records opt-in stage timings while preserving each action and its original exception.</summary>
    internal sealed class NativeExecutionTimings
    {
        /// <summary>Opt-in directory; the caller must create it before starting the owned host.</summary>
        internal const string RootEnvironmentVariable = "VBAi_TEST_NATIVE_EXECUTION_TIMINGS";

        /// <summary>Bounds diagnostic output independently of the number of native operations.</summary>
        internal const int MaximumRecords = 8192;

        /// <summary>Disabled observer used by direct execution fixtures and ordinary product sessions.</summary>
        internal static readonly NativeExecutionTimings Disabled = new NativeExecutionTimings(null);

        /// <summary>Serializes event publication and enforces the diagnostic output limit.</summary>
        private readonly object gate = new object();

        /// <summary>Sink cleared after a diagnostic failure; it never executes the observed native action.</summary>
        private Action<StageEvent> publish;

        /// <summary>Monotonic span identifier distinguishes overlapping or nested measurements.</summary>
        private long spans;

        /// <summary>Number of diagnostic events successfully published.</summary>
        private int records;

        /// <summary>Distinguishes recorders sharing a process, including temporary coverage services.</summary>
        private readonly string observerId = Guid.NewGuid().ToString("N");

        /// <summary>Gets the first diagnostic failure without replacing a native operation failure.</summary>
        internal Exception DiagnosticFailure { get; private set; }

        /// <summary>Gets whether the configured event bound was reached.</summary>
        internal bool Truncated { get; private set; }

        /// <summary>Creates an observer with an optional event sink; null disables measurement.</summary>
        /// <param name="publish">Diagnostic sink; its failures are retained and disable further publication.</param>
        internal NativeExecutionTimings(Action<StageEvent> publish) { this.publish = publish; }

        /// <summary>Creates a recorder only for an explicitly supplied, existing diagnostic directory.</summary>
        /// <returns>Opt-in recorder, or a disabled observer when no directory was configured.</returns>
        internal static NativeExecutionTimings ForCurrentProcess() => ForRoot(Environment.GetEnvironmentVariable(RootEnvironmentVariable));

        /// <summary>Validates a caller-owned directory and selects one diagnostic file for this process.</summary>
        /// <param name="root">Absolute existing directory, or null/empty to disable diagnostics.</param>
        /// <returns>Recorder that writes metadata only, without creating directories or changing host settings.</returns>
        internal static NativeExecutionTimings ForRoot(string root)
        {
            if (string.IsNullOrWhiteSpace(root)) return Disabled;
            try
            {
                if (!Path.IsPathRooted(root) || !Directory.Exists(root))
                    throw new ArgumentException("Native timing diagnostics require an existing absolute directory.");
                int processId;
                using (var process = Process.GetCurrentProcess()) processId = process.Id;
                string path = Path.Combine(Path.GetFullPath(root), "native-stages-" + processId + ".jsonl");
                var json = new JavaScriptSerializer();
                string mvid = typeof(NativeExecutionTimings).Module.ModuleVersionId.ToString("D");
                return new NativeExecutionTimings(value => File.AppendAllText(path,
                    json.Serialize(new { Version = 1, ProcessId = processId, ProductMvid = mvid, Event = value }) + "\n",
                    new UTF8Encoding(false)));
            }
            catch (Exception error) { return new NativeExecutionTimings(null) { DiagnosticFailure = error }; }
        }

        /// <summary>Executes an action exactly once and records its entry and terminal outcome when enabled.</summary>
        /// <param name="stage">Fixed stage identifier, never user-provided text.</param>
        /// <param name="action">Original operation, whose return or exception remains authoritative.</param>
        internal void Measure(NativeTestTimingStage stage, Action action)
        {
            if (action == null) throw new ArgumentNullException(nameof(action));
            Measure<object>(stage, () => { action(); return null; });
        }

        /// <summary>Measures one operation while preserving its result, thread and exception identity.</summary>
        /// <typeparam name="T">Original operation result type.</typeparam>
        /// <param name="stage">Fixed stage identifier.</param>
        /// <param name="action">Original operation; it is never retried after a diagnostic or native failure.</param>
        /// <returns>The exact original result.</returns>
        internal T Measure<T>(NativeTestTimingStage stage, Func<T> action)
        {
            if (action == null) throw new ArgumentNullException(nameof(action));
            if (publish == null) return action();
            long span = Interlocked.Increment(ref spans);
            var clock = Stopwatch.StartNew();
            Emit(stage, span, "Started", 0, null);
            Exception failure = null;
            try { return action(); }
            catch (Exception error) { failure = error; throw; }
            finally { clock.Stop(); Emit(stage, span, failure == null ? "Returned" : "Failed", clock.Elapsed.TotalMilliseconds, failure?.GetType().FullName); }
        }

        /// <summary>Publishes bounded metadata; observer failures disable recording rather than native execution.</summary>
        /// <param name="stage">Fixed stage identifier.</param>
        /// <param name="span">Correlation identifier allocated before the operation.</param>
        /// <param name="state">Entry or terminal observation.</param>
        /// <param name="milliseconds">Elapsed stage duration, including diagnostic entry publication.</param>
        /// <param name="exceptionType">Exception type only; messages and user data are omitted.</param>
        private void Emit(NativeTestTimingStage stage, long span, string state, double milliseconds, string exceptionType)
        {
            lock (gate)
            {
                if (publish == null) return;
                if (records >= MaximumRecords) { Truncated = true; publish = null; return; }
                try
                {
                    publish(new StageEvent
                    {
                        ObserverId = observerId, Sequence = records + 1, Span = span, Stage = stage.ToString(), State = state,
                        Utc = DateTime.UtcNow.ToString("o"), Milliseconds = milliseconds,
                        ManagedThreadId = Thread.CurrentThread.ManagedThreadId,
                        Apartment = Thread.CurrentThread.GetApartmentState().ToString(), ExceptionType = exceptionType
                    });
                    records++;
                }
                catch (Exception error) { DiagnosticFailure = error; publish = null; }
            }
        }

        /// <summary>Contains stage metadata only; it excludes user code and native method arguments.</summary>
        internal sealed class StageEvent
        {
            /// <summary>Gets or sets the monotonically increasing event sequence.</summary>
            public int Sequence { get; set; }
            /// <summary>Gets or sets the recorder identity shared by entry and completion.</summary>
            public string ObserverId { get; set; }
            /// <summary>Gets or sets the operation span within this recorder.</summary>
            public long Span { get; set; }
            /// <summary>Gets or sets the fixed diagnostic stage name.</summary>
            public string Stage { get; set; }
            /// <summary>Gets or sets entry, successful return or failure state.</summary>
            public string State { get; set; }
            /// <summary>Gets or sets the UTC observation timestamp.</summary>
            public string Utc { get; set; }
            /// <summary>Gets or sets the measured stage duration.</summary>
            public double Milliseconds { get; set; }
            /// <summary>Gets or sets the original caller's managed thread identity.</summary>
            public int ManagedThreadId { get; set; }
            /// <summary>Gets or sets the original caller's apartment.</summary>
            public string Apartment { get; set; }
            /// <summary>Gets or sets an exception type, excluding its message and payload.</summary>
            public string ExceptionType { get; set; }
        }
    }
}
