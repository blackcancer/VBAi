using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;

namespace VBAi
{

    /// <summary>Optional bounded phase evidence, containing no request, source or inspected values.</summary>
    internal sealed class VbeInspectionTrace
    {

        /// <summary>Environment variable that opts into writing bounded VBE inspection diagnostics.</summary>
        internal const string EnvironmentName = "VBAi_VBE_INSPECTION_TRACE";

        /// <summary>Maximum trace rows emitted by one trace instance.</summary>
        internal const int MaximumEvents = 128;

        /// <summary>Maximum bytes appended to one trace destination.</summary>
        internal const long MaximumFileBytes = 1024 * 1024;

        /// <summary>Async-local trace scope used to attach phase events to the current inspection flow.</summary>
        private static readonly AsyncLocal<VbeInspectionTrace> ambient = new AsyncLocal<VbeInspectionTrace>();

        /// <summary>Serializes appends from concurrent trace scopes targeting the same file.</summary>
        private static readonly object fileGate = new object();

        /// <summary>Optional line writer; invalid opt-in paths disable tracing by leaving it absent.</summary>
        private readonly Action<string> write;

        /// <summary>Monotonic clock used to timestamp relative phase latency.</summary>
        private readonly Stopwatch clock = Stopwatch.StartNew();

        /// <summary>Opaque per-instance correlation identifier containing no request or inspected values.</summary>
        private readonly string correlation = Guid.NewGuid().ToString("N");

        /// <summary>Counts the count maintained by vbe inspection trace.</summary>
        private int count;

        /// <summary>Gets the current.</summary>
        /// <value>Current current exposed by vbe inspection trace.</value>
        internal static VbeInspectionTrace Current => ambient.Value;

        /// <summary>Fixed phase vocabulary prevents request content from becoming diagnostic text.</summary>
        internal enum Phase
        {

            /// <summary>Identifies the enqueue case of phase.</summary>
            Enqueue,

            /// <summary>Identifies the callback entered case of phase.</summary>
            CallbackEntered,

            /// <summary>Identifies the owner sta case of phase.</summary>
            OwnerSta,

            /// <summary>Identifies the context validation case of phase.</summary>
            ContextValidation,

            /// <summary>Identifies the context validated case of phase.</summary>
            ContextValidated,

            /// <summary>Identifies the command229 before case of phase.</summary>
            Command229Before,

            /// <summary>Identifies the command229 returned case of phase.</summary>
            Command229Returned,

            /// <summary>Identifies the observer entered case of phase.</summary>
            ObserverEntered,

            /// <summary>Identifies the observer dialog found case of phase.</summary>
            ObserverDialogFound,

            /// <summary>Identifies the observer read complete case of phase.</summary>
            ObserverReadComplete,

            /// <summary>Identifies the observer terminal case of phase.</summary>
            ObserverTerminal,

            /// <summary>Identifies the continuation enqueued case of phase.</summary>
            ContinuationEnqueued,

            /// <summary>Identifies the continuation entered case of phase.</summary>
            ContinuationEntered,

            /// <summary>Identifies the continuation returned case of phase.</summary>
            ContinuationReturned,

            /// <summary>Identifies the core entered case of phase.</summary>
            CoreEntered,

            /// <summary>Identifies the core terminal case of phase.</summary>
            CoreTerminal,

            /// <summary>Identifies the terminal case of phase.</summary>
            Terminal,

            /// <summary>Identifies the options combo inspection case of phase.</summary>
            OptionsComboInspection,

            /// <summary>Identifies the continuation post returned case of phase.</summary>
            ContinuationPostReturned,

            /// <summary>Identifies the continuation post failed case of phase.</summary>
            ContinuationPostFailed
        }

        /// <summary>Lists the supported options reader values.</summary>
        internal enum OptionsReader
        {

            /// <summary>Identifies the native combo case of options reader.</summary>
            NativeCombo,

            /// <summary>Identifies the ui automation combo case of options reader.</summary>
            UiAutomationCombo
        }

        /// <summary>Lists the supported options role values.</summary>
        internal enum OptionsRole
        {

            /// <summary>Identifies the other case of options role.</summary>
            Other,

            /// <summary>Identifies the font case of options role.</summary>
            Font,

            /// <summary>Identifies the size case of options role.</summary>
            Size,

            /// <summary>Identifies the palette case of options role.</summary>
            Palette
        }

        /// <summary>Numeric native observations only; no labels, values or request text.</summary>
        internal sealed class OptionsComboEvidence
        {

            /// <summary>Reader implementation used for the bounded combo inspection.</summary>
            public OptionsReader Reader;

            /// <summary>Semantic role assigned to the observed options combo.</summary>
            public OptionsRole Role;

            /// <summary>Native HWND represented as a signed pointer-sized integer.</summary>
            public long Window;

            /// <summary>Parent HWND, when the reader could retrieve it.</summary>
            public long? Parent;

            /// <summary>Native owner process ID, when available.</summary>
            public uint? OwnerProcessId;

            /// <summary>Native owner UI-thread ID, when available.</summary>
            public uint? OwnerThreadId;

            /// <summary>Native dialog control ID, when available.</summary>
            public int? ControlId;

            /// <summary>Numeric Win32 style bits observed on the combo.</summary>
            public int? Style;

            /// <summary>Item count read before focus or expansion attempts.</summary>
            public int? CountBefore;

            /// <summary>Item count after the focus attempt.</summary>
            public int? CountAfterFocus;

            /// <summary>Item count after the dropdown expansion attempt.</summary>
            public int? CountAfterExpansion;

            /// <summary>Selected item index when the native reader could obtain it.</summary>
            public int? SelectedIndex;

            /// <summary>Dropdown visibility observed before expansion.</summary>
            public bool? DropDownBefore;

            /// <summary>Dropdown visibility observed after expansion.</summary>
            public bool? DropDownAfterExpansion;

            /// <summary>Dropdown visibility observed after cleanup attempts.</summary>
            public bool? DropDownAfterCleanup;

            /// <summary>Whether the reader attempted to expand the native dropdown.</summary>
            public bool ExpansionAttempted;

            /// <summary>Whether the reader attempted to focus the native combo.</summary>
            public bool FocusAttempted;

            /// <summary>Whether all planned numeric observations completed successfully.</summary>
            public bool ReadCompleted;
        }

        /// <summary>Creates a trace instance with a line writer and fresh correlation/timing state.</summary>
        /// <param name="writer">Optional callback that appends one serialized JSON event line.</param>
        internal VbeInspectionTrace(Action<string> writer) { write = writer; }

        /// <summary>Captures the opt-in destination once; invalid or unavailable logging never changes execution.</summary>
        /// <returns>vbe inspection trace produced by the operation for begin on vbe inspection trace.</returns>
        internal static VbeInspectionTrace Begin()
        {
            try { return ForPath(Environment.GetEnvironmentVariable(EnvironmentName)); }
            catch { return null; }
        }

        /// <summary>Creates an opt-in writer only for a local fixed/removable/RAM drive path without reparse ancestors.</summary>
        /// <param name="path">Absolute drive-letter destination; UNC paths, invalid paths, and reparse paths disable tracing.</param>
        /// <returns>Bounded append trace, or null when the destination is outside the accepted local path rules.</returns>
        internal static VbeInspectionTrace ForPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || path.Length < 3 || !char.IsLetter(path[0]) || path[1] != ':' ||
                (path[2] != '\\' && path[2] != '/') || !Path.IsPathRooted(path) || path.StartsWith(@"\\", StringComparison.Ordinal) ||
                path.IndexOf(':', 2) >= 0) return null;
            string destination = Path.GetFullPath(path);
            var drive = new DriveInfo(Path.GetPathRoot(destination));
            if (drive.DriveType != DriveType.Fixed && drive.DriveType != DriveType.Removable && drive.DriveType != DriveType.Ram) return null;
            for (string ancestor = destination; !string.IsNullOrEmpty(ancestor); ancestor = Path.GetDirectoryName(ancestor))
                if ((File.Exists(ancestor) || Directory.Exists(ancestor)) && (File.GetAttributes(ancestor) & FileAttributes.ReparsePoint) != 0) return null;
            return new VbeInspectionTrace(line =>
            {
                lock (fileGate)
                {
                    byte[] bytes = new UTF8Encoding(false).GetBytes(line + "\n");
                    using (var file = new FileStream(destination, FileMode.OpenOrCreate, FileAccess.Write, FileShare.Read))
                    {
                        if (file.Length + bytes.Length > MaximumFileBytes) return;
                        file.Position = file.Length;
                        file.Write(bytes, 0, bytes.Length);
                        file.Flush();
                    }
                }
            });
        }

        /// <summary>Sets this instance as the ambient trace until the returned scope is disposed.</summary>
        /// <returns>Scope that restores the previous ambient trace, supporting nested inspections.</returns>
        internal IDisposable Enter()
        {
            var previous = ambient.Value;
            ambient.Value = this;
            return new Scope(() => ambient.Value = previous);
        }

        /// <summary>Appends a fixed-vocabulary phase event without request or inspected values.</summary>
        /// <param name="phase">Known lifecycle phase to record; unknown enum values and events beyond the cap are ignored.</param>
        /// <param name="error">Exception describing the error failure.</param>
        internal void Record(Phase phase, Exception error = null)
        {
            try
            {
                int sequence = Interlocked.Increment(ref count);
                if (sequence > MaximumEvents || !Enum.IsDefined(typeof(Phase), phase)) return;
                var row = new
                {
                    Correlation = correlation,
                    Sequence = sequence,
                    Utc = DateTime.UtcNow.ToString("o"),
                    HostProcessId = Process.GetCurrentProcess().Id,
                    ThreadId = Thread.CurrentThread.ManagedThreadId,
                    Apartment = Thread.CurrentThread.GetApartmentState().ToString(),
                    clock.ElapsedMilliseconds,
                    Phase = phase.ToString(),
                    ErrorType = error?.GetType().Name
                };
                write?.Invoke(new JavaScriptSerializer().Serialize(row));
            }
            catch { /* Evidence is optional and must not affect a native operation or its original failure. */ }
        }

        /// <summary>Appends numeric combo evidence after validating its reader and role enums.</summary>
        /// <param name="observed">Numeric HWND/style/count/selection observations; text labels and option values are not represented.</param>
        /// <param name="error">Exception describing the error failure.</param>
        internal void RecordOptionsCombo(OptionsComboEvidence observed, Exception error = null)
        {
            try
            {
                int sequence = Interlocked.Increment(ref count);
                if (sequence > MaximumEvents || observed == null ||
                    !Enum.IsDefined(typeof(OptionsReader), observed.Reader) ||
                    !Enum.IsDefined(typeof(OptionsRole), observed.Role)) return;
                var row = new
                {
                    Correlation = correlation,
                    Sequence = sequence,
                    Utc = DateTime.UtcNow.ToString("o"),
                    HostProcessId = Process.GetCurrentProcess().Id,
                    ThreadId = Thread.CurrentThread.ManagedThreadId,
                    Apartment = Thread.CurrentThread.GetApartmentState().ToString(),
                    clock.ElapsedMilliseconds,
                    Phase = Phase.OptionsComboInspection.ToString(),
                    ErrorType = error?.GetType().Name,
                    Reader = observed.Reader.ToString(),
                    Role = observed.Role.ToString(),
                    Native = observed
                };
                write?.Invoke(new JavaScriptSerializer().Serialize(row));
            }
            catch { /* Diagnostics cannot alter a native read, closure or original exception. */ }
        }

        /// <summary>Restores the prior AsyncLocal trace when an inspection scope ends.</summary>
        private sealed class Scope : IDisposable
        {

            /// <summary>One-use restoration callback, cleared atomically during disposal.</summary>
            private Action restore;

            /// <summary>Creates a trace scope with a prior-context restoration callback.</summary>
            /// <param name="action">Callback restoring the ambient trace that was active before entry.</param>
            internal Scope(Action action) { restore = action; }

            /// <summary>Disposes  for scope.</summary>
            public void Dispose() { Interlocked.Exchange(ref restore, null)?.Invoke(); }
        }
    }
}
