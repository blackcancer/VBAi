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

        /// <summary>Maintains the environment name state for vbe inspection trace.</summary>
        internal const string EnvironmentName = "VBAi_VBE_INSPECTION_TRACE";

        /// <summary>Maintains the maximum events state for vbe inspection trace.</summary>
        internal const int MaximumEvents = 128;

        /// <summary>Maintains the maximum file bytes state for vbe inspection trace.</summary>
        internal const long MaximumFileBytes = 1024 * 1024;

        /// <summary>Maintains the ambient state for vbe inspection trace.</summary>
        private static readonly AsyncLocal<VbeInspectionTrace> ambient = new AsyncLocal<VbeInspectionTrace>();

        /// <summary>Maintains the file gate state for vbe inspection trace.</summary>
        private static readonly object fileGate = new object();

        /// <summary>Maintains the write state for vbe inspection trace.</summary>
        private readonly Action<string> write;

        /// <summary>Maintains the clock state for vbe inspection trace.</summary>
        private readonly Stopwatch clock = Stopwatch.StartNew();

        /// <summary>Maintains the correlation state for vbe inspection trace.</summary>
        private readonly string correlation = Guid.NewGuid().ToString("N");

        /// <summary>Counts the count maintained by vbe inspection trace.</summary>
        private int count;

        /// <summary>Gets the current.</summary>
        /// <value>Current current exposed by vbe inspection trace.</value>
        internal static VbeInspectionTrace Current => ambient.Value;

        /// <summary>Fixed phase vocabulary prevents request content from becoming diagnostic text.</summary>
        internal enum Phase {

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
ContinuationPostFailed }

        /// <summary>Lists the supported options reader values.</summary>
        internal enum OptionsReader {

/// <summary>Identifies the native combo case of options reader.</summary>
NativeCombo,

/// <summary>Identifies the ui automation combo case of options reader.</summary>
UiAutomationCombo }

        /// <summary>Lists the supported options role values.</summary>
        internal enum OptionsRole {

/// <summary>Identifies the other case of options role.</summary>
Other,

/// <summary>Identifies the font case of options role.</summary>
Font,

/// <summary>Identifies the size case of options role.</summary>
Size,

/// <summary>Identifies the palette case of options role.</summary>
Palette }

        /// <summary>Numeric native observations only; no labels, values or request text.</summary>
        internal sealed class OptionsComboEvidence
        {

            /// <summary>Maintains the reader state for options combo evidence.</summary>
            public OptionsReader Reader;

            /// <summary>Maintains the role state for options combo evidence.</summary>
            public OptionsRole Role;

            /// <summary>Maintains the window state for options combo evidence.</summary>
            public long Window;

            /// <summary>Maintains the parent state for options combo evidence.</summary>
            public long? Parent;

            /// <summary>Identifies the owner process id associated with options combo evidence.</summary>
            public uint? OwnerProcessId;

            /// <summary>Identifies the owner thread id associated with options combo evidence.</summary>
            public uint? OwnerThreadId;

            /// <summary>Identifies the control id associated with options combo evidence.</summary>
            public int? ControlId;

            /// <summary>Maintains the style state for options combo evidence.</summary>
            public int? Style;

            /// <summary>Maintains the count before state for options combo evidence.</summary>
            public int? CountBefore;

            /// <summary>Maintains the count after focus state for options combo evidence.</summary>
            public int? CountAfterFocus;

            /// <summary>Maintains the count after expansion state for options combo evidence.</summary>
            public int? CountAfterExpansion;

            /// <summary>Maintains the selected index state for options combo evidence.</summary>
            public int? SelectedIndex;

            /// <summary>Maintains the drop down before state for options combo evidence.</summary>
            public bool? DropDownBefore;

            /// <summary>Maintains the drop down after expansion state for options combo evidence.</summary>
            public bool? DropDownAfterExpansion;

            /// <summary>Maintains the drop down after cleanup state for options combo evidence.</summary>
            public bool? DropDownAfterCleanup;

            /// <summary>Maintains the expansion attempted state for options combo evidence.</summary>
            public bool ExpansionAttempted;

            /// <summary>Maintains the focus attempted state for options combo evidence.</summary>
            public bool FocusAttempted;

            /// <summary>Maintains the read completed state for options combo evidence.</summary>
            public bool ReadCompleted;
        }

        /// <summary>Initializes a VbeInspectionTrace instance with the supplied state.</summary>
        /// <param name="writer">action&lt;string&gt; that supplies the writer for this operation.</param>
        internal VbeInspectionTrace(Action<string> writer) { write = writer; }

        /// <summary>Captures the opt-in destination once; invalid or unavailable logging never changes execution.</summary>
        /// <returns>vbe inspection trace produced by the operation for begin on vbe inspection trace.</returns>
        internal static VbeInspectionTrace Begin()
        {
            try { return ForPath(Environment.GetEnvironmentVariable(EnvironmentName)); }
            catch { return null; }
        }

        /// <summary>Handles for path for vbe inspection trace.</summary>
        /// <param name="path">Path used for the path being processed.</param>
        /// <returns>vbe inspection trace produced by the operation for for path on vbe inspection trace.</returns>
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
            return new VbeInspectionTrace(line => {
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

        /// <summary>Handles enter for vbe inspection trace.</summary>
        /// <returns>i disposable produced by the operation for enter on vbe inspection trace.</returns>
        internal IDisposable Enter()
        {
            var previous = ambient.Value;
            ambient.Value = this;
            return new Scope(() => ambient.Value = previous);
        }

        /// <summary>Handles record for vbe inspection trace.</summary>
        /// <param name="phase">phase that supplies the phase for this operation.</param>
        /// <param name="error">Exception describing the error failure.</param>
        internal void Record(Phase phase, Exception error = null)
        {
            try
            {
                int sequence = Interlocked.Increment(ref count);
                if (sequence > MaximumEvents || !Enum.IsDefined(typeof(Phase), phase)) return;
                var row = new { Correlation = correlation, Sequence = sequence, Utc = DateTime.UtcNow.ToString("o"),
                    HostProcessId = Process.GetCurrentProcess().Id, ThreadId = Thread.CurrentThread.ManagedThreadId,
                    Apartment = Thread.CurrentThread.GetApartmentState().ToString(), ElapsedMilliseconds = clock.ElapsedMilliseconds,
                    Phase = phase.ToString(), ErrorType = error == null ? null : error.GetType().Name };
                write?.Invoke(new JavaScriptSerializer().Serialize(row));
            }
            catch { /* Evidence is optional and must not affect a native operation or its original failure. */ }
        }

        /// <summary>Handles record options combo for vbe inspection trace.</summary>
        /// <param name="observed">options combo evidence that supplies the observed for this operation.</param>
        /// <param name="error">Exception describing the error failure.</param>
        internal void RecordOptionsCombo(OptionsComboEvidence observed, Exception error = null)
        {
            try
            {
                int sequence = Interlocked.Increment(ref count);
                if (sequence > MaximumEvents || observed == null ||
                    !Enum.IsDefined(typeof(OptionsReader), observed.Reader) ||
                    !Enum.IsDefined(typeof(OptionsRole), observed.Role)) return;
                var row = new { Correlation = correlation, Sequence = sequence, Utc = DateTime.UtcNow.ToString("o"),
                    HostProcessId = Process.GetCurrentProcess().Id, ThreadId = Thread.CurrentThread.ManagedThreadId,
                    Apartment = Thread.CurrentThread.GetApartmentState().ToString(), ElapsedMilliseconds = clock.ElapsedMilliseconds,
                    Phase = Phase.OptionsComboInspection.ToString(), ErrorType = error?.GetType().Name,
                    Reader = observed.Reader.ToString(), Role = observed.Role.ToString(), Native = observed };
                write?.Invoke(new JavaScriptSerializer().Serialize(row));
            }
            catch { /* Diagnostics cannot alter a native read, closure or original exception. */ }
        }

        /// <summary>Owns the scope state and operations.</summary>
        private sealed class Scope : IDisposable
        {

            /// <summary>Maintains the restore state for scope.</summary>
            private Action restore;

            /// <summary>Initializes a Scope instance with the supplied state.</summary>
            /// <param name="action">action that supplies the action for this operation.</param>
            internal Scope(Action action) { restore = action; }

            /// <summary>Disposes  for scope.</summary>
            public void Dispose() { Interlocked.Exchange(ref restore, null)?.Invoke(); }
        }
    }
}
