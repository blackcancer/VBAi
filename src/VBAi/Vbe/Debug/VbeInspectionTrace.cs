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
        internal const string EnvironmentName = "VBAi_VBE_INSPECTION_TRACE";
        internal const int MaximumEvents = 128;
        internal const long MaximumFileBytes = 1024 * 1024;
        private static readonly AsyncLocal<VbeInspectionTrace> ambient = new AsyncLocal<VbeInspectionTrace>();
        private static readonly object fileGate = new object();
        private readonly Action<string> write;
        private readonly Stopwatch clock = Stopwatch.StartNew();
        private readonly string correlation = Guid.NewGuid().ToString("N");
        private int count;
        internal static VbeInspectionTrace Current => ambient.Value;

        /// <summary>Fixed phase vocabulary prevents request content from becoming diagnostic text.</summary>
        internal enum Phase { Enqueue, CallbackEntered, OwnerSta, ContextValidation, ContextValidated,
            Command229Before, Command229Returned, ObserverEntered, ObserverDialogFound, ObserverReadComplete,
            ObserverTerminal, ContinuationEnqueued, ContinuationEntered, ContinuationReturned,
            CoreEntered, CoreTerminal, Terminal, OptionsComboInspection }

        internal enum OptionsReader { NativeCombo, UiAutomationCombo }
        internal enum OptionsRole { Other, Font, Size, Palette }

        /// <summary>Numeric native observations only; no labels, values or request text.</summary>
        internal sealed class OptionsComboEvidence
        {
            public OptionsReader Reader;
            public OptionsRole Role;
            public long Window;
            public long? Parent;
            public uint? OwnerProcessId;
            public uint? OwnerThreadId;
            public int? ControlId;
            public int? Style;
            public int? CountBefore;
            public int? CountAfterExpansion;
            public int? SelectedIndex;
            public bool? DropDownBefore;
            public bool? DropDownAfterExpansion;
            public bool? DropDownAfterCleanup;
            public bool ExpansionAttempted;
            public bool ReadCompleted;
        }

        internal VbeInspectionTrace(Action<string> writer) { write = writer; }

        /// <summary>Captures the opt-in destination once; invalid or unavailable logging never changes execution.</summary>
        internal static VbeInspectionTrace Begin()
        {
            try { return ForPath(Environment.GetEnvironmentVariable(EnvironmentName)); }
            catch { return null; }
        }

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

        internal IDisposable Enter()
        {
            var previous = ambient.Value;
            ambient.Value = this;
            return new Scope(() => ambient.Value = previous);
        }

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

        private sealed class Scope : IDisposable
        {
            private Action restore;
            internal Scope(Action action) { restore = action; }
            public void Dispose() { Interlocked.Exchange(ref restore, null)?.Invoke(); }
        }
    }
}
