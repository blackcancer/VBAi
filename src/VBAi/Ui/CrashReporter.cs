using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.IO;
using System.Threading.Tasks;

namespace VBAi
{

    /// <summary>Capture les erreurs du complément. Les événements fatals n’ouvrent aucune UI pendant la terminaison de l’hôte.</summary>
    internal sealed class CrashReporter : IDisposable
    {

        /// <summary>Callback that displays a report to the user.</summary>
        private readonly Action<CrashReport> show;

        /// <summary>Callback that persists a report before optional display.</summary>
        private readonly Action<CrashReport> save;

        /// <summary>Directory used by default for pending crash reports.</summary>
        private readonly string directory;

        /// <summary>Interlocked guard against overlapping report capture and recovery.</summary>
        private int reporting;

        /// <summary>Whether this reporter has removed its global exception subscriptions.</summary>
        private bool disposed;

        /// <summary>Enumerates pending reports in the owned directory with the native filesystem by default.</summary>
        internal Func<string, string[]> ReadPendingFiles = folder => Directory.GetFiles(folder, "*.pending");

        /// <summary>Creates a reporter and subscribes to fatal and unobserved task exceptions.</summary>
        /// <param name="show">Callback used to show a report when interactive display is safe.</param>
        /// <param name="save">Optional persistence callback; defaults to writing a pending report.</param>
        /// <param name="directory">Optional report directory; defaults to <see cref="CrashReport.DirectoryPath"/>.</param>
        internal CrashReporter(Action<CrashReport> show, Action<CrashReport> save = null, string directory = null)
        {
            this.show = show;
            this.directory = directory ?? CrashReport.DirectoryPath;
            this.save = save ?? (report => report.SavePending(this.directory));
            AppDomain.CurrentDomain.UnhandledException += FatalError;
            TaskScheduler.UnobservedTaskException += TaskError;
        }

        /// <summary>Rapporte une erreur de programmation interceptée dans le complément ; les erreurs opérationnelles restent gérées par leur UI.</summary>
        /// <param name="error">Exception intercepted by the add-in.</param>
        internal void ReportUnexpected(Exception error)
        {
            if (error == null || error is ArgumentException || error is InvalidOperationException || error is COMException ||
                error is OperationCanceledException || !CrashReport.IsOwned(error)) return;
            Capture(error, true);
        }

        /// <summary>Persists an add-in exception report and optionally displays it when safe to do so.</summary>
        /// <param name="error">Exception to capture.</param><param name="display">Whether to invoke the interactive report callback.</param>
        internal void Capture(Exception error, bool display)
        {
            if (disposed || error == null || !CrashReport.IsOwned(error) || Interlocked.CompareExchange(ref reporting, 1, 0) != 0) return;
            try
            {
                var report = new CrashReport(error);
                try { save(report); }
                catch (Exception) { LoadLog.Write("Crash report could not be saved."); }
                if (display)
                    try { show(report); MarkReviewed(report); }
                    catch (Exception) { LoadLog.Write("Crash report window could not be opened."); }
            }
            finally { Volatile.Write(ref reporting, 0); }
        }

        /// <summary>Propose au démarrage suivant les rapports sauvegardés avant une terminaison fatale.</summary>
        /// <param name="directory">Directory to scan, or null to use this reporter's configured directory.</param>
        internal void RecoverPending(string directory = null)
        {
            directory = directory ?? this.directory;
            if (disposed || !Directory.Exists(directory) || Interlocked.CompareExchange(ref reporting, 1, 0) != 0) return;
            try
            {
                foreach (string path in ReadPendingFiles(directory))
                {
                    try
                    {
                        var report = CrashReport.ReadPending(path);
                        show(report);
                        File.Delete(path); // Markdown backup remains available.
                        break; // One dialog per connection; remaining reports are preserved.
                    }
                    catch (Exception) { LoadLog.Write("Pending crash report could not be opened."); }
                }
            }
            catch (Exception) { LoadLog.Write("Pending crash report directory could not be read."); }
            finally { Volatile.Write(ref reporting, 0); }
        }

        /// <summary>Removes the pending marker after a report was shown; the Markdown backup remains.</summary>
        /// <param name="report">Report whose marker should be removed.</param>
        private void MarkReviewed(CrashReport report)
        {
            try { File.Delete(Path.Combine(directory, report.Id + ".pending")); }
            catch (Exception) { LoadLog.Write("Crash report review state could not be saved."); }
        }

        /// <summary>Captures an unhandled exception without displaying UI during host termination.</summary>
        /// <param name="sender">AppDomain that raised the event.</param><param name="args">Fatal exception event data.</param>
        private void FatalError(object sender, UnhandledExceptionEventArgs args) { Capture(args.ExceptionObject as Exception, false); }

        /// <summary>Captures an unobserved task exception without displaying UI from the finalizer path.</summary>
        /// <param name="sender">Task scheduler that raised the event.</param><param name="args">Unobserved task exception data.</param>
        private void TaskError(object sender, UnobservedTaskExceptionEventArgs args) { Capture(args.Exception, false); }

        /// <summary>Unsubscribes from process-wide exception events.</summary>
        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            AppDomain.CurrentDomain.UnhandledException -= FatalError;
            TaskScheduler.UnobservedTaskException -= TaskError;
        }
    }
}
