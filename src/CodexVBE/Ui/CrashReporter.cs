using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.IO;
using System.Threading.Tasks;

namespace CodexVBE
{
    /// <summary>Capture les erreurs du complément. Les événements fatals n’ouvrent aucune UI pendant la terminaison de l’hôte.</summary>
    internal sealed class CrashReporter : IDisposable
    {
        private readonly Action<CrashReport> show;
        private readonly Action<CrashReport> save;
        private readonly string directory;
        private int reporting;
        private bool disposed;
        /// <summary>Enumerates pending reports in the owned directory with the native filesystem by default.</summary>
        internal Func<string, string[]> ReadPendingFiles = folder => Directory.GetFiles(folder, "*.pending");
        internal CrashReporter(Action<CrashReport> show, Action<CrashReport> save = null, string directory = null)
        {
            this.show = show;
            this.directory = directory ?? CrashReport.DirectoryPath;
            this.save = save ?? (report => report.SavePending(this.directory));
            AppDomain.CurrentDomain.UnhandledException += FatalError;
            TaskScheduler.UnobservedTaskException += TaskError;
        }

        /// <summary>Rapporte une erreur de programmation interceptée dans le complément ; les erreurs opérationnelles restent gérées par leur UI.</summary>
        internal void ReportUnexpected(Exception error)
        {
            if (error == null || error is ArgumentException || error is InvalidOperationException || error is COMException ||
                error is OperationCanceledException || !CrashReport.IsOwned(error)) return;
            Capture(error, true);
        }

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
        private void MarkReviewed(CrashReport report)
        {
            try { File.Delete(Path.Combine(directory, report.Id + ".pending")); }
            catch (Exception) { LoadLog.Write("Crash report review state could not be saved."); }
        }
        private void FatalError(object sender, UnhandledExceptionEventArgs args) { Capture(args.ExceptionObject as Exception, false); }
        private void TaskError(object sender, UnobservedTaskExceptionEventArgs args) { Capture(args.Exception, false); }
        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            AppDomain.CurrentDomain.UnhandledException -= FatalError;
            TaskScheduler.UnobservedTaskException -= TaskError;
        }
    }
}
