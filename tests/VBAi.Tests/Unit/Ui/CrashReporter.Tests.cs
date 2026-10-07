using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Threading.Tasks;
namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class CrashReporterTests
    {
        [TestMethod]
        public void ReporterHandlesSaveDisplayReviewDirectoryAndRuntimeEventFailuresWithoutRecursing()
        {
            string directory = Path.Combine(Path.GetTempPath(), "VBAi-crash-contract-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(directory);
            try
            {
                var owned = new NullReferenceException("Synthetic failure", CrashReportTests.OwnedError());
                using (var reporter = new CrashReporter(report => { throw new IOException("Owned display unavailable"); }, report => { throw new IOException("Owned save unavailable"); }, directory)) reporter.ReportUnexpected(owned);
                int saved = 0;
                using (var reporter = new CrashReporter(report => { }, report => { saved++; Directory.CreateDirectory(Path.Combine(directory, report.Id + ".pending")); }, directory)) reporter.Capture(owned, true);
                Assert.AreEqual(1, saved);
                using (var reporter = new CrashReporter(report => { }, report => { saved++; }, directory))
                {
                    reporter.ReadPendingFiles = path => { throw new DirectoryNotFoundException("Directory removed during listing"); };
                    reporter.RecoverPending(directory);
                    var flags = System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic;
                    typeof(CrashReporter).GetMethod("FatalError", flags).Invoke(reporter, new object[] { null, new UnhandledExceptionEventArgs(owned, true) });
                    typeof(CrashReporter).GetMethod("FatalError", flags).Invoke(reporter, new object[] { null, new UnhandledExceptionEventArgs("nonexception", false) });
                    typeof(CrashReporter).GetMethod("TaskError", flags).Invoke(reporter, new object[] { null, new UnobservedTaskExceptionEventArgs(new AggregateException(owned)) });
                    Assert.AreEqual(3, saved);
                }
            }
            finally { Directory.Delete(directory, true); }
        }
        [TestMethod]
        public void CaptureKeepsIdentityPreventsRecursionAndIgnoresForeignErrorsAndDisposedReporter()
        {
            int saved = 0, shown = 0; CrashReport stored = null; CrashReporter reporter = null;
            reporter = new CrashReporter(report => { shown++; Assert.AreSame(stored, report); reporter.Capture(CrashReportTests.OwnedError(), true); }, report => { saved++; stored = report; });
            using (reporter)
            {
                reporter.Capture(new Exception("Foreign error"), true); Assert.AreEqual(0, shown);
                reporter.ReportUnexpected(CrashReportTests.OwnedError()); Assert.AreEqual(0, shown);
                reporter.Capture(CrashReportTests.OwnedError(), true); Assert.AreEqual(1, saved); Assert.AreEqual(1, shown);
                reporter.Capture(CrashReportTests.OwnedError(), false); Assert.AreEqual(2, saved); Assert.AreEqual(1, shown);
            }
            reporter.Dispose(); reporter.Capture(CrashReportTests.OwnedError(), true); Assert.AreEqual(2, saved);
        }
        [TestMethod]
        public void RecoveryShowsPendingReportOnceAndPreservesMarkdownBackup()
        {
            string directory = Path.Combine(Path.GetTempPath(), "vbai-recovery-" + Guid.NewGuid().ToString("N"));
            try
            {
                var report = new CrashReport(); report.SavePending(directory); int shown = 0;
                using (var reporter = new CrashReporter(restored => { shown++; Assert.AreEqual(report.Id, restored.Id); }))
                {
                    reporter.RecoverPending(directory); reporter.RecoverPending(directory);
                    Assert.AreEqual(1, shown); Assert.IsTrue(File.Exists(Path.Combine(directory, report.Id + ".md")));
                    report.SavePending(directory);
                }
                using (var reporter = new CrashReporter(restored => throw new InvalidOperationException("Window unavailable"))) reporter.RecoverPending(directory);
                Assert.IsTrue(File.Exists(Path.Combine(directory, report.Id + ".pending")));
            }
            finally { if (Directory.Exists(directory)) Directory.Delete(directory, true); }
        }
    }
}
