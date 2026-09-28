using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class CrashReporterTests
    {
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
