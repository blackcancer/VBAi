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
    public sealed class CrashReportTests
    {
        [TestMethod]
        public void TechnicalReportExcludesExceptionMessagesPathsAndDataButKeepsUserDescription()
        {
            var error = new InvalidOperationException(@"secret-token C:\private\macro.swp VBA code", new Exception("private conversation"));
            error.Data["key"] = "password";
            var report = new CrashReport(error);
            string body = report.Body("Reproduction", "User-approved description");
            StringAssert.Contains(body, "User-approved description");
            StringAssert.Contains(body, "System.InvalidOperationException");
            StringAssert.Contains(body, report.Id);
            foreach (string secret in new[] { "secret-token", "private conversation", @"C:\private", "password", "VBA code" }) Assert.IsFalse(body.Contains(secret), secret);
            Assert.ThrowsException<ArgumentException>(() => report.Body(" ", ""));
            Assert.ThrowsException<ArgumentException>(() => report.Body(new string('x', 181), ""));
            Assert.ThrowsException<ArgumentException>(() => report.Body("Title", new string('x', 8001)));
            Assert.IsFalse(CrashReport.IsOwned(error));
            Assert.IsTrue(CrashReport.IsOwned(OwnedError()));
        }
        internal static Exception OwnedError()
        {
            try { new CrashReport().Body("", ""); }
            catch (Exception error) { return error; }
            throw new AssertFailedException();
        }
        [TestMethod]
        public void PendingSnapshotRoundTripsIdentityAndRejectsMismatchedOrOversizedFiles()
        {
            string directory = Path.Combine(Path.GetTempPath(), "vbai-crash-" + Guid.NewGuid().ToString("N"));
            try
            {
                var report = new CrashReport(OwnedError()); report.SavePending(directory);
                string path = Path.Combine(directory, report.Id + ".pending");
                var restored = CrashReport.ReadPending(path);
                Assert.AreEqual(report.Id, restored.Id); Assert.AreEqual(report.TechnicalDetails, restored.TechnicalDetails);
                Assert.IsTrue(File.Exists(Path.Combine(directory, report.Id + ".md")));
                string invalid = Path.Combine(directory, "wrong.pending"); File.Copy(path, invalid);
                Assert.ThrowsException<InvalidDataException>(() => CrashReport.ReadPending(invalid));
                File.WriteAllText(path, new string('x', 50001));
                Assert.ThrowsException<InvalidDataException>(() => CrashReport.ReadPending(path));
            }
            finally { if (Directory.Exists(directory)) Directory.Delete(directory, true); }
        }
    }
}
