using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using VBAi.Tests.Infrastructure;
namespace VBAi.Tests.Unit
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
        [TestMethod]
        public void ReportMetadataAndFrameGuardsRemainBoundedAndExcludePrivateContents()
        {
            using (var scope = new TechnicalUiMetadataScope())
            {
                CrashReport.MetadataAssembly = TechnicalUiMetadataScope.WithoutInformation; CrashReport.ProcessIs64Bit = () => false; CrashReport.RuntimeVersion = () => null;
                var dynamicMethod = new System.Reflection.Emit.DynamicMethod("OwnedDynamicFrame", typeof(void), Type.EmptyTypes);
                var ownedMethod = typeof(CrashReport).GetMethod("Body", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
                CrashReport.FrameSnapshot = error => new System.Diagnostics.StackFrame[] { new OwnedTechnicalStackFrame(null), new OwnedTechnicalStackFrame(dynamicMethod), new OwnedTechnicalStackFrame(ownedMethod) };
                Exception nested = new ArgumentException("private-sixth"); for (int i = 0; i < 5; i++) nested = new IOException("private-outer", nested);
                var report = new CrashReport(nested); StringAssert.Contains(report.TechnicalDetails, "5.6.7.8"); StringAssert.Contains(report.TechnicalDetails, "x86");
                Assert.AreEqual(5, report.TechnicalDetails.Split(new[] { "Exception:" }, StringSplitOptions.None).Length - 1);
                Assert.IsFalse(report.TechnicalDetails.Contains("ArgumentException")); Assert.IsFalse(report.TechnicalDetails.Contains("private")); StringAssert.Contains(report.TechnicalDetails, ".OwnedDynamicFrame");
                Assert.IsTrue(CrashReport.IsOwned(new AggregateException(new Exception("private"))));
                CrashReport.FrameSnapshot = error => new System.Diagnostics.StackFrame[] { new OwnedTechnicalStackFrame(null), new OwnedTechnicalStackFrame(dynamicMethod) };
                Assert.IsFalse(CrashReport.IsOwned(new AggregateException(new Exception("private")))); Assert.IsFalse(CrashReport.IsOwned(new AggregateException()));
                Assert.IsFalse(CrashReport.IsOwned(null)); Assert.IsNotNull(report.Body("Title", null));
                Assert.IsFalse(string.IsNullOrEmpty(CrashReport.DirectoryPath));
            }
        }

        [TestMethod]
        public void PendingSnapshotsRejectEveryIncompleteMetadataFieldBeforeConstruction()
        {
            using (var host = new EditorFixture())
            {
                Directory.CreateDirectory(host.Root);
                using (var scope = new TechnicalUiMetadataScope()) { CrashReport.ReportDirectory = () => host.Root; var report = new CrashReport(); Assert.AreEqual(Path.Combine(host.Root, report.Id + ".md"), report.Save("owned body")); report.SavePending(); Assert.IsTrue(File.Exists(Path.Combine(host.Root, report.Id + ".pending"))); }
                string id = Guid.NewGuid().ToString("N"), path = Path.Combine(host.Root, id + ".pending");
                var json = new System.Web.Script.Serialization.JavaScriptSerializer();
                foreach (string state in new[] { "null", "id", "details", "title", "long-title" })
                {
                    File.WriteAllText(path, state == "null" ? "null" : json.Serialize(new CrashReport.Snapshot { Id = state == "id" ? "invalid" : id, Details = state == "details" ? "" : "owned details", Title = state == "title" ? " " : state == "long-title" ? new string('x', 181) : "Owned title" }));
                    Assert.ThrowsException<InvalidDataException>(() => CrashReport.ReadPending(path), state);
                }
            }
        }
    }
}
