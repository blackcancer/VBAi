using System;
using System.IO;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Web.Script.Serialization;
using System.Security.Cryptography;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ExcelVbeFixtureEmbeddedGitTests
    {
        [TestMethod]
        public void RawHashReadsUnchangedWorkbookWhileAnOwnedWriterAllowsReaders()
        {
            WithFile(path => {
                byte[] original = { 0, 1, 2, 3, 4, 255 };
                File.WriteAllBytes(path, original);
                string expected;
                using (var sha = SHA256.Create())
                    expected = BitConverter.ToString(sha.ComputeHash(original)).Replace("-", "");
                using (var owner = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.Read))
                {
                    Assert.AreEqual(expected, ExcelVbeFixture.EmbeddedRawHash(path));
                    Assert.AreEqual(original.Length, owner.Length);
                    var observed = new byte[original.Length];
                    Assert.AreEqual(original.Length, owner.Read(observed, 0, observed.Length));
                    CollectionAssert.AreEqual(original, observed);
                }
                // The hash reader must have released its handle, including when the owner has closed.
                using (var exclusive = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                    Assert.AreEqual(original.Length, exclusive.Length);
                CollectionAssert.AreEqual(original, File.ReadAllBytes(path));
            });
        }

        [TestMethod]
        public void RawHashPreservesMissingFileFailureAndDoesNotCreateAFile()
        {
            WithFile(path => {
                Assert.ThrowsException<FileNotFoundException>(() => ExcelVbeFixture.EmbeddedRawHash(path));
                Assert.IsFalse(File.Exists(path));
            });
        }

        [DataTestMethod]
        [DataRow("Application.VBE")]
        [DataRow("VBE.MainWindow")]
        public void RetainedVbeReadRecordsExactlyOneNamedGetterOnTheCallingThread(string getter)
        {
            var events = new List<object>(); int reads = 0;
            int thread = System.Threading.Thread.CurrentThread.ManagedThreadId;
            ExcelVbeFixture.ObserveEmbeddedVbeRead(getter, () => reads++, events.Add, 41);
            Assert.AreEqual(1, reads); Assert.AreEqual(2, events.Count);
            foreach (var item in events)
            {
                var row = Row(item);
                Assert.AreEqual(getter, row["Getter"]); Assert.AreEqual(41, row["ProcessId"]);
                Assert.AreEqual(thread, row["OwnerThread"]);
            }
            Assert.AreEqual("VbeAccessIntent", Row(events[0])["Phase"]);
            Assert.AreEqual("VbeAccessReturned", Row(events[1])["Phase"]);
        }

        [TestMethod]
        public void RetainedVbeReadPreservesNativeGetterErrorAndNeverReplaysIt()
        {
            var events = new List<object>(); int reads = 0;
            var primary = new COMException("Synthetic native getter failure", unchecked((int)0x800A03EC));
            var actual = Assert.ThrowsException<COMException>(() => ExcelVbeFixture.ObserveEmbeddedVbeRead(
                "Application.VBE", () => { reads++; throw primary; }, events.Add, 41));
            Assert.AreSame(primary, actual); Assert.AreEqual(1, reads); Assert.AreEqual(2, events.Count);
            Assert.AreEqual("VbeAccessFailed", Row(events[1])["Phase"]);
            Assert.AreEqual("Application.VBE", Row(events[1])["Getter"]);
            Assert.AreEqual("0x800A03EC", Row(events[1])["HResult"]);
        }

        [TestMethod]
        public void RetainedVbeReadKeepsPrimaryGetterFailureWhenItsFailureReceiptThrows()
        {
            int receipts = 0, reads = 0;
            var primary = new COMException("Synthetic getter error", unchecked((int)0x800A03EC));
            var actual = Assert.ThrowsException<COMException>(() => ExcelVbeFixture.ObserveEmbeddedVbeRead(
                "VBE.MainWindow", () => { reads++; throw primary; }, item => {
                    receipts++; if (receipts == 2) throw new IOException("Synthetic receipt failure");
                }, 41));
            Assert.AreSame(primary, actual); Assert.AreEqual(1, reads); Assert.AreEqual(2, receipts);
            StringAssert.Contains(Convert.ToString(actual.Data["VbeAccessEvidenceError"]), "Synthetic receipt failure");
        }

        [TestMethod]
        public void RetainedVbeReadDoesNotAccessNativeObjectWhenIntentReceiptFails()
        {
            int reads = 0;
            Assert.ThrowsException<IOException>(() => ExcelVbeFixture.ObserveEmbeddedVbeRead(
                "Application.VBE", () => reads++, item => { throw new IOException("Intent unavailable"); }, 41));
            Assert.AreEqual(0, reads);
        }

        [TestMethod]
        public void RetainedVbeReadLeavesAcquiredLeaseWithCallerWhenReturnedReceiptFails()
        {
            object lease = null; int reads = 0, releases = 0, receipts = 0;
            var acquired = new object();
            Assert.ThrowsException<IOException>(() => {
                try
                {
                    ExcelVbeFixture.ObserveEmbeddedVbeRead("VBE.MainWindow", () => { reads++; lease = acquired; }, item => {
                        receipts++; if (receipts == 2) throw new IOException("Returned receipt unavailable");
                    }, 41);
                }
                finally { if (lease != null) { Assert.AreSame(acquired, lease); releases++; lease = null; } }
            });
            Assert.AreEqual(1, reads); Assert.AreEqual(1, releases); Assert.AreEqual(2, receipts);
            Assert.IsNull(lease);
        }

        private static IDictionary<string, object> Row(object value)
        { return new JavaScriptSerializer().Deserialize<IDictionary<string, object>>(new JavaScriptSerializer().Serialize(value)); }

        private static void WithFile(Action<string> action)
        {
            string root = Path.Combine(Path.GetTempPath(), "VBAi-embedded-hash-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try { action(Path.Combine(root, "Owned.xlsm")); }
            finally { Directory.Delete(root, true); }
        }
    }
}
