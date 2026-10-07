using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Web.Script.Serialization;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ExcelVbeFixtureEmbeddedGitTests
    {
        [TestMethod]
        public void RawHashReadsUnchangedWorkbookWhileAnOwnedWriterAllowsReaders()
        {
            WithFile(path =>
            {
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
            WithFile(path =>
            {
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
                "VBE.MainWindow", () => { reads++; throw primary; }, item =>
                {
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
            Assert.ThrowsException<IOException>(() =>
            {
                try
                {
                    ExcelVbeFixture.ObserveEmbeddedVbeRead("VBE.MainWindow", () => { reads++; lease = acquired; }, item =>
                    {
                        receipts++; if (receipts == 2) throw new IOException("Returned receipt unavailable");
                    }, 41);
                }
                finally { if (lease != null) { Assert.AreSame(acquired, lease); releases++; lease = null; } }
            });
            Assert.AreEqual(1, reads); Assert.AreEqual(1, releases); Assert.AreEqual(2, receipts);
            Assert.IsNull(lease);
        }

        [DataTestMethod]
        [DataRow(null, false)]
        [DataRow(null, true)]
        [DataRow("", false)]
        [DataRow("", true)]
        [DataRow("BeforeAfterCopy", true)]
        public void RetainedVbeLifetimeOptInIsExactAndScoped(string setting, bool retained)
        {
            Assert.AreEqual(setting == "BeforeAfterCopy",
                ExcelVbeFixture.RetainedVbeLifetimeEnabled(setting, retained));
        }

        [DataTestMethod]
        [DataRow("BeforeAfterCopy", false)]
        [DataRow("beforeaftercopy", true)]
        [DataRow("AfterInitialCapture", true)]
        [DataRow("0", false)]
        [DataRow(" ", true)]
        public void RetainedVbeLifetimeRefusesEveryOtherConfiguredValue(string setting, bool retained)
        {
            Assert.ThrowsException<InvalidOperationException>(() =>
                ExcelVbeFixture.RetainedVbeLifetimeEnabled(setting, retained));
        }

        [TestMethod]
        public void BeforeCopyVbeReadOrdersExactReceiptsReleaseAndOneReplacement()
        {
            var order = new List<string>();
            var acquired = new object(); int reads = 0, releases = 0, replacements = 0;
            int owner = System.Threading.Thread.CurrentThread.ManagedThreadId;
            ExcelVbeFixture.ObserveRetainedVbeBeforeCopy(
                () => { order.Add("read"); reads++; return acquired; },
                value => { Assert.AreSame(acquired, value); order.Add("release"); releases++; },
                () => { order.Add("replace"); replacements++; },
                item =>
                {
                    var receipt = Row(item);
                    Assert.AreEqual("BeforeCopy", receipt["Stage"]);
                    Assert.AreEqual("Application.VBE", receipt["Getter"]);
                    Assert.AreEqual(41, receipt["ProcessId"]);
                    Assert.AreEqual(owner, receipt["OwnerThread"]);
                    order.Add(Convert.ToString(receipt["Phase"]));
                }, 41);
            CollectionAssert.AreEqual(new[] { "VbeAccessIntent", "read", "VbeAccessReturned", "release", "replace" }, order);
            Assert.AreEqual(1, reads); Assert.AreEqual(1, releases); Assert.AreEqual(1, replacements);
        }

        [TestMethod]
        public void BeforeCopyVbeReadValidatesDelegatesBeforeReadingOrReplacing()
        {
            int reads = 0, releases = 0, replacements = 0;
            Func<object> read = () => { reads++; return new object(); };
            Action<object> release = _ => releases++;
            Action replace = () => replacements++;
            Action<object> evidence = _ => { };
            Assert.ThrowsException<ArgumentNullException>(() => ExcelVbeFixture.ObserveRetainedVbeBeforeCopy(null, release, replace, evidence, 41));
            Assert.ThrowsException<ArgumentNullException>(() => ExcelVbeFixture.ObserveRetainedVbeBeforeCopy(read, null, replace, evidence, 41));
            Assert.ThrowsException<ArgumentNullException>(() => ExcelVbeFixture.ObserveRetainedVbeBeforeCopy(read, release, null, evidence, 41));
            Assert.ThrowsException<ArgumentNullException>(() => ExcelVbeFixture.ObserveRetainedVbeBeforeCopy(read, release, replace, null, 41));
            Assert.AreEqual(0, reads); Assert.AreEqual(0, releases); Assert.AreEqual(0, replacements);
        }

        [DataTestMethod]
        [DataRow("intent")]
        [DataRow("getter")]
        [DataRow("returned")]
        [DataRow("null")]
        [DataRow("release")]
        public void BeforeCopyVbeReadNeverReplacesWhenAnyPrerequisiteFails(string fault)
        {
            int reads = 0, releases = 0, replacements = 0, receipts = 0;
            var acquired = new object();
            var getterError = new COMException("Synthetic VBE getter failed", unchecked((int)0x800A03EC));
            var receiptError = new IOException("Synthetic receipt failed");
            var releaseError = new InvalidOperationException("Synthetic release failed");
            Action invoke = () => ExcelVbeFixture.ObserveRetainedVbeBeforeCopy(
                () => { reads++; if (fault == "getter") throw getterError; return fault == "null" ? null : acquired; },
                value =>
                {
                    if (fault == "getter" || fault == "intent" || fault == "null") Assert.IsNull(value);
                    else Assert.AreSame(acquired, value);
                    releases++;
                    if (fault == "release") throw releaseError;
                },
                () => replacements++,
                _ => { receipts++; if (fault == "intent" || fault == "returned" && receipts == 2) throw receiptError; }, 41);
            if (fault == "intent" || fault == "returned") Assert.AreSame(receiptError, Assert.ThrowsException<IOException>(invoke));
            else if (fault == "getter") Assert.AreSame(getterError, Assert.ThrowsException<COMException>(invoke));
            else if (fault == "null") Assert.ThrowsException<InvalidOperationException>(invoke);
            else
            {
                var failure = Assert.ThrowsException<AggregateException>(invoke);
                CollectionAssert.AreEqual(new[] { releaseError }, new List<Exception>(failure.InnerExceptions));
            }
            Assert.AreEqual(fault == "intent" ? 0 : 1, reads);
            Assert.AreEqual(1, releases, "Even an absent lease is visited once by the scoped cleanup helper.");
            Assert.AreEqual(0, replacements);
            Assert.AreEqual(fault == "intent" ? 1 : 2, receipts);
        }

        [TestMethod]
        public void BeforeCopyVbeReadPreservesOriginalGetterFailureWhenFailureReceiptAlsoFails()
        {
            int reads = 0, releases = 0, replacements = 0, receipts = 0;
            var original = new COMException("Synthetic native VBE failure", unchecked((int)0x800A03EC));
            var observed = Assert.ThrowsException<COMException>(() => ExcelVbeFixture.ObserveRetainedVbeBeforeCopy(
                () => { reads++; throw original; }, _ => releases++, () => replacements++, item =>
                {
                    receipts++;
                    if (receipts == 2)
                    {
                        Assert.AreEqual("VbeAccessFailed", Row(item)["Phase"]);
                        Assert.AreEqual("BeforeCopy", Row(item)["Stage"]);
                        Assert.AreEqual("0x800A03EC", Row(item)["HResult"]);
                        throw new IOException("Synthetic failure receipt unavailable");
                    }
                }, 41));
            Assert.AreSame(original, observed);
            Assert.AreEqual(1, reads); Assert.AreEqual(2, receipts);
            Assert.AreEqual(1, releases); Assert.AreEqual(0, replacements);
            StringAssert.Contains(Convert.ToString(original.Data["VbeAccessEvidenceError"]), "Synthetic failure receipt unavailable");
        }

        [TestMethod]
        public void BeforeCopyVbeReadRetainsPrimaryThenCleanupErrorWithoutReplacement()
        {
            int reads = 0, releases = 0, replacements = 0;
            var native = new COMException("Synthetic native getter failure", unchecked((int)0x800A03EC));
            var cleanup = new IOException("Synthetic release failure");
            var failure = Assert.ThrowsException<AggregateException>(() => ExcelVbeFixture.ObserveRetainedVbeBeforeCopy(
                () => { reads++; throw native; }, value =>
                {
                    Assert.IsNull(value); releases++; throw cleanup;
                }, () => replacements++, _ => { }, 41));
            Assert.AreEqual(2, failure.InnerExceptions.Count);
            Assert.AreSame(native, failure.InnerExceptions[0]);
            Assert.AreSame(cleanup, failure.InnerExceptions[1]);
            Assert.AreEqual(1, reads); Assert.AreEqual(1, releases); Assert.AreEqual(0, replacements);
        }

        [TestMethod]
        public void BeforeCopyVbeReadDoesNotRetryAThrowingReplacement()
        {
            int reads = 0, releases = 0, replacements = 0;
            var replacementError = new InvalidOperationException("Synthetic replacement outcome unknown");
            var observed = Assert.ThrowsException<InvalidOperationException>(() => ExcelVbeFixture.ObserveRetainedVbeBeforeCopy(
                () => { reads++; return new object(); }, _ => releases++,
                () => { replacements++; throw replacementError; }, _ => { }, 41));
            Assert.AreSame(replacementError, observed);
            Assert.AreEqual(1, reads); Assert.AreEqual(1, releases); Assert.AreEqual(1, replacements);
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
