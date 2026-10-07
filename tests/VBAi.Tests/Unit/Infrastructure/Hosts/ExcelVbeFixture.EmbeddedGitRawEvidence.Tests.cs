using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Web.Script.Serialization;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ExcelVbeFixtureEmbeddedGitRawEvidenceTests
    {
        [TestMethod]
        public void RetainsExactTargetAndActualBytesForEveryFileBeforeTheReceipt()
        {
            WithRoot(root =>
            {
                var target = Files(0x25); var actual = Files(0x27); int receipts = 0;
                ExcelVbeFixture.RetainEmbeddedImportRawEvidence(root, target, actual, receipt =>
                {
                    receipts++;
                    Assert.IsTrue(File.Exists(Path.Combine(root, "post-import-raw", "raw-evidence.json")));
                    foreach (var side in new[] { "target", "actual" })
                        foreach (var name in target.Keys)
                            Assert.IsTrue(File.Exists(Path.Combine(root, "post-import-raw", side, name)));
                });
                Assert.AreEqual(1, receipts);
                foreach (var side in new[] { "target", "actual" })
                {
                    var source = side == "target" ? target : actual;
                    foreach (var file in source)
                        CollectionAssert.AreEqual(file.Value, File.ReadAllBytes(Path.Combine(root, "post-import-raw", side, file.Key)));
                }
                var manifest = (Dictionary<string, object>)new JavaScriptSerializer().DeserializeObject(
                    File.ReadAllText(Path.Combine(root, "post-import-raw", "raw-evidence.json")));
                var rows = (object[])manifest["Actual"];
                Assert.AreEqual("PreparedBaseline", manifest["TargetRole"]);
                Assert.AreEqual("ActualPostImportCapture", manifest["ActualRole"]);
                Assert.AreEqual(actual.Count, rows.Length);
                foreach (Dictionary<string, object> row in rows)
                {
                    string name = (string)row["Name"];
                    using (var sha = SHA256.Create())
                        Assert.AreEqual(BitConverter.ToString(sha.ComputeHash(actual[name])).Replace("-", ""), row["Sha256"]);
                }
            });
        }

        [DataTestMethod]
        [DataRow("null-root")]
        [DataRow("relative-root")]
        [DataRow("null-target")]
        [DataRow("null-actual")]
        [DataRow("null-receipt")]
        [DataRow("unsafe-name")]
        [DataRow("null-bytes")]
        public void RefusesMissingOrUnsafeEvidenceInputs(string fault)
        {
            WithRoot(root =>
            {
                var target = Files(0x25); var actual = Files(0x27);
                if (fault == "unsafe-name") actual.Add("../escape.frx", new byte[] { 1 });
                if (fault == "null-bytes") actual["EmbeddedForm.frx"] = null;
                int receipts = 0;
                Action call = () => ExcelVbeFixture.RetainEmbeddedImportRawEvidence(
                    fault == "null-root" ? null : fault == "relative-root" ? "relative-q027-evidence" : root,
                    fault == "null-target" ? null : target,
                    fault == "null-actual" ? null : actual,
                    fault == "null-receipt" ? null : new Action<object>(_ => receipts++));
                try { call(); Assert.Fail("Missing or unsafe raw evidence must be refused."); }
                catch (ArgumentException) { }
                catch (InvalidOperationException) { }
                Assert.AreEqual(0, receipts);
            });
        }

        [TestMethod]
        public void ReceiptFailureKeepsDurableRawBytesAndDoesNotRepeatCapture()
        {
            WithRoot(root =>
            {
                int receipts = 0;
                var failure = new IOException("Synthetic receipt failure");
                Assert.AreSame(failure, Assert.ThrowsException<IOException>(() =>
                    ExcelVbeFixture.RetainEmbeddedImportRawEvidence(root, Files(0x25), Files(0x27), _ =>
                    {
                        receipts++; throw failure;
                    })));
                Assert.AreEqual(1, receipts);
                CollectionAssert.AreEqual(new byte[] { 0x25 },
                    File.ReadAllBytes(Path.Combine(root, "post-import-raw", "target", "EmbeddedForm.frx")));
                CollectionAssert.AreEqual(new byte[] { 0x27 },
                    File.ReadAllBytes(Path.Combine(root, "post-import-raw", "actual", "EmbeddedForm.frx")));
                Assert.IsTrue(File.Exists(Path.Combine(root, "post-import-raw", "raw-evidence.json")));
            });
        }

        [TestMethod]
        public void RefusesAnExistingEvidenceRootBeforeWritingOrRepeatingTheReceipt()
        {
            WithRoot(root =>
            {
                int receipts = 0; var target = Files(0x25);
                ExcelVbeFixture.RetainEmbeddedImportRawEvidence(root, target, Files(0x27), _ => receipts++);
                Assert.ThrowsException<IOException>(() => ExcelVbeFixture.RetainEmbeddedImportRawEvidence(
                    root, target, Files(0x28), _ => receipts++));
                Assert.AreEqual(1, receipts);
                CollectionAssert.AreEqual(new byte[] { 0x25 },
                    File.ReadAllBytes(Path.Combine(root, "post-import-raw", "target", "EmbeddedForm.frx")));
            });
        }

        [DataTestMethod]
        [DataRow("none")]
        [DataRow("snapshot")]
        [DataRow("native")]
        [DataRow("both")]
        public void ReadsNativeOnceAfterSnapshotFailureAndPreservesBothOriginalExceptions(string fault)
        {
            int snapshots = 0, native = 0;
            var primary = new InvalidOperationException("Exact FRX differs");
            var readback = new IOException("Native readback failed");
            Action call = () => ExcelVbeFixture.VerifyEmbeddedImportReadbacks(
                () => { snapshots++; if (fault == "snapshot" || fault == "both") throw primary; },
                () => { native++; if (fault == "native" || fault == "both") throw readback; });
            if (fault == "snapshot") Assert.AreSame(primary, Assert.ThrowsException<InvalidOperationException>(call));
            else if (fault == "native") Assert.AreSame(readback, Assert.ThrowsException<IOException>(call));
            else if (fault == "both")
            {
                var combined = Assert.ThrowsException<AggregateException>(call);
                CollectionAssert.AreEqual(new Exception[] { primary, readback }, combined.InnerExceptions.ToArray());
            }
            else call();
            Assert.AreEqual(1, snapshots); Assert.AreEqual(1, native);
        }

        [TestMethod]
        public void ValidatesBothReadbackCallbacksBeforeAnyAction()
        {
            int calls = 0;
            Assert.ThrowsException<ArgumentNullException>(() => ExcelVbeFixture.VerifyEmbeddedImportReadbacks(null, () => calls++));
            Assert.ThrowsException<ArgumentNullException>(() => ExcelVbeFixture.VerifyEmbeddedImportReadbacks(() => calls++, null));
            Assert.AreEqual(0, calls);
        }

        private static Dictionary<string, byte[]> Files(byte resource)
        {
            return new Dictionary<string, byte[]>
            {
                ["EmbeddedForm.frm"] = new byte[] { 0x41, 0x0A },
                ["EmbeddedForm.frx"] = new byte[] { resource },
                ["EmbeddedClass.cls"] = new byte[] { 0x43 },
                ["manifest.json"] = new byte[] { 0x7B, 0x7D }
            };
        }

        private static void WithRoot(Action<string> action)
        {
            string root = Path.Combine(Path.GetTempPath(), "Q027-RawEvidence-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try { action(root); }
            finally { Directory.Delete(root, true); }
        }
    }
}
