using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class RetainedRootFontWorkbookTests
    {
        [TestMethod]
        public void RetainedMarkerIsIndependentOfFreshCorrelationNonce()
        {
            const string nonce = "fresh-correlation";
            Assert.AreEqual(RetainedRootFontWorkbook.SourceMarker,
                RetainedRootFontWorkbook.BaselineMarker(RootFontObservationManifest.RetainedSyntheticTahoma825, nonce));
            Assert.AreNotEqual(nonce, RetainedRootFontWorkbook.SourceMarker);
            Assert.AreEqual(nonce, RetainedRootFontWorkbook.BaselineMarker(null, nonce));
            Assert.AreEqual(nonce, RetainedRootFontWorkbook.BaselineMarker(
                RootFontObservationManifest.SyntheticExplicitArial9, nonce));
        }

        [TestMethod]
        public void FirstCaptureMustVerifyBeforeEvidenceOrAnyFollowingReadback()
        {
            var order = new List<string>();
            string snapshot = RetainedRootFontWorkbook.AttestFirstCapture(
                () => { order.Add("capture"); return "snapshot"; },
                value => { Assert.AreEqual("snapshot", value); order.Add("verify"); return new byte[] { 1 }; },
                (value, bytes) => { Assert.AreEqual("snapshot", value); Assert.AreEqual((byte)1, bytes[0]); order.Add("record"); });
            Assert.AreEqual("snapshot", snapshot);
            CollectionAssert.AreEqual(new[] { "capture", "verify", "record" }, order);
            order.Clear();
            var original = new InvalidOperationException("baseline drift");
            Assert.AreSame(original, Assert.ThrowsException<InvalidOperationException>(() =>
                RetainedRootFontWorkbook.AttestFirstCapture(
                    () => { order.Add("capture"); return "snapshot"; },
                    value => { order.Add("verify"); throw original; },
                    (value, bytes) => order.Add("record"))));
            CollectionAssert.AreEqual(new[] { "capture", "verify" }, order);
        }

        [TestMethod]
        public void CopyCreatesOneExactDisposableFileAndRefusesCollisionOrDrift()
        {
            WithRoot(root => {
                string source = Path.Combine(root, "source.xlsm");
                string destination = Path.Combine(root, "EmbeddedGit.xlsm");
                byte[] bytes = { 1, 2, 3, 4, 5 };
                File.WriteAllBytes(source, bytes);
                string sha = Sha(bytes);
                RetainedRootFontWorkbook.CopyCreateNew(source, destination, root, sha);
                CollectionAssert.AreEqual(bytes, File.ReadAllBytes(destination));
                Assert.ThrowsException<InvalidOperationException>(() =>
                    RetainedRootFontWorkbook.CopyCreateNew(source, destination, root, sha));
                CollectionAssert.AreEqual(bytes, File.ReadAllBytes(destination));
                File.Delete(destination);
                File.WriteAllBytes(source, new byte[] { 9 });
                Assert.ThrowsException<InvalidOperationException>(() =>
                    RetainedRootFontWorkbook.CopyCreateNew(source, destination, root, sha));
                Assert.IsFalse(File.Exists(destination));
            });
        }

        [TestMethod]
        public void CopyRefusesReparseSourceAndDestinationOutsideFixture()
        {
            WithRoot(root => {
                string source = Path.Combine(root, "source.xlsm");
                string destination = Path.Combine(root, "EmbeddedGit.xlsm");
                File.WriteAllBytes(source, new byte[] { 1 });
                string sha = Sha(new byte[] { 1 });
                Assert.ThrowsException<InvalidOperationException>(() =>
                    RetainedRootFontWorkbook.CopyCreateNew(source, destination, root, sha,
                        path => path == source ? FileAttributes.ReparsePoint : File.GetAttributes(path)));
                Assert.IsFalse(File.Exists(destination));
                Assert.ThrowsException<InvalidOperationException>(() =>
                    RetainedRootFontWorkbook.CopyCreateNew(source, Path.Combine(Path.GetDirectoryName(root),
                        "EmbeddedGit.xlsm"), root, sha));
                Assert.IsFalse(File.Exists(destination));
            });
        }

        [TestMethod]
        public void RetainedDescriptorRequiresExactStoredTahomaAndSingleRoot()
        {
            byte[] exact = RetainedRootFontWorkbook.Tahoma825Descriptor();
            CollectionAssert.AreEqual(exact, RootFontObservationManifest.RequireRoot(
                new[] { new FormStreamPadding.FormFontBinding("", exact, 7) },
                "AfterInitialCapture", RootFontObservationManifest.RetainedSyntheticTahoma825));
            foreach (byte[] invalid in new[] { exact.Take(exact.Length - 1).ToArray(),
                RootFontObservationManifest.SyntheticArial9Descriptor() })
                Assert.ThrowsException<InvalidOperationException>(() => RootFontObservationManifest.RequireRoot(
                    new[] { new FormStreamPadding.FormFontBinding("", invalid, 7) },
                    "AfterInitialCapture", RootFontObservationManifest.RetainedSyntheticTahoma825));
            Assert.ThrowsException<InvalidOperationException>(() => RootFontObservationManifest.RequireRoot(
                Array.Empty<FormStreamPadding.FormFontBinding>(), "AfterInitialCapture",
                RootFontObservationManifest.RetainedSyntheticTahoma825));
        }

        private static void WithRoot(Action<string> action)
        {
            string root = Path.Combine(Path.GetTempPath(), "VBAi-retained-font-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try { action(root); }
            finally { Directory.Delete(root, true); }
        }

        private static string Sha(byte[] bytes)
        {
            using (var hash = SHA256.Create())
                return BitConverter.ToString(hash.ComputeHash(bytes)).Replace("-", "");
        }
    }
}
