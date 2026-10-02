using System;
using System.IO;
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

        private static void WithFile(Action<string> action)
        {
            string root = Path.Combine(Path.GetTempPath(), "VBAi-embedded-hash-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try { action(Path.Combine(root, "Owned.xlsm")); }
            finally { Directory.Delete(root, true); }
        }
    }
}
