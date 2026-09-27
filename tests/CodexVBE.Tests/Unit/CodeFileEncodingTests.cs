using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class CodeFileEncodingTests
    {
        [TestMethod]
        public void InspectionDistinguishesAsciiUtf8BomAndAmbiguousAnsi()
        {
            var root = Path.Combine(Path.GetTempPath(), "CodexVBE-EncodingTests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try
            {
                var reader = new VbeCodeNavigation(null, new VbeForms(null));
                var ascii = Path.Combine(root, "plain.bas");
                File.WriteAllBytes(ascii, Encoding.ASCII.GetBytes("Option Explicit\r\n"));
                dynamic plain = reader.InspectCodeFile(ascii);
                Assert.AreEqual("utf-8", (string)plain.DefaultEncoding);
                Assert.IsFalse((bool)plain.ExplicitEncodingRequired);
                Assert.IsTrue((bool)plain.StrictUtf8Valid);
                Assert.IsFalse((bool)plain.ContentIncluded);

                var utf8 = Path.Combine(root, "utf8.bas");
                File.WriteAllBytes(utf8, Combine(new byte[] { 0xEF, 0xBB, 0xBF }, Encoding.UTF8.GetBytes("' été\r\n")));
                dynamic withBom = reader.InspectCodeFile(utf8);
                Assert.AreEqual("utf-8", (string)withBom.Bom);
                Assert.AreEqual("utf-8", (string)withBom.DefaultEncoding);
                Assert.IsTrue((bool)withBom.ContainsNonAscii);
                Assert.IsFalse((bool)withBom.ExplicitEncodingRequired);

                var ansi = Path.Combine(root, "ansi.bas");
                File.WriteAllBytes(ansi, new byte[] { 0x27, 0x20, 0xE9, 0x0D, 0x0A });
                dynamic ambiguous = reader.InspectCodeFile(ansi);
                Assert.IsNull((object)ambiguous.Bom);
                Assert.IsNull((object)ambiguous.DefaultEncoding);
                Assert.IsFalse((bool)ambiguous.StrictUtf8Valid);
                Assert.IsTrue((bool)ambiguous.ExplicitEncodingRequired);
            }
            finally { Directory.Delete(root, true); }
        }

        [TestMethod]
        public void InspectionReportsUtf16AndRejectsUnsafeFileInputs()
        {
            var root = Path.Combine(Path.GetTempPath(), "CodexVBE-EncodingTests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try
            {
                var reader = new VbeCodeNavigation(null, new VbeForms(null));
                Assert.ThrowsException<ArgumentException>(() => reader.InspectCodeFile("relative.bas"));
                Assert.ThrowsException<FileNotFoundException>(() => reader.InspectCodeFile(Path.Combine(root, "missing.bas")));
                var empty = Path.Combine(root, "empty.bas");
                File.WriteAllBytes(empty, new byte[0]);
                Assert.ThrowsException<ArgumentException>(() => reader.InspectCodeFile(empty));
                var oversized = Path.Combine(root, "oversized.bas");
                File.WriteAllBytes(oversized, new byte[256 * 1024 + 1]);
                Assert.ThrowsException<ArgumentException>(() => reader.InspectCodeFile(oversized));

                var utf16 = Path.Combine(root, "utf16.bas");
                File.WriteAllBytes(utf16, Combine(new byte[] { 0xFF, 0xFE }, Encoding.Unicode.GetBytes("' été")));
                dynamic unicode = reader.InspectCodeFile(utf16);
                Assert.AreEqual("utf-16le", (string)unicode.Bom);
                Assert.IsTrue((bool)unicode.ContainsNulByte);
                Assert.IsNull((object)unicode.StrictUtf8Valid);
                Assert.IsFalse((bool)unicode.ExplicitEncodingRequired);
            }
            finally { Directory.Delete(root, true); }
        }

        [TestMethod]
        public void InsertionRequiresExplicitEncodingAndRollsBackCorruptedUnicode()
        {
            var root = Path.Combine(Path.GetTempPath(), "CodexVBE-EncodingTests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try
            {
                var source = Path.Combine(root, "accent.bas");
                File.WriteAllBytes(source, Encoding.GetEncoding(1252).GetBytes("' été"));
                var module = new VbeSessionContractTests.FakeModule();
                var project = new VbeSessionContractTests.FakeProject { Name = "Projet", Mode = 2 };
                project.VBComponents.Add(new VbeSessionContractTests.FakeComponent {
                    Name = "Module1", Type = 1, CodeModule = module
                });
                var host = new VbeSessionContractTests.FakeVbe();
                host.VBProjects.Add(project);
                var navigation = new VbeCodeNavigation(host, new VbeForms(host));
                var request = new Request { Project = "Projet", Module = "Module1", Path = source,
                    StartLine = 1, ExpectedSha256 = Hash(string.Empty) };
                Assert.ThrowsException<ArgumentException>(() => navigation.InsertCodeFile(request));
                request.SourceEncoding = "utf-8";
                Assert.ThrowsException<ArgumentException>(() => navigation.InsertCodeFile(request));
                request.SourceEncoding = "windows-1252";
                module.CorruptNonAsciiOnInsert = true;
                Assert.ThrowsException<InvalidOperationException>(() => navigation.InsertCodeFile(request));
                Assert.AreEqual(0, module.CountOfLines, "Unicode corruption must roll back all inserted lines.");
                module.CorruptNonAsciiOnInsert = false;
                dynamic inserted = navigation.InsertCodeFile(request);
                Assert.AreEqual("windows-1252", (string)inserted.SourceEncoding);
                Assert.AreEqual(1, (int)inserted.InsertedLineCount);
                StringAssert.Contains(module.Code, "été");
                Assert.IsFalse((bool)inserted.CompilationVerified);
                Assert.ThrowsException<InvalidOperationException>(() => navigation.InsertCodeFile(request));
            }
            finally { Directory.Delete(root, true); }
        }

        private static string Hash(string value)
        {
            using (var sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(value)))
                    .Replace("-", "").ToLowerInvariant();
        }

        private static byte[] Combine(byte[] prefix, byte[] body)
        {
            var result = new byte[prefix.Length + body.Length];
            Buffer.BlockCopy(prefix, 0, result, 0, prefix.Length);
            Buffer.BlockCopy(body, 0, result, prefix.Length, body.Length);
            return result;
        }
    }
}
