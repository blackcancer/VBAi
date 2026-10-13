using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Linq;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks source-literal provenance and strict rejection of malformed synthetic tamper inputs.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class SyntheticSignatureTamperTests
    {
        [DataTestMethod]
        [DataRow("0103B000414243", "ABC", "4,5,6")]
        [DataRow("0105B0084142430020", "ABCABC", "4,5,6,4,5,6")]
        [DataRow("0103B002410000", "AAAA", "4,4,4,4")]
        public void LiteralAndCopyTokensRetainExactOrigin(string hex, string expected, string expectedOrigins)
        {
            int[] origins;
            byte[] decoded = SyntheticSignatureTamper.Decompress(Hex(hex), out origins);
            Assert.AreEqual(expected, System.Text.Encoding.ASCII.GetString(decoded));
            Assert.AreEqual(expectedOrigins, string.Join(",", origins));
        }

        [TestMethod]
        public void RawChunkHasOneDistinctOriginForEachByte()
        {
            byte[] raw = new byte[4099]; raw[0] = 1; raw[1] = 0xff; raw[2] = 0x3f;
            for (int i = 3; i < raw.Length; i++) raw[i] = (byte)(i % 251);
            int[] origins;
            byte[] decoded = SyntheticSignatureTamper.Decompress(raw, out origins);
            CollectionAssert.AreEqual(raw.Skip(3).ToArray(), decoded);
            CollectionAssert.AreEqual(Enumerable.Range(3, 4096).ToArray(), origins);
        }

        [DataTestMethod]
        [DataRow("")]
        [DataRow("02")]
        [DataRow("0100")]
        [DataRow("01000000")]
        [DataRow("0101B000")]
        [DataRow("0101B00100")]
        [DataRow("0102B0010000")]
        [DataRow("0103B002410020")]
        [DataRow("01003000")]
        public void InvalidContainersAreRefused(string hex)
        {
            Assert.ThrowsException<InvalidDataException>(() => SyntheticSignatureTamper.Decompress(Hex(hex), out _));
        }

        [TestMethod]
        public void NullAndOversizedContainersAreRefused()
        {
            Assert.ThrowsException<InvalidDataException>(() => SyntheticSignatureTamper.Decompress(null, out _));
            byte[] bytes = new byte[1048577]; bytes[0] = 1;
            Assert.ThrowsException<InvalidDataException>(() => SyntheticSignatureTamper.Decompress(bytes, out _));
        }

        [TestMethod]
        public void CopyRefusesAWorkbookOutsideTheOwnedResultRootBeforeOpeningIt()
        {
            string root = Path.Combine(Path.GetTempPath(), "VBAi-Signature-" + Guid.NewGuid().ToString("N"));
            string outside = Path.Combine(Path.GetTempPath(), "SyntheticSignature.xlsm");
            Assert.ThrowsException<InvalidOperationException>(() => SyntheticSignatureTamper.CreateCopy(outside, root));
            Assert.IsFalse(Directory.Exists(root), "Admission failure must not create or alter any directory.");
        }

        [TestMethod]
        public void CompoundMutationChangesExactlyTheKnownSourceLetterAndPreservesTheInput()
        {
            byte[] original = Compound("0123456789abcdef0123456789abcdef");
            byte[] snapshot = (byte[])original.Clone();
            byte[] mutated = SyntheticSignatureTamper.PatchProject(original);
            CollectionAssert.AreEqual(snapshot, original);
            CollectionAssert.AreEqual(Compound("a123456789abcdef0123456789abcdef"), mutated);
            Assert.AreEqual(1, original.Where((value, index) => value != mutated[index]).Count());
        }

        [TestMethod]
        public void ASourceWithoutTheExactSyntheticMarkerCannotBePatched()
        {
            byte[] original = Compound("not-a-guid");
            byte[] snapshot = (byte[])original.Clone();
            Assert.ThrowsException<InvalidDataException>(() => SyntheticSignatureTamper.PatchProject(original));
            CollectionAssert.AreEqual(snapshot, original);
        }

        [TestMethod]
        public void ZipCopyPreservesSignaturePayloadsAndRefusesAnExistingDestination()
        {
            string root = Path.Combine(Path.GetTempPath(), "VBAi-Signature-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try
            {
                string original = Path.Combine(root, "SyntheticSignature.xlsm");
                byte[] signature = new byte[] { 17, 19, 23, 29 }; // Synthetic payload proves preservation, not cryptographic validity.
                using (var archive = System.IO.Compression.ZipFile.Open(original, System.IO.Compression.ZipArchiveMode.Create))
                {
                    using (var stream = archive.CreateEntry("xl/vbaProject.bin").Open())
                    {
                        byte[] compound = Compound("0123456789abcdef0123456789abcdef");
                        stream.Write(compound, 0, compound.Length);
                    }
                    using (var stream = archive.CreateEntry("xl/vbaProjectSignature.bin").Open()) stream.Write(signature, 0, signature.Length);
                }
                byte[] snapshot = File.ReadAllBytes(original);
                string copy = SyntheticSignatureTamper.CreateCopy(original, root);
                CollectionAssert.AreEqual(snapshot, File.ReadAllBytes(original));
                using (var archive = System.IO.Compression.ZipFile.OpenRead(copy))
                using (var stream = archive.GetEntry("xl/vbaProjectSignature.bin").Open())
                using (var data = new MemoryStream())
                {
                    stream.CopyTo(data); CollectionAssert.AreEqual(signature, data.ToArray());
                }
                byte[] existing = File.ReadAllBytes(copy);
                Assert.ThrowsException<IOException>(() => SyntheticSignatureTamper.CreateCopy(original, root));
                CollectionAssert.AreEqual(existing, File.ReadAllBytes(copy));
            }
            finally { Directory.Delete(root, true); }
        }

        private static byte[] Compound(string nonce)
        {
            byte[] name = System.Text.Encoding.ASCII.GetBytes("SignatureFixture");
            byte[] directory;
            using (var data = new MemoryStream())
            using (var writer = new BinaryWriter(data))
            {
                foreach (ushort tag in new ushort[] { 0x19, 0x1a }) { writer.Write(tag); writer.Write(name.Length); writer.Write(name); }
                writer.Write((ushort)0x31); writer.Write(4); writer.Write(0);
                writer.Write((ushort)0x21); writer.Write(0);
                writer.Write((ushort)0x2b); writer.Write(0);
                directory = Literals(data.ToArray());
            }
            byte[] module = Literals(System.Text.Encoding.ASCII.GetBytes(
                "Attribute VB_Name = \"SignatureFixture\"\r\nOption Explicit\r\nPrivate Const Marker As String = \"SIGNATURE_" + nonce + "\"\r\n"));
            int directorySectors = (directory.Length + 63) / 64, moduleSectors = (module.Length + 63) / 64;
            Assert.IsTrue(directorySectors + moduleSectors <= 8);
            byte[] bytes = new byte[3072];
            Buffer.BlockCopy(Hex("D0CF11E0A1B11AE1"), 0, bytes, 0, 8);
            Put16(bytes, 24, 0x3e); Put16(bytes, 26, 3); Put16(bytes, 28, 0xfffe);
            Put16(bytes, 30, 9); Put16(bytes, 32, 6);
            Put32(bytes, 44, 1); Put32(bytes, 48, 1); Put32(bytes, 56, 4096);
            Put32(bytes, 60, 3); Put32(bytes, 64, 1); Put32(bytes, 68, 0xfffffffe);
            for (int i = 0; i < 109; i++) Put32(bytes, 76 + i * 4, 0xffffffff);
            Put32(bytes, 76, 0);
            for (int i = 0; i < 128; i++) { Put32(bytes, 512 + i * 4, 0xffffffff); Put32(bytes, 2048 + i * 4, 0xffffffff); }
            Put32(bytes, 512, 0xfffffffd); Put32(bytes, 516, 2);
            foreach (int sector in new[] { 2, 3, 4 }) Put32(bytes, 512 + sector * 4, 0xfffffffe);
            Entry(bytes, 1024, "Root Entry", 5, 1, 0xffffffff, 4, 512);
            Entry(bytes, 1152, "VBA", 1, 2, 0xffffffff, 0xfffffffe, 0);
            Entry(bytes, 1280, "dir", 2, 0xffffffff, 3, 0, directory.Length);
            Entry(bytes, 1408, "SignatureFixture", 2, 0xffffffff, 0xffffffff, (uint)directorySectors, module.Length);
            for (int i = 0; i < directorySectors; i++) Put32(bytes, 2048 + i * 4, i + 1 == directorySectors ? 0xfffffffe : (uint)i + 1);
            for (int i = 0; i < moduleSectors; i++) Put32(bytes, 2048 + (directorySectors + i) * 4, i + 1 == moduleSectors ? 0xfffffffe : (uint)(directorySectors + i + 1));
            Buffer.BlockCopy(directory, 0, bytes, 2560, directory.Length);
            Buffer.BlockCopy(module, 0, bytes, 2560 + directorySectors * 64, module.Length);
            return bytes;
        }

        private static byte[] Literals(byte[] source)
        {
            using (var result = new MemoryStream())
            using (var writer = new BinaryWriter(result))
            {
                writer.Write((byte)1);
                int body = source.Length + (source.Length + 7) / 8;
                writer.Write((ushort)(0xb000 | (body - 1)));
                for (int i = 0; i < source.Length; i += 8) { writer.Write((byte)0); writer.Write(source, i, Math.Min(8, source.Length - i)); }
                return result.ToArray();
            }
        }

        private static void Entry(byte[] data, int offset, string name, byte kind, uint child, uint right, uint start, int size)
        {
            byte[] encoded = System.Text.Encoding.Unicode.GetBytes(name + "\0");
            Buffer.BlockCopy(encoded, 0, data, offset, encoded.Length); Put16(data, offset + 64, (ushort)encoded.Length);
            data[offset + 66] = kind; data[offset + 67] = 1;
            Put32(data, offset + 68, 0xffffffff); Put32(data, offset + 72, right); Put32(data, offset + 76, child);
            Put32(data, offset + 116, start); Put32(data, offset + 120, (uint)size);
        }

        private static void Put16(byte[] data, int offset, ushort value) { Buffer.BlockCopy(BitConverter.GetBytes(value), 0, data, offset, 2); }
        private static void Put32(byte[] data, int offset, uint value) { Buffer.BlockCopy(BitConverter.GetBytes(value), 0, data, offset, 4); }

        private static byte[] Hex(string value)
        {
            return Enumerable.Range(0, value.Length / 2).Select(i => Convert.ToByte(value.Substring(i * 2, 2), 16)).ToArray();
        }
    }
}
