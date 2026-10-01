namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class FormResourcePreflightTests
    {
        [TestMethod]
        public void BoundedCompoundStorageAndMiniStreamAreValidatedWithoutChangingRawBytes()
        {
            byte[] resource = Resource();
            byte[] original = (byte[])resource.Clone();
            FormResourcePreflight.ValidateOleObjectBlob(resource, 0);
            CollectionAssert.AreEqual(original, resource);
            byte[] prefixed = new byte[resource.Length + 13];
            Buffer.BlockCopy(resource, 0, prefixed, 13, resource.Length);
            FormResourcePreflight.ValidateOleObjectBlob(prefixed, 13);
            Assert.ThrowsException<InvalidOperationException>(() => FormResourcePreflight.ValidateOleObjectBlob(resource, -1));
            Assert.ThrowsException<InvalidOperationException>(() => FormResourcePreflight.ValidateOleObjectBlob(resource, resource.Length));
        }

        [TestMethod]
        public void ComparisonIgnoresOnlyCompoundAllocationSlackAndTimestamps()
        {
            byte[] original = Resource();
            byte[] other = (byte[])original.Clone();
            // Root directory creation/modification dates and unallocated stream slack.
            other[24 + 1024 + 100] = 31; other[24 + 1024 + 108] = 42;
            other[24 + 2048 + 7] = 99;
            CollectionAssert.AreEqual(FormResourcePreflight.ComparisonBytes(original, new[] { 0 }),
                FormResourcePreflight.ComparisonBytes(other, new[] { 0 }));
            Assert.AreEqual((byte)99, other[24 + 2048 + 7], "Raw transport bytes were modified.");
            // Reallocate the same four logical stream bytes to mini sector 1.
            Write(other, 24 + 1024 + 120, 128);
            Write(other, 24 + 1024 + 128 + 116, 1);
            Write(other, 24 + 1536, uint.MaxValue); Write(other, 24 + 1536 + 4, 0xfffffffe);
            Buffer.BlockCopy(original, 24 + 2048, other, 24 + 2048 + 64, 4);
            CollectionAssert.AreEqual(FormResourcePreflight.ComparisonBytes(original, new[] { 0 }),
                FormResourcePreflight.ComparisonBytes(other, new[] { 0 }));
        }

        [TestMethod]
        public void ComparisonRetainsLogicalStreamEnvelopeNamesClsidStateAndTrailingResources()
        {
            byte[] original = Resource();
            byte[] baseline = FormResourcePreflight.ComparisonBytes(original, new[] { 0 });
            foreach (int offset in new[] { 8, 24 + 8, 24 + 1024 + 80, 24 + 1024 + 96,
                24 + 1024 + 128, 24 + 2048 })
            {
                byte[] other = (byte[])original.Clone(); other[offset]++;
                Assert.IsFalse(System.Linq.Enumerable.SequenceEqual(baseline,
                    FormResourcePreflight.ComparisonBytes(other, new[] { 0 })), "Lost significant byte at " + offset);
            }
            byte[] trailing = (byte[])original.Clone(); Array.Resize(ref trailing, trailing.Length + 1); trailing[trailing.Length - 1] = 7;
            Assert.IsFalse(System.Linq.Enumerable.SequenceEqual(baseline, FormResourcePreflight.ComparisonBytes(trailing, new[] { 0 })));
        }

        [TestMethod]
        public void FormPaddingNormalizationRequiresKnownStorageClassIdentity()
        {
            byte[][] streams = FormStreamPaddingTests.NativeStreams();
            byte[] known = Resource();
            int directory = 24 + 1024, mini = 24 + 2048;
            Buffer.BlockCopy(new Guid("C62A69F0-16DC-11CE-9E98-00AA00574A4F").ToByteArray(), 0, known, directory + 80, 16);
            Write(known, directory + 120, 384);
            Entry(known, directory + 128, "f", 2, uint.MaxValue, 0, (uint)streams[0].Length);
            Write(known, directory + 128 + 72, 2);
            Entry(known, directory + 256, "o", 2, uint.MaxValue, 3, (uint)streams[1].Length);
            for (int i = 0; i < 6; i++) Write(known, 24 + 1536 + i * 4, i == 2 || i == 5 ? 0xfffffffe : (uint)(i + 1));
            Buffer.BlockCopy(streams[0], 0, known, mini, streams[0].Length);
            Buffer.BlockCopy(streams[1], 0, known, mini + 192, streams[1].Length);
            byte[] other = (byte[])known.Clone(); other[mini + 192 + 139] ^= 1;
            CollectionAssert.AreEqual(FormResourcePreflight.ComparisonBytes(known, new[] { 0 }), FormResourcePreflight.ComparisonBytes(other, new[] { 0 }));
            Array.Clear(known, directory + 80, 16); Array.Clear(other, directory + 80, 16);
            Assert.IsFalse(System.Linq.Enumerable.SequenceEqual(FormResourcePreflight.ComparisonBytes(known, new[] { 0 }),
                FormResourcePreflight.ComparisonBytes(other, new[] { 0 })), "Unknown OLE classes must retain every logical stream byte.");
        }

        [TestMethod]
        public void ManifestPreflightRejectsCorruptOleResourcesBeforeAnAdapterCanImport()
        {
            var manifest = new VbaGitManifest { References = "", Components = new[] {
                new VbaGitComponent { Name = "Form1", Type = 3, HasResources = true } } };
            var files = new Dictionary<string, byte[]> {
                ["Form1.frm"] = VbaGitSnapshot.Utf8.GetBytes("VERSION 5.00\nBegin SyntheticForm\n OleObjectBlob = \"Form1.frx\":0000\nEnd\nAttribute VB_Name = \"Form1\"\n"),
                ["Form1.frx"] = Resource() };
            var accepted = new VbaGitSnapshot(manifest, files);
            CollectionAssert.AreEqual(files["Form1.frx"], accepted.Files["Form1.frx"]);
            files["Form1.frx"] = new byte[accepted.Files["Form1.frx"].Length];
            Assert.ThrowsException<InvalidOperationException>(() => new VbaGitSnapshot(manifest, files));
        }

        [TestMethod]
        public void VersionFourNormalStreamsAndEmptyMiniStorageRespectDeclaredSectorSizes()
        {
            var bytes = new byte[24 + 4 * 4096];
            byte[] source = Resource(); Buffer.BlockCopy(source, 0, bytes, 0, 24 + 512);
            Write(bytes, 4, 4 * 4096);
            bytes[24 + 26] = 4; bytes[24 + 30] = 12;
            Write(bytes, 24 + 40, 1); Write(bytes, 24 + 60, 0xfffffffe); Write(bytes, 24 + 64, 0);
            for (int i = 0; i < 1024; i++) Write(bytes, 24 + 4096 + i * 4,
                i == 0 ? 0xfffffffd : i < 3 ? 0xfffffffe : uint.MaxValue);
            Entry(bytes, 24 + 8192, "Root Entry", 5, 1, 0xfffffffe, 0);
            Entry(bytes, 24 + 8192 + 128, "Object", 2, uint.MaxValue, 2, 4096);
            bytes[24 + 12288] = 42;
            FormResourcePreflight.ValidateOleObjectBlob(bytes, 0);
            Write(bytes, 24 + 8192 + 128 + 124, 1); // v4 honors the high stream-size DWORD
            Assert.ThrowsException<InvalidOperationException>(() => FormResourcePreflight.ValidateOleObjectBlob(bytes, 0));
        }

        [TestMethod]
        public void DifatExtensionIsBoundedAndCannotReferenceItselfAsFat()
        {
            var bytes = new byte[24 + 5 * 512];
            byte[] source = Resource(); Buffer.BlockCopy(source, 0, bytes, 0, 24 + 512);
            Write(bytes, 24 + 44, 2); Write(bytes, 24 + 48, 3);
            Write(bytes, 24 + 60, 0xfffffffe); Write(bytes, 24 + 64, 0);
            Write(bytes, 24 + 68, 1); Write(bytes, 24 + 72, 1);
            for (int i = 0; i < 128; i++)
            {
                Write(bytes, 24 + 512 + i * 4,
                    i == 0 || i == 2 ? 0xfffffffd : i == 1 ? 0xfffffffc : i == 3 ? 0xfffffffe : uint.MaxValue);
                Write(bytes, 24 + 1024 + i * 4, i == 0 ? 2u : i == 127 ? 0xfffffffe : uint.MaxValue);
            }
            Entry(bytes, 24 + 2048, "Root Entry", 5, uint.MaxValue, 0xfffffffe, 0);
            FormResourcePreflight.ValidateOleObjectBlob(bytes, 0);
            Write(bytes, 24 + 1024, 1);
            Assert.ThrowsException<InvalidOperationException>(() => FormResourcePreflight.ValidateOleObjectBlob(bytes, 0));
        }

        [TestMethod]
        public void InvalidEnvelopeHeaderPointersCyclesLengthsAndAllocationAliasesAreRefused()
        {
            // Offsets are prescribed by MS-CFB, not calculated by production code.
            var corruptions = new Action<byte[]>[] {
                b => b[0] = 0,
                b => Write(b, 4, uint.MaxValue),
                b => Write(b, 4, 512),
                b => b[24] = 0,
                b => b[24 + 28] = 0,
                b => b[24 + 30] = 10,
                b => Write(b, 24 + 44, uint.MaxValue),
                b => Write(b, 24 + 76, 100),
                b => Write(b, 24 + 80, 0), // duplicate FAT allocation
                b => Write(b, 24 + 48, 0), // directory aliases FAT
                b => Write(b, 24 + 512 + 4, 1), // directory self-cycle
                b => Write(b, 24 + 1024 + 76, 20), // invalid root child
                b => Write(b, 24 + 1024 + 128 + 68, 1), // directory entry self-cycle
                b => { b[24 + 1024 + 128] = 0; b[24 + 1024 + 128 + 1] = 0xd8; }, // unpaired UTF-16 name cannot collapse during comparison
                b => Write(b, 24 + 1024 + 120, uint.MaxValue),
                b => Write(b, 24 + 1024 + 128 + 120, uint.MaxValue),
                b => Write(b, 24 + 1536, 0), // mini-stream self-cycle
                b => Write(b, 24 + 1024 + 128 + 116, 99),
                b => Write(b, 24 + 512 + 3 * 4, 0), // root mini-stream aliases FAT
                b => Write(b, 24 + 68, 2), // empty DIFAT must terminate
                b => Write(b, 24 + 72, uint.MaxValue),
                b => Write(b, 24 + 60, 3), // mini FAT aliases mini stream
                b => Write(b, 24 + 64, uint.MaxValue)
            };
            for (int i = 0; i < corruptions.Length; i++)
            {
                byte[] resource = Resource(); corruptions[i](resource);
                Assert.ThrowsException<InvalidOperationException>(() => FormResourcePreflight.ValidateOleObjectBlob(resource, 0), "Corruption " + i);
            }
            byte[] truncated = Resource(); Array.Resize(ref truncated, truncated.Length - 1);
            Assert.ThrowsException<InvalidOperationException>(() => FormResourcePreflight.ValidateOleObjectBlob(truncated, 0));
        }

        /// <summary>Hand-built synthetic MS-CFB v3 container: FAT, directory, mini FAT and one mini stream.</summary>
        internal static byte[] Resource()
        {
            var bytes = new byte[24 + 5 * 512];
            bytes[0] = 0x4c; bytes[1] = 0x42; bytes[2] = 8;
            Write(bytes, 4, 5 * 512);
            byte[] signature = { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 };
            Buffer.BlockCopy(signature, 0, bytes, 24, 8);
            bytes[24 + 24] = 0x3e; bytes[24 + 26] = 3;
            bytes[24 + 28] = 0xfe; bytes[24 + 29] = 0xff; bytes[24 + 30] = 9; bytes[24 + 32] = 6;
            Write(bytes, 24 + 44, 1); Write(bytes, 24 + 48, 1); Write(bytes, 24 + 56, 4096);
            Write(bytes, 24 + 60, 2); Write(bytes, 24 + 64, 1); Write(bytes, 24 + 68, 0xfffffffe);
            for (int i = 0; i < 109; i++) Write(bytes, 24 + 76 + i * 4, i == 0 ? 0u : uint.MaxValue);
            for (int i = 0; i < 128; i++)
            {
                Write(bytes, 24 + 512 + i * 4, i == 0 ? 0xfffffffd : i < 4 ? 0xfffffffe : uint.MaxValue);
                Write(bytes, 24 + 1536 + i * 4, i == 0 ? 0xfffffffe : uint.MaxValue);
            }
            Entry(bytes, 24 + 1024, "Root Entry", 5, 1, 3, 64);
            Entry(bytes, 24 + 1024 + 128, "f", 2, uint.MaxValue, 0, 4);
            bytes[24 + 2048] = 17; bytes[24 + 2049] = 18; bytes[24 + 2050] = 19; bytes[24 + 2051] = 20;
            return bytes;
        }

        private static void Entry(byte[] bytes, int position, string name, byte type, uint child, uint start, uint size)
        {
            byte[] text = Encoding.Unicode.GetBytes(name + "\0"); Buffer.BlockCopy(text, 0, bytes, position, text.Length);
            bytes[position + 64] = (byte)text.Length; bytes[position + 66] = type; bytes[position + 67] = 1;
            Write(bytes, position + 68, uint.MaxValue); Write(bytes, position + 72, uint.MaxValue);
            Write(bytes, position + 76, child); Write(bytes, position + 116, start); Write(bytes, position + 120, size);
        }

        private static void Write(byte[] bytes, int position, uint value) { Buffer.BlockCopy(BitConverter.GetBytes(value), 0, bytes, position, 4); }
    }
}
