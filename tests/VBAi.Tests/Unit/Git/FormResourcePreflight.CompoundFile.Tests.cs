using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;

namespace VBAi.Tests.Unit
{
    public sealed partial class FormResourcePreflightTests
    {
        [DataTestMethod]
        [DataRow("fat marker")]
        [DataRow("directory absent")]
        [DataRow("root kind")]
        [DataRow("root name")]
        [DataRow("root left")]
        [DataRow("root right")]
        [DataRow("duplicate path")]
        [DataRow("unreachable entry")]
        [DataRow("stream child")]
        [DataRow("unknown directory kind")]
        [DataRow("short name")]
        [DataRow("long name")]
        [DataRow("odd name")]
        [DataRow("unterminated name")]
        [DataRow("empty name")]
        [DataRow("forbidden name")]
        [DataRow("missing mini FAT")]
        [DataRow("short root chain")]
        public void EveryReachableCompoundDirectoryAndAllocationGuardRefusesMalformedPreflight(string defect)
        {
            byte[] bytes = Resource();
            FormResourcePreflight.ValidateOleObjectBlob(bytes, 0);
            const int header = 24, fat = 24 + 512, directory = 24 + 1024, leaf = directory + 128;
            switch (defect)
            {
                case "fat marker": Write(bytes, fat, 0xfffffffe); break;
                case "directory absent": Write(bytes, header + 48, 0xfffffffe); break;
                case "root kind": bytes[directory + 66] = 1; break;
                case "root name": Entry(bytes, directory, "Wrong root", 5, 1, 3, 64); break;
                case "root left": Write(bytes, directory + 68, 0); break;
                case "root right": Write(bytes, directory + 72, 0); break;
                case "duplicate path":
                    Entry(bytes, directory + 256, "F", 2, uint.MaxValue, 0xfffffffe, 0);
                    Write(bytes, leaf + 72, 2); break;
                case "unreachable entry": Entry(bytes, directory + 256, "orphan", 2, uint.MaxValue, 0xfffffffe, 0); break;
                case "stream child": Write(bytes, leaf + 76, 0); break;
                case "unknown directory kind": bytes[leaf + 66] = 6; break;
                case "short name": Write16(bytes, leaf + 64, 0); break;
                case "long name": Write16(bytes, leaf + 64, 66); break;
                case "odd name": Write16(bytes, leaf + 64, 3); break;
                case "unterminated name": bytes[leaf + 2] = 65; break;
                case "empty name": Write16(bytes, leaf, 0); Write16(bytes, leaf + 64, 2); break;
                case "forbidden name": Entry(bytes, leaf, "bad/name", 2, uint.MaxValue, 0, 4); break;
                case "missing mini FAT": Write(bytes, header + 60, 0xfffffffe); Write(bytes, header + 64, 0); break;
                case "short root chain": Write(bytes, directory + 120, 513); break;
                default: Assert.Fail("Unknown prepared corruption: " + defect); break;
            }
            AssertInvalidCompoundResource(bytes, defect);
        }

        [TestMethod]
        public void ExactFatCapacityAcceptsAFullSectorButInsufficientOrMissingFatIsRefused()
        {
            // 128 FAT slots cover exactly 128 physical sectors, including allocated and free sectors.
            byte[] exact = ExtendResource(128);
            FormResourcePreflight.ValidateOleObjectBlob(exact, 0);
            CollectionAssert.AreEqual(FormResourcePreflight.ComparisonBytes(Resource(), new[] { 0 }),
                FormResourcePreflight.ComparisonBytes(exact, new[] { 0 }), "Free allocation does not alter logical data.");
            AssertInvalidCompoundResource(ExtendResource(129), "One FAT cannot cover 129 physical sectors.");
            byte[] missing = Resource(); Write(missing, 24 + 44, 0); Write(missing, 24 + 76, uint.MaxValue);
            AssertInvalidCompoundResource(missing, "A compound file requires a FAT even when the declared FAT count is zero.");
        }

        [TestMethod]
        public void DifatSectorMustCarryItsOwnAllocationMarker()
        {
            byte[] bytes = EmptyDifatResource(); FormResourcePreflight.ValidateOleObjectBlob(bytes, 0);
            Write(bytes, 24 + 512 + 4, 0xfffffffe);
            AssertInvalidCompoundResource(bytes, "The declared DIFAT sector has an ordinary-chain allocation marker.");
        }

        [TestMethod]
        public void VersionFourDeclaredDirectoryCountMustEqualItsActualChainLength()
        {
            byte[] bytes = EmptyVersionFourResource(); FormResourcePreflight.ValidateOleObjectBlob(bytes, 0);
            Write(bytes, 24 + 40, 2);
            AssertInvalidCompoundResource(bytes, "Version4 declares two directory sectors but has one.");
        }

        [TestMethod]
        public void EmptyMiniStreamAndTwoSectorRootMiniStreamUseTheirExactDeclaredLengths()
        {
            byte[] empty = Resource(); Write(empty, 24 + 60, 0xfffffffe); Write(empty, 24 + 64, 0);
            Write(empty, 24 + 1024 + 116, 0xfffffffe); Write(empty, 24 + 1024 + 120, 0);
            Write(empty, 24 + 1024 + 128 + 116, 0xfffffffe); Write(empty, 24 + 1024 + 128 + 120, 0);
            FormResourcePreflight.ValidateOleObjectBlob(empty, 0);
            byte[] twoSectors = ExtendResource(5); Write(twoSectors, 24 + 1024 + 120, 513);
            Write(twoSectors, 24 + 512 + 3 * 4, 4); Write(twoSectors, 24 + 512 + 4 * 4, 0xfffffffe);
            FormResourcePreflight.ValidateOleObjectBlob(twoSectors, 0);
            Write(twoSectors, 24 + 512 + 3 * 4, 0xfffffffe);
            AssertInvalidCompoundResource(twoSectors, "Root stream has one sector but requires two, despite enough file capacity.");
        }

        [TestMethod]
        public void CumulativeDirectoryPathBudgetIsBoundedIndependentlyOfSmallRawResourceSize()
        {
            // Identical 31-character names are valid at different levels. No sibling/path duplicate is introduced.
            // Sum of visited path lengths is32*(n*(n+1)/2), below32MiB for1400 and above it for1500.
            byte[] below = NestedStorageResource(1400), above = NestedStorageResource(1500);
            Assert.IsTrue(below.Length < 256 * 1024 && above.Length < 256 * 1024);
            FormResourcePreflight.ValidateOleObjectBlob(below, 0);
            AssertInvalidCompoundResource(above, "Deep path expansion exceeds the comparison's32MiB directory path budget.");
        }

        [TestMethod]
        public void MultipleResourceOffsetsPreserveGapsAndTrailingBytesButRejectOverlappingBlobs()
        {
            byte[] blob = Resource(); var bytes = new byte[blob.Length * 2 + 11];
            Buffer.BlockCopy(blob, 0, bytes, 0, blob.Length); Buffer.BlockCopy(blob, 0, bytes, blob.Length + 7, blob.Length);
            bytes[blob.Length + 2] = 41; bytes[bytes.Length - 1] = 52;
            byte[] original = (byte[])bytes.Clone();
            byte[] expected = FormResourcePreflight.ComparisonBytes(bytes, new[] { 0, blob.Length + 7 });
            CollectionAssert.AreEqual(expected, FormResourcePreflight.ComparisonBytes(bytes, new[] { blob.Length + 7, 0, 0 }));
            bytes[blob.Length + 2] ^= 1;
            Assert.IsFalse(expected.SequenceEqual(FormResourcePreflight.ComparisonBytes(bytes, new[] { 0, blob.Length + 7 })));
            bytes = original; bytes[bytes.Length - 1] ^= 1;
            Assert.IsFalse(expected.SequenceEqual(FormResourcePreflight.ComparisonBytes(bytes, new[] { 0, blob.Length + 7 })));
            Assert.ThrowsException<InvalidOperationException>(() => FormResourcePreflight.ComparisonBytes(bytes, new[] { 0, 10 }));
        }

        private static void AssertInvalidCompoundResource(byte[] bytes, string reason)
        {
            byte[] original = (byte[])bytes.Clone();
            var error = Assert.ThrowsException<InvalidOperationException>(() => FormResourcePreflight.ValidateOleObjectBlob(bytes, 0), reason);
            Assert.AreEqual("Invalid or truncated UserForm OLE resource container.", error.Message);
            Assert.ThrowsException<InvalidOperationException>(() => FormResourcePreflight.ComparisonBytes(bytes, new[] { 0 }), reason);
            CollectionAssert.AreEqual(original, bytes, "A refusal must preserve exact transport bytes: " + reason);
        }

        private static byte[] ExtendResource(int sectors)
        {
            var bytes = new byte[24 + (sectors + 1) * 512]; byte[] original = Resource();
            Buffer.BlockCopy(original, 0, bytes, 0, original.Length); Write(bytes, 4, (uint)(bytes.Length - 24)); return bytes;
        }

        private static byte[] EmptyDifatResource()
        {
            var bytes = new byte[24 + 5 * 512]; Buffer.BlockCopy(Resource(), 0, bytes, 0, 24 + 512);
            Write(bytes, 4, 5 * 512); Write(bytes, 24 + 44, 2); Write(bytes, 24 + 48, 3);
            Write(bytes, 24 + 60, 0xfffffffe); Write(bytes, 24 + 64, 0); Write(bytes, 24 + 68, 1); Write(bytes, 24 + 72, 1);
            for (int i = 0; i < 128; i++)
            {
                Write(bytes, 24 + 512 + i * 4, i == 0 || i == 2 ? 0xfffffffd : i == 1 ? 0xfffffffc : i == 3 ? 0xfffffffe : uint.MaxValue);
                Write(bytes, 24 + 1024 + i * 4, i == 0 ? 2u : i == 127 ? 0xfffffffe : uint.MaxValue);
            }
            Entry(bytes, 24 + 2048, "Root Entry", 5, uint.MaxValue, 0xfffffffe, 0); return bytes;
        }

        private static byte[] EmptyVersionFourResource()
        {
            var bytes = new byte[24 + 3 * 4096]; Buffer.BlockCopy(Resource(), 0, bytes, 0, 24 + 512);
            Write(bytes, 4, 3 * 4096); bytes[24 + 26] = 4; bytes[24 + 30] = 12;
            Write(bytes, 24 + 40, 1); Write(bytes, 24 + 60, 0xfffffffe); Write(bytes, 24 + 64, 0);
            for (int i = 0; i < 1024; i++) Write(bytes, 24 + 4096 + i * 4, i == 0 ? 0xfffffffd : i == 1 ? 0xfffffffe : uint.MaxValue);
            Entry(bytes, 24 + 8192, "Root Entry", 5, uint.MaxValue, 0xfffffffe, 0); return bytes;
        }

        /// <summary>Hand-built MS-CFB v3 directory chain with no streams; file size remains bounded independently of path expansion.</summary>
        private static byte[] NestedStorageResource(int children)
        {
            int directories = (children + 4) / 4, fats = (directories + 126) / 127, sectors = directories + fats;
            var bytes = new byte[24 + (sectors + 1) * 512]; Buffer.BlockCopy(Resource(), 0, bytes, 0, 24 + 512);
            Write(bytes, 4, (uint)(bytes.Length - 24)); Write(bytes, 24 + 44, (uint)fats); Write(bytes, 24 + 48, (uint)fats);
            Write(bytes, 24 + 60, 0xfffffffe); Write(bytes, 24 + 64, 0);
            for (int i = 0; i < 109; i++) Write(bytes, 24 + 76 + i * 4, i < fats ? (uint)i : uint.MaxValue);
            for (int i = 0; i < fats * 128; i++)
                Write(bytes, 24 + 512 + i * 4, i < fats ? 0xfffffffd : i < sectors - 1 ? (uint)(i + 1) : i == sectors - 1 ? 0xfffffffe : uint.MaxValue);
            int directory = 24 + (fats + 1) * 512;
            Entry(bytes, directory, "Root Entry", 5, children == 0 ? uint.MaxValue : 1, 0xfffffffe, 0);
            for (int i = 1; i <= children; i++) Entry(bytes, directory + i * 128, new string('S', 31), 1,
                i == children ? uint.MaxValue : (uint)(i + 1), 0xfffffffe, 0);
            return bytes;
        }

        private static void Write16(byte[] bytes, int position, ushort value)
        { Buffer.BlockCopy(BitConverter.GetBytes(value), 0, bytes, position, 2); }
    }
}
