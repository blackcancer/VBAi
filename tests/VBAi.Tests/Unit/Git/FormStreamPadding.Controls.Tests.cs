using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    public sealed partial class FormStreamPaddingTests
    {
        [TestMethod]
        [DataRow(16)] [DataRow(47)] [DataRow(18)] [DataRow(12)]
        public void RetainedFlatControlPairsNormalizeOnlyDocumentedPadding(int type)
        {
            byte[][] input = RetainedControlStreams(type), unchanged = RetainedControlStreams(type, true);
            byte[][] backup = new[] { (byte[])input[0].Clone(), (byte[])input[1].Clone() };
            byte[][] normalized = FormStreamPadding.Normalize(input[0], input[1]);
            byte[][] other = FormStreamPadding.Normalize(unchanged[0], unchanged[1]);
            for (int stream = 0; stream < 2; stream++)
            {
                Assert.AreNotSame(input[stream], normalized[stream]);
                CollectionAssert.AreEqual(normalized[stream], other[stream]);
                CollectionAssert.AreEqual(backup[stream], input[stream]);
                var padding = RetainedControlPadding(type, stream);
                for (int i = 0; i < input[stream].Length; i++)
                    Assert.AreEqual(padding.Contains(i) ? (byte)0 : input[stream][i], normalized[stream][i], "stream " + stream + ", offset " + i);
                foreach (int i in padding) input[stream][i] ^= 0xa5;
            }
            byte[][] alteredPadding = FormStreamPadding.Normalize(input[0], input[1]);
            byte[][] repeated = FormStreamPadding.Normalize(normalized[0], normalized[1]);
            for (int stream = 0; stream < 2; stream++)
            {
                CollectionAssert.AreEqual(normalized[stream], alteredPadding[stream]);
                CollectionAssert.AreEqual(normalized[stream], repeated[stream]);
            }
        }

        [TestMethod]
        [DataRow(16)] [DataRow(47)] [DataRow(18)] [DataRow(12)]
        public void EveryRetainedFlatControlNonPaddingByteRemainsSignificant(int type)
        {
            byte[][] input = RetainedControlStreams(type), expected = FormStreamPadding.Normalize(input[0], input[1]);
            for (int stream = 0; stream < 2; stream++)
            {
                var padding = RetainedControlPadding(type, stream);
                for (int i = 0; i < input[stream].Length; i++)
                {
                    if (padding.Contains(i)) continue;
                    byte[] changed = (byte[])input[stream].Clone(); changed[i] ^= 1;
                    byte[][] actual = FormStreamPadding.Normalize(stream == 0 ? changed : input[0], stream == 1 ? changed : input[1]);
                    Assert.AreNotEqual(expected[stream][i], actual[stream][i], "type " + type + ", stream " + stream + ", offset " + i);
                }
            }
        }

        [TestMethod]
        [DataRow(16)] [DataRow(47)] [DataRow(18)] [DataRow(12)]
        public void AllFlatControlTruncationsAndTrailingBytesRetainBothOriginalStreams(int type)
        {
            byte[][] input = RetainedControlStreams(type);
            for (int stream = 0; stream < 2; stream++)
            {
                for (int length = 0; length < input[stream].Length; length++)
                {
                    byte[] truncated = Slice(input[stream], 0, length);
                    AssertFallback(stream == 0 ? truncated : input[0], stream == 1 ? truncated : input[1]);
                }
                byte[] extended = Join(input[stream], new byte[] { 0x65 });
                AssertFallback(stream == 0 ? extended : input[0], stream == 1 ? extended : input[1]);
            }
        }

        [TestMethod]
        [DataRow(16, 4, 3)] [DataRow(47, 8, 3)] [DataRow(18, 3, 4)] [DataRow(12, 0, 9)]
        public void ReservedMaskBitsAndMissingRequiredSizeAreConservative(int type, int unused, int required)
        {
            byte[][] input = RetainedControlStreams(type); int start = type == 12 ? 128 : 132;
            uint mask = BitConverter.ToUInt32(input[1], start + 4);
            byte[] changed = (byte[])input[1].Clone(); Write32(changed, start + 4, mask | (1u << unused));
            AssertFallback(input[0], changed);
            changed = (byte[])input[1].Clone(); Write32(changed, start + 4, mask & ~(1u << required));
            AssertFallback(input[0], changed);
            if (type == 18)
            {
                changed = (byte[])input[1].Clone(); Write32(changed, start + 4, mask & ~(1u << 19));
                AssertFallback(input[0], changed);
            }
        }

        [TestMethod]
        [DataRow(16)] [DataRow(47)] [DataRow(18)] [DataRow(12)]
        public void RichFlatControlsReadEveryOptionalFieldAndOnlyClearAlignment(int type)
        {
            byte[] leaf = RichFlatControl(type), form = SingleControlForm((uint)type, leaf.Length);
            byte[] formBackup = (byte[])form.Clone(), leafBackup = (byte[])leaf.Clone();
            byte[][] normalized = FormStreamPadding.Normalize(form, leaf);
            Assert.AreNotSame(form, normalized[0]); Assert.AreNotSame(leaf, normalized[1]);
            CollectionAssert.AreEqual(formBackup, form); CollectionAssert.AreEqual(leafBackup, leaf);
            for (int stream = 0; stream < 2; stream++)
            {
                byte[] bytes = stream == 0 ? form : leaf;
                for (int i = 0; i < bytes.Length; i++)
                    if (bytes[i] != normalized[stream][i])
                    {
                        Assert.AreEqual((byte)0xe7, bytes[i], "Only fixture padding is0xE7 at " + i);
                        Assert.AreEqual((byte)0, normalized[stream][i]);
                    }
            }
            // Headers, property values, both picture envelopes/payloads and array strings remain significant.
            for (int i = 0; i < leaf.Length; i++)
            {
                if (leaf[i] != normalized[1][i]) continue;
                byte[] changed = (byte[])leaf.Clone(); changed[i] ^= 1;
                Assert.AreNotEqual(normalized[1][i], FormStreamPadding.Normalize(form, changed)[1][i], "type " + type + ", byte " + i);
            }
        }

        [TestMethod]
        [DataRow(16, 8)] [DataRow(47, 9)]
        public void SpinAndScrollEnabledMaskDependenciesRejectInconsistentPreviousOrNext(int type, int previous)
        {
            byte[] leaf = RichFlatControl(type), form = SingleControlForm((uint)type, leaf.Length);
            uint mask = BitConverter.ToUInt32(leaf, 4);
            foreach (int bit in new[] { previous, previous + 1 })
            {
                byte[] changed = (byte[])leaf.Clone(); Write32(changed, 4, mask & ~(1u << bit)); AssertFallback(form, changed);
            }
            byte[] enabled = (byte[])leaf.Clone(); Write32(enabled, 16, 0x1b); AssertFallback(form, enabled);
        }

        [TestMethod]
        [DataRow(16)] [DataRow(47)] [DataRow(18)] [DataRow(12)]
        public void PictureEnvelopeUnknownGuidSentinelPreambleAndOversizeNeverPartiallyNormalize(int type)
        {
            byte[] leaf = RichFlatControl(type), form = SingleControlForm((uint)type, leaf.Length);
            byte[] guid = new Guid("0BE35204-8F91-11CE-9DE3-00AA004BB851").ToByteArray();
            int offset = FindBytes(leaf, guid); Assert.IsTrue(offset >= 0);
            foreach (int i in new[] { offset, offset + 16, offset + 20 })
            {
                byte[] changed = (byte[])leaf.Clone(); changed[i] ^= 1;
                if (i == offset + 20) Write32(changed, i, uint.MaxValue);
                AssertFallback(form, changed);
            }
            // Every byte of this envelope's retained102-byte BMP payload must survive.
            for (int i = offset + 24; i < offset + 24 + 102; i++)
            {
                byte[] changed = (byte[])leaf.Clone(); changed[i] ^= 1;
                Assert.AreEqual(changed[i], FormStreamPadding.Normalize(form, changed)[1][i]);
            }
            // Main-block picture/icon sentinels are outside the envelope and remain validated.
            int blockEnd = 4 + BitConverter.ToUInt16(leaf, 2);
            for (int i = 8; i < blockEnd - 1; i++)
                if (leaf[i] == 0xff && leaf[i + 1] == 0xff)
                {
                    byte[] changed = (byte[])leaf.Clone(); changed[i] = 0; AssertFallback(form, changed);
                }
        }

        [TestMethod]
        public void TabArraysUseCharacterLengthsAndPreserveUnicodeZeroLengthAndExactFlagOrder()
        {
            byte[] leaf = RichFlatControl(18), form = SingleControlForm(18, leaf.Length);
            Assert.AreNotSame(leaf, FormStreamPadding.Normalize(form, leaf)[1]);
            // ArrayString uncompressed count3 is six bytes, not three, and includes nonASCII U+03A9.
            Assert.IsTrue(FindBytes(leaf, Encoding.Unicode.GetBytes("AΩZ")) >= 0);
            byte[] zero = SimpleTab(ArrayString(new byte[0], true), 1, 1);
            Assert.AreNotSame(zero, FormStreamPadding.Normalize(SingleControlForm(18, zero.Length), zero)[1]);
            byte[] noArrays = SimpleTab(null, 0, 0);
            Assert.AreNotSame(noArrays, FormStreamPadding.Normalize(SingleControlForm(18, noArrays.Length), noArrays)[1]);
            byte[] invalidFlag = (byte[])leaf.Clone(); Write32(invalidFlag, invalidFlag.Length - 4, 4); AssertFallback(form, invalidFlag);
            byte[] wrongCount = SimpleTab(ArrayString(new byte[] { 65 }, true), 2, 1);
            AssertFallback(SingleControlForm(18, wrongCount.Length), wrongCount);
            byte[] tooFewAllocated = SimpleTab(ArrayString(new byte[] { 65 }, true), 1, 0);
            AssertFallback(SingleControlForm(18, tooFewAllocated.Length), tooFewAllocated);
        }

        [TestMethod]
        public void TabArraySizeOverflowTruncationAndMismatchedArrayCountsRetainOriginals()
        {
            byte[] leaf = RichFlatControl(18), form = SingleControlForm(18, leaf.Length);
            // Rich tab fields: ListIndex/BackColor/ForeColor/ItemsSize; verify offset before mutation.
            Assert.AreEqual(20u, BitConverter.ToUInt32(leaf, 20));
            foreach (uint size in new[] { 0u, 1u, 19u, 21u, uint.MaxValue })
            {
                byte[] changed = (byte[])leaf.Clone(); Write32(changed, 20, size); AssertFallback(form, changed);
            }
            // Array two begins after main data+Size. Corrupt its character count beyond bounded length.
            int descriptor = FindBytes(leaf, new byte[] { 3, 0, 0, 0, 65, 0, 0xa9, 3, 90, 0 });
            Assert.IsTrue(descriptor >= 0);
            byte[] malformed = (byte[])leaf.Clone(); Write32(malformed, descriptor, 0x7fffffff); AssertFallback(form, malformed);
            byte[] oneArray = ArrayString(new byte[] { 65 }, true);
            var b = new BlockBuilder(0x200, 0x000a0030); // Items+Names, required Size/NewVersion.
            b.Field(4, (uint)oneArray.Length); b.Field(4, (uint)(oneArray.Length * 2)); b.Align(4); b.Bytes(new byte[8]);
            b.Bytes(oneArray); b.Bytes(Join(oneArray, oneArray));
            malformed = Join(b.Finish(), RichText()); AssertFallback(SingleControlForm(18, malformed.Length), malformed);
        }

        private static byte[] RichFlatControl(int type)
        {
            if (type == 16 || type == 47)
            {
                bool scroll = type == 47; var b = new BlockBuilder(0x200, scroll ? 0x1feffu : 0x7fefu);
                b.Field(4, 0x112233); b.Field(4, 0x445566); b.Field(4, 0x19); // Disabled => both Prev/Next present.
                if (scroll) b.Field(1, 1);
                for (int i = 0; i < 8; i++) b.Field(4, (uint)(i + 1));
                if (scroll) { b.Field(2, 1); b.Field(4, 50); }
                b.Field(2, 0xffff); if (!scroll) b.Field(1, 1);
                b.Align(4); b.Bytes(new byte[8]); return Join(b.Finish(), TestPicture());
            }
            if (type == 12)
            {
                var b = new BlockBuilder(0x200, 0x7ffcu);
                b.Field(4, 0x112233); b.Field(4, 0x445566);
                for (int i = 0; i < 4; i++) b.Field(1, 1);
                b.Field(2, 0xffff); b.Field(1, 1); b.Field(4, 0x1b); b.Field(2, 0xffff);
                b.Align(4); b.Bytes(new byte[8]); return Join(Join(b.Finish(), TestPicture()), TestPicture());
            }
            if (type == 18)
            {
                byte[] array = Join(ArrayString(new byte[] { 65, 66, 67 }, true), ArrayString(Encoding.Unicode.GetBytes("AΩZ"), false));
                var b = new BlockBuilder(0x200, 0x1febf77u);
                b.Field(4, 0); b.Field(4, 0x112233); b.Field(4, 0x445566); b.Field(4, (uint)array.Length); b.Field(1, 1);
                for (int i = 0; i < 4; i++) b.Field(4, 1);
                b.Field(4, (uint)array.Length); b.Field(4, (uint)array.Length); b.Field(4, 0x1b); b.Field(4, 2);
                b.Field(4, (uint)array.Length); b.Field(4, 2); b.Field(4, (uint)array.Length); b.Field(2, 0xffff);
                b.Align(4); b.Bytes(new byte[8]); for (int i = 0; i < 5; i++) b.Bytes(array);
                return Join(Join(Join(b.Finish(), TestPicture()), RichText()), new byte[] { 1, 0, 0, 0, 2, 0, 0, 0 });
            }
            throw new ArgumentOutOfRangeException(nameof(type));
        }

        private static byte[] SimpleTab(byte[] array, uint count, uint allocated)
        {
            var b = new BlockBuilder(0x200, array == null ? 0x80010u : 0x580030u);
            if (array != null) { b.Field(4, (uint)array.Length); b.Field(4, allocated); b.Field(4, count); }
            b.Align(4); b.Bytes(new byte[8]); if (array != null) b.Bytes(array);
            byte[] result = Join(b.Finish(), RichText());
            for (uint i = 0; i < count; i++) result = Join(result, new byte[] { 3, 0, 0, 0 });
            return result;
        }

        private static byte[] SingleControlForm(uint type, int size)
        {
            var header = new BlockBuilder(0x400, 0x08000040); header.Field(4, 0x8004); header.Field(4, 1);
            var site = new BlockBuilder(0, 0xe0); site.Field(4, (uint)size); site.Field(2, 0); site.Field(2, type);
            using (var stream = new MemoryStream()) using (var writer = new BinaryWriter(stream))
            {
                byte[] child = site.Finish(); writer.Write(header.Finish()); writer.Write(1u); writer.Write((uint)(4 + child.Length));
                writer.Write(new byte[] { 0, 1, 0xe7, 0xe7 }); writer.Write(child); return stream.ToArray();
            }
        }

        private static byte[] ArrayString(byte[] bytes, bool compressed)
        {
            Assert.IsTrue(compressed || bytes.Length % 2 == 0);
            using (var stream = new MemoryStream()) using (var writer = new BinaryWriter(stream))
            {
                writer.Write((uint)(compressed ? bytes.Length : bytes.Length / 2) | (compressed ? 0x80000000u : 0)); writer.Write(bytes);
                while (stream.Length % 4 != 0) writer.Write((byte)0xe7); return stream.ToArray();
            }
        }

        private static byte[] TestPicture()
        {
            byte[] payload = Slice(Hex(ImageObjects), 172, 102);
            using (var stream = new MemoryStream()) using (var writer = new BinaryWriter(stream))
            {
                writer.Write(new Guid("0BE35204-8F91-11CE-9DE3-00AA004BB851").ToByteArray()); writer.Write(0x746cu);
                writer.Write((uint)payload.Length); writer.Write(payload); return stream.ToArray();
            }
        }

        private static byte[] Slice(byte[] bytes, int start, int length)
        {
            var result = new byte[length]; Array.Copy(bytes, start, result, 0, length); return result;
        }

        private static int FindBytes(byte[] bytes, byte[] wanted)
        {
            for (int i = 0; i <= bytes.Length - wanted.Length; i++)
            {
                int j = 0; while (j < wanted.Length && bytes[i + j] == wanted[j]) j++;
                if (j == wanted.Length) return i;
            }
            return -1;
        }
    }
}
