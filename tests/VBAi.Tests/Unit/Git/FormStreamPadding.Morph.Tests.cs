using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    public sealed partial class FormStreamPaddingTests
    {
        [DataTestMethod]
        [DataRow(23)] [DataRow(24)] [DataRow(25)] [DataRow(26)] [DataRow(27)] [DataRow(28)]
        public void MorphMinimalAndAllApplicableFieldsKeepExactValuesAndNormalizeOnlyPadding(int type)
        {
            foreach (bool rich in new[] { false, true })
            {
                MorphSample sample = MakeMorph(type, rich);
                byte[] form = MorphForm(new[] { type }, new[] { sample.Bytes.Length });
                byte[] original = (byte[])sample.Bytes.Clone();
                byte[][] result = FormStreamPadding.Normalize(form, sample.Bytes);
                Assert.AreNotSame(form, result[0], "Type " + type + ", rich=" + rich);
                Assert.AreNotSame(sample.Bytes, result[1]);
                CollectionAssert.AreEqual(original, sample.Bytes);
                for (int i = 0; i < original.Length; i++)
                    Assert.AreEqual(sample.Padding.Contains(i) ? (byte)0 : original[i], result[1][i], "Byte " + i);
                foreach (int offset in sample.Padding) sample.Bytes[offset] ^= 0x3a;
                byte[][] altered = FormStreamPadding.Normalize(form, sample.Bytes);
                CollectionAssert.AreEqual(result[1], altered[1]);
                byte[][] repeated = FormStreamPadding.Normalize(result[0], result[1]);
                CollectionAssert.AreEqual(result[0], repeated[0]); CollectionAssert.AreEqual(result[1], repeated[1]);
            }
        }

        [DataTestMethod]
        [DataRow(23)] [DataRow(24)] [DataRow(25)] [DataRow(26)] [DataRow(27)] [DataRow(28)]
        public void EveryMorphPropertyAndFontByteRemainsSignificant(int type)
        {
            MorphSample sample = MakeMorph(type, true, unicode: true);
            byte[] form = MorphForm(new[] { type }, new[] { sample.Bytes.Length });
            byte[][] baseline = FormStreamPadding.Normalize(form, sample.Bytes);
            Assert.AreNotSame(sample.Bytes, baseline[1]);
            for (int i = 0; i < sample.Bytes.Length; i++)
            {
                if (sample.Padding.Contains(i)) continue;
                byte[] changed = (byte[])sample.Bytes.Clone(); changed[i] ^= 1;
                byte[][] normalized = FormStreamPadding.Normalize(form, changed);
                Assert.AreNotEqual(baseline[1][i], normalized[1][i], "Type " + type + ", byte " + i);
            }
        }

        [DataTestMethod]
        [DataRow(23)] [DataRow(24)] [DataRow(25)] [DataRow(26)] [DataRow(27)] [DataRow(28)]
        public void EachOptionalMorphPropertyUsesNaturalAlignmentOnlyWhenPresent(int type)
        {
            foreach (int bit in MorphFieldBits(type))
            {
                MorphSample sample = MakeMorph(type, false, extraBit: bit);
                byte[] form = MorphForm(new[] { type }, new[] { sample.Bytes.Length });
                byte[][] normalized = FormStreamPadding.Normalize(form, sample.Bytes);
                Assert.AreNotSame(sample.Bytes, normalized[1], "Type " + type + ", single field " + bit);
                foreach (int offset in sample.Padding) Assert.AreEqual((byte)0, normalized[1][offset]);
            }
            // A TextBox may omit DisplayStyle: the documented default is TextBox, not a byte read from Size.
            if (type == 23)
            {
                MorphSample sample = MakeMorph(type, false, omitStyle: true);
                byte[] form = MorphForm(new[] { type }, new[] { sample.Bytes.Length });
                Assert.AreNotSame(sample.Bytes, FormStreamPadding.Normalize(form, sample.Bytes)[1]);
            }
        }

        [DataTestMethod]
        [DataRow(23)] [DataRow(24)] [DataRow(25)] [DataRow(26)] [DataRow(27)] [DataRow(28)]
        public void MorphExtraDataUsesSizeThenValueCaptionGroupAndExactStringLengths(int type)
        {
            foreach (bool unicode in new[] { false, true })
            foreach (bool empty in new[] { false, true })
            {
                MorphSample sample = MakeMorph(type, true, unicode: unicode, empty: empty);
                byte[] form = MorphForm(new[] { type }, new[] { sample.Bytes.Length });
                byte[][] result = FormStreamPadding.Normalize(form, sample.Bytes);
                Assert.AreNotSame(sample.Bytes, result[1]);
                for (int i = sample.SizeOffset; i < sample.SizeOffset + 8; i++) Assert.AreEqual(sample.Bytes[i], result[1][i]);
                foreach (int bit in new[] { 22, 23, 32 })
                {
                    if (!sample.Fields.ContainsKey(bit)) continue;
                    foreach (uint invalid in new[] { 1u, 3u, 0x7fffffffu, 0xffffffffu })
                    {
                        byte[] changed = (byte[])sample.Bytes.Clone(); Write32(changed, sample.Fields[bit], invalid);
                        AssertFallback(form, changed);
                    }
                }
            }
        }

        [DataTestMethod]
        [DataRow(23)] [DataRow(24)] [DataRow(25)] [DataRow(26)] [DataRow(27)] [DataRow(28)]
        public void MorphUnknownMasksMissingRequiredBitsAndInapplicableFieldsRefuseBothStreams(int type)
        {
            MorphSample sample = MakeMorph(type, true);
            byte[] form = MorphForm(new[] { type }, new[] { sample.Bytes.Length });
            foreach (int bit in new[] { 8, 31 })
            {
                byte[] changed = (byte[])sample.Bytes.Clone(); Write32(changed, 4, BitConverter.ToUInt32(changed, 4) & ~(1u << bit));
                AssertFallback(form, changed);
            }
            foreach (int bit in new[] { 19, 30 })
            {
                byte[] changed = (byte[])sample.Bytes.Clone(); Write32(changed, 4, BitConverter.ToUInt32(changed, 4) | (1u << bit));
                AssertFallback(form, changed);
            }
            for (int bit = 1; bit < 32; bit++)
            {
                byte[] changed = (byte[])sample.Bytes.Clone(); Write32(changed, 8, BitConverter.ToUInt32(changed, 8) | (1u << bit));
                AssertFallback(form, changed);
            }
            var applicable = new HashSet<int>(MorphFieldBits(type));
            for (int bit = 0; bit <= 32; bit++)
            {
                if (applicable.Contains(bit) || bit == 8 || bit == 19 || bit == 30 || bit == 31) continue;
                MorphSample wrong = MakeMorph(type, true, extraBit: bit);
                AssertFallback(MorphForm(new[] { type }, new[] { wrong.Bytes.Length }), wrong.Bytes);
            }
            foreach (int picture in new[] { 27, 28 })
            {
                MorphSample unsupported = MakeMorph(type, true, extraBit: picture);
                AssertFallback(MorphForm(new[] { type }, new[] { unsupported.Bytes.Length }), unsupported.Bytes);
            }
        }

        [DataTestMethod]
        [DataRow(23)] [DataRow(24)] [DataRow(25)] [DataRow(26)] [DataRow(27)] [DataRow(28)]
        public void MorphCachedTypeDisplayStyleAndMandatoryValuesCannotDisagree(int type)
        {
            MorphSample sample = MakeMorph(type, true);
            byte[] form = MorphForm(new[] { type }, new[] { sample.Bytes.Length });
            for (int style = 0; style <= 8; style++)
            {
                byte[] changed = (byte[])sample.Bytes.Clone(); changed[sample.Fields[6]] = (byte)style;
                if (style == type - 22 || (type == 25 && style == 7))
                    Assert.AreNotSame(changed, FormStreamPadding.Normalize(form, changed)[1]);
                else AssertFallback(form, changed);
            }
            byte[] foreign = MorphForm(new[] { 15 }, new[] { sample.Bytes.Length });
            AssertFallback(foreign, sample.Bytes); // Generic MorphData needs an explicitly supported identity route.
            if (type == 24)
            {
                byte[] changed = (byte[])sample.Bytes.Clone(); changed[sample.Fields[5]] = 0;
                AssertFallback(form, changed);
            }
            if (type == 28)
            {
                byte[] changed = (byte[])sample.Bytes.Clone(); Write32(changed, sample.Fields[26], 0);
                AssertFallback(form, changed);
            }
            foreach (int bit in new[] { 12, 13 })
            {
                if (!sample.Fields.ContainsKey(bit)) continue;
                byte[] changed = (byte[])sample.Bytes.Clone();
                Array.Copy(BitConverter.GetBytes((ushort)0xfffe), 0, changed, sample.Fields[bit], 2);
                AssertFallback(form, changed);
            }
        }

        [DataTestMethod]
        [DataRow(24)] [DataRow(25)]
        public void MorphColumnsFollowTextPropertiesAndPreserveEverySignedWidth(int type)
        {
            MorphSample sample = MakeMorph(type, true, widths: new int?[] { null, -1, 0, 1900, int.MinValue, int.MaxValue });
            byte[] form = MorphForm(new[] { type }, new[] { sample.Bytes.Length });
            byte[][] result = FormStreamPadding.Normalize(form, sample.Bytes);
            Assert.AreNotSame(sample.Bytes, result[1]);
            for (int i = sample.ColumnStart; i < sample.Bytes.Length; i++) Assert.AreEqual(sample.Bytes[i], result[1][i]);
            foreach (int offset in sample.ColumnOffsets)
            {
                foreach (int fieldOffset in new[] { 0, 1, 2, 4 })
                {
                    byte[] changed = (byte[])sample.Bytes.Clone(); changed[offset + fieldOffset] ^= 0x40;
                    AssertFallback(form, changed);
                }
            }
            byte[] mismatched = (byte[])sample.Bytes.Clone();
            Array.Copy(BitConverter.GetBytes((ushort)7), 0, mismatched, sample.Fields[15], 2);
            AssertFallback(form, mismatched);
            Array.Copy(BitConverter.GetBytes((ushort)5), 0, mismatched, sample.Fields[15], 2);
            AssertFallback(form, mismatched);
            Array.Copy(BitConverter.GetBytes(ushort.MaxValue), 0, mismatched, sample.Fields[15], 2);
            AssertFallback(form, mismatched);
        }

        [DataTestMethod]
        [DataRow(23)] [DataRow(24)] [DataRow(25)] [DataRow(26)] [DataRow(27)] [DataRow(28)]
        public void EveryMorphTruncationHeaderAndTrailingExtensionReturnsOriginalPair(int type)
        {
            MorphSample sample = MakeMorph(type, true, unicode: true);
            byte[] form = MorphForm(new[] { type }, new[] { sample.Bytes.Length });
            for (int i = 0; i < form.Length; i++) AssertFallback(form.Take(i).ToArray(), sample.Bytes);
            for (int i = 0; i < sample.Bytes.Length; i++) AssertFallback(form, sample.Bytes.Take(i).ToArray());
            foreach (int offset in new[] { 0, 1, 2, 3, sample.FontStart, sample.FontStart + 1, sample.FontStart + 2 })
            {
                byte[] changed = (byte[])sample.Bytes.Clone(); changed[offset] ^= 0x40;
                AssertFallback(form, changed);
            }
            foreach (ushort size in new ushort[] { 0, 4, 7, ushort.MaxValue })
            {
                byte[] changed = (byte[])sample.Bytes.Clone(); Array.Copy(BitConverter.GetBytes(size), 0, changed, 2, 2);
                AssertFallback(form, changed);
            }
            byte[] extended = Join(sample.Bytes, new byte[] { 0 });
            AssertFallback(MorphForm(new[] { type }, new[] { extended.Length }), extended);
            AssertFallback(Join(form, new byte[] { 0 }), sample.Bytes);
            var next = MakeMorph(23, true);
            byte[] combined = Join(sample.Bytes, next.Bytes);
            byte[] correct = MorphForm(new[] { type, 23 }, new[] { sample.Bytes.Length, next.Bytes.Length });
            Assert.AreNotSame(combined, FormStreamPadding.Normalize(correct, combined)[1]);
            AssertFallback(MorphForm(new[] { type, 23 }, new[] { sample.Bytes.Length - 1, next.Bytes.Length + 1 }), combined);
            AssertFallback(MorphForm(new[] { type, 23 }, new[] { sample.Bytes.Length + 1, next.Bytes.Length - 1 }), combined);
        }

        /// <summary>Builds independent, explicitly typed scalar/extra/text/column sections from the Microsoft grammar.</summary>
        private static MorphSample MakeMorph(int type, bool rich, bool unicode = false, bool empty = false,
            int extraBit = -1, bool omitStyle = false, int?[] widths = null)
        {
            var sample = new MorphSample();
            var bits = new HashSet<int>(rich ? MorphFieldBits(type) : new[] { 6 });
            if (type == 24) bits.Add(5);
            if (omitStyle) bits.Remove(6);
            if (extraBit >= 0) bits.Add(extraBit);
            if (widths != null) bits.Add(15);
            ulong mask = (1UL << 8) | (1UL << 31);
            foreach (int bit in bits) mask |= 1UL << bit;
            byte[] value = empty ? new byte[0] : unicode ? Encoding.Unicode.GetBytes("Value é") : Encoding.ASCII.GetBytes("Value");
            byte[] caption = empty ? new byte[0] : unicode ? Encoding.Unicode.GetBytes("Caption à") : Encoding.ASCII.GetBytes("Caption");
            byte[] group = empty ? new byte[0] : unicode ? Encoding.Unicode.GetBytes("Group ø") : Encoding.ASCII.GetBytes("Group");
            using (var stream = new MemoryStream())
            using (var writer = new BinaryWriter(stream))
            {
                writer.Write((ushort)0x0200); writer.Write((ushort)0); writer.Write((uint)mask); writer.Write((uint)(mask >> 32));
                int[] sizes = { 4, 4, 4, 4, 1, 1, 1, 1, 0, 2, 4, 2, 2, 2, 2, 2, 1, 1, 1, 0, 1, 1, 4, 4, 4, 4, 4, 2, 2, 2, 0, 0, 4 };
                uint[] values = { 0x2c80481b, 0x112233, 0x445566, 123, 1, type == 24 ? 3u : 2u,
                    (uint)(type - 22), 4, 0, 0x2022, 1800, 2, 0xffff, 3, 5, (uint)(widths?.Length ?? 0),
                    1, 1, 2, 0, 1, 0, Descriptor(value, unicode), Descriptor(caption, unicode), 0x70001,
                    0x334455, 2, 0xffff, 0xffff, 0x03a9, 0, 0, Descriptor(group, unicode) };
                foreach (int bit in Enumerable.Range(0, 33))
                {
                    if (!bits.Contains(bit)) continue;
                    int size = sizes[bit]; if (size == 0) continue;
                    PadMorph(writer, stream, size, sample);
                    sample.Fields[bit] = (int)stream.Length;
                    if (size == 4) writer.Write(values[bit]); else if (size == 2) writer.Write((ushort)values[bit]); else writer.Write((byte)values[bit]);
                }
                PadMorph(writer, stream, 4, sample); sample.SizeOffset = (int)stream.Length;
                writer.Write(0x1020304u); writer.Write(0x5060708u);
                foreach (int bit in new[] { 22, 23, 32 })
                {
                    if (!bits.Contains(bit)) continue;
                    writer.Write(bit == 22 ? value : bit == 23 ? caption : group);
                    PadMorph(writer, stream, 4, sample);
                }
                sample.FontStart = (int)stream.Length;
                byte[] font = RichText();
                // RichText has padding before Weight, at the end of scalar data, and after its font name.
                // These offsets are independently fixed by that existing fixture, not inferred from output.
                foreach (int offset in new[] { 23, 26, 27, 31 }) sample.Padding.Add(sample.FontStart + offset);
                writer.Write(font); sample.ColumnStart = (int)stream.Length;
                if (widths != null)
                foreach (int? width in widths)
                {
                    sample.ColumnOffsets.Add((int)stream.Length);
                    writer.Write((ushort)0x0200); writer.Write((ushort)(width.HasValue ? 8 : 4));
                    writer.Write(width.HasValue ? 1u : 0u); if (width.HasValue) writer.Write(width.Value);
                }
                sample.Bytes = stream.ToArray();
                Array.Copy(BitConverter.GetBytes((ushort)(sample.FontStart - 4)), 0, sample.Bytes, 2, 2);
                return sample;
            }
        }

        private static uint Descriptor(byte[] text, bool unicode) { return (uint)text.Length | (unicode ? 0u : 0x80000000u); }

        private static int[] MorphFieldBits(int type)
        {
            var common = new List<int> { 0, 1, 2, 6, 7, 22, 26 };
            if (type == 23) common.AddRange(new[] { 3, 4, 5, 9, 18, 20, 25 });
            if (type == 24 || type == 25) common.AddRange(new[] { 4, 10, 11, 12, 13, 15, 16, 17, 25 });
            if (type == 24) common.AddRange(new[] { 5, 21 });
            if (type == 25) common.AddRange(new[] { 3, 14, 18, 20 });
            if (type >= 26) common.AddRange(new[] { 21, 23, 24, 29 });
            if (type == 26 || type == 27) common.Add(32);
            return common.ToArray();
        }

        private static void PadMorph(BinaryWriter writer, MemoryStream stream, int alignment, MorphSample sample)
        {
            while (stream.Length % alignment != 0) { sample.Padding.Add((int)stream.Length); writer.Write((byte)0xe7); }
        }

        private static byte[] MorphForm(int[] types, int[] sizes)
        {
            using (var stream = new MemoryStream())
            using (var writer = new BinaryWriter(stream))
            {
                writer.Write(EmptyForm(0x8004, (uint)types.Length, 0), 0, 20);
                var sites = new List<byte>();
                for (int i = 0; i < types.Length; i++) { sites.Add(0); sites.Add(1); }
                while ((sites.Count & 3) != 0) sites.Add(0xe7);
                for (int i = 0; i < types.Length; i++) sites.AddRange(RichSite((uint)types[i], sizes[i]));
                writer.Write((uint)sites.Count); writer.Write(sites.ToArray()); return stream.ToArray();
            }
        }

        private sealed class MorphSample
        {
            internal byte[] Bytes;
            internal int SizeOffset, FontStart, ColumnStart;
            internal readonly Dictionary<int, int> Fields = new Dictionary<int, int>();
            internal readonly HashSet<int> Padding = new HashSet<int>();
            internal readonly List<int> ColumnOffsets = new List<int>();
        }
    }
}
