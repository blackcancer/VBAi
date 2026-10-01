namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class FormStreamPaddingTests
    {
        // Logical streams from an owned, synthetic Excel UserForm export. All strings are test data.
        private const string NativeForm =
            "00044000080c080c020000001b00008002000000007d00006b1f0000c6140000" +
            "000000000000000053796e74686574696320476974207175616c696669636174" +
            "696f6e00000002000000700000000082016f00003400f5010000120000800100" +
            "0000320000004c000000000015005175616c696669636174696f6e4c6162656c" +
            "0000a7010000a701000000003000e50100001300008002000000400000000100" +
            "11005175616c696669636174696f6e427574746f6e00a7010000f6040000";
        private const string NativeObjects =
            "00022c00280000001a0000804f726967696e616c2073796e7468657469632063" +
            "617074696f6e6e00d81300007b020000000218003500000006000080a5000000" +
            "000200005461686f6d61000000022000280000001000008053796e7468657469" +
            "6320627574746f6eec0900004f030000000218007500000006000080a5000000" +
            "000203005461686f6d617469";

        internal static byte[][] NativeStreams() { return new[] { Hex(NativeForm), Hex(NativeObjects) }; }

        [TestMethod]
        public void NativePaddingDifferencesNormalizeWithoutMutatingInputs()
        {
            byte[] form = Hex(NativeForm), objects = Hex(NativeObjects);
            byte[] originalForm = (byte[])form.Clone(), originalObjects = (byte[])objects.Clone();
            byte[][] expected = FormStreamPadding.Normalize(form, objects);
            Assert.AreNotSame(form, expected[0]); Assert.AreNotSame(objects, expected[1]);
            foreach (int i in FormPadding()) form[i] ^= 0xa5;
            foreach (int i in ObjectPadding()) objects[i] ^= 0xa5;
            byte[][] result = FormStreamPadding.Normalize(form, objects);
            CollectionAssert.AreEqual(expected[0], result[0]); CollectionAssert.AreEqual(expected[1], result[1]);
            foreach (int i in FormPadding()) originalForm[i] ^= 0xa5;
            foreach (int i in ObjectPadding()) originalObjects[i] ^= 0xa5;
            CollectionAssert.AreEqual(originalForm, form); CollectionAssert.AreEqual(originalObjects, objects);
            byte[][] repeated = FormStreamPadding.Normalize(result[0], result[1]);
            CollectionAssert.AreEqual(result[0], repeated[0]); CollectionAssert.AreEqual(result[1], repeated[1]);
        }

        [TestMethod]
        public void EveryNativeNonPaddingByteRemainsSignificant()
        {
            byte[] form = Hex(NativeForm), objects = Hex(NativeObjects);
            byte[][] original = FormStreamPadding.Normalize(form, objects);
            for (int stream = 0; stream < 2; stream++)
            {
                var padding = new HashSet<int>(stream == 0 ? FormPadding() : ObjectPadding());
                byte[] source = stream == 0 ? form : objects;
                for (int i = 0; i < source.Length; i++)
                {
                    if (padding.Contains(i)) continue;
                    byte[] altered = (byte[])source.Clone(); altered[i] ^= 1;
                    byte[][] normalized = FormStreamPadding.Normalize(stream == 0 ? altered : form, stream == 1 ? altered : objects);
                    Assert.AreNotEqual(original[stream][i], normalized[stream][i], "Stream " + stream + ", byte " + i);
                }
            }
        }

        [TestMethod]
        public void TruncationAndUnknownChildrenReturnBothOriginalStreams()
        {
            byte[] form = Hex(NativeForm), objects = Hex(NativeObjects);
            for (int i = 0; i < form.Length; i++)
            {
                byte[] truncated = new byte[i]; Array.Copy(form, truncated, i);
                AssertFallback(truncated, objects);
            }
            for (int i = 0; i < objects.Length; i++)
            {
                byte[] truncated = new byte[i]; Array.Copy(objects, truncated, i);
                AssertFallback(form, truncated);
            }
            foreach (int offset in new[] { 108, 160 })
            {
                byte[] unknown = (byte[])form.Clone(); unknown[offset] = 23;
                AssertFallback(unknown, objects);
            }
            byte[] extended = new byte[objects.Length + 1]; Array.Copy(objects, extended, objects.Length);
            AssertFallback(form, extended);
            AssertFallback(null, objects); AssertFallback(form, null);
        }

        [TestMethod]
        public void UnsupportedFormFontPicturesClassTableDepthAndUnknownMasksAreConservative()
        {
            byte[] form = Hex(NativeForm), objects = Hex(NativeObjects);
            foreach (int bit in new[] { 0, 4, 5, 14, 15, 20, 21, 28, 31 })
            {
                byte[] changed = (byte[])form.Clone(); Write32(changed, 4, BitConverter.ToUInt32(changed, 4) | (1u << bit));
                AssertFallback(changed, objects);
            }
            foreach (int offset in new[] { 68, 78, 79, 80, 84, 88 })
            {
                byte[] changed = (byte[])form.Clone(); changed[offset] = 0xff;
                AssertFallback(changed, objects);
            }
            foreach (int maskOffset in new[] { 4, 80 })
            {
                byte[] changed = (byte[])objects.Clone(); Write32(changed, maskOffset, BitConverter.ToUInt32(changed, maskOffset) | (1u << 10));
                AssertFallback(form, changed);
            }
        }

        [TestMethod]
        public void EmptyFormsAndCompressedOrUnicodeStringsUseDocumentedLengths()
        {
            byte[] empty = EmptyForm(4, 0, 0);
            Assert.AreNotSame(empty, FormStreamPadding.Normalize(empty, new byte[0])[0]);
            byte[] noClassTable = EmptyForm(0x8004, 0, 0);
            Assert.AreNotSame(noClassTable, FormStreamPadding.Normalize(noClassTable, new byte[0])[0]);
            AssertFallback(EmptyForm(0x4004, 0, 0), new byte[0]);
            AssertFallback(EmptyForm(4, uint.MaxValue, 0), new byte[0]);
            AssertFallback(EmptyForm(4, 0, uint.MaxValue), new byte[0]);
            byte[] form = Hex(NativeForm), objects = Hex(NativeObjects);
            byte[] unicode = (byte[])objects.Clone(); Write32(unicode, 8, 26); // Same byte count, Unicode encoding.
            Assert.AreNotSame(unicode, FormStreamPadding.Normalize(form, unicode)[1]);
            Write32(unicode, 8, 25); AssertFallback(form, unicode); // Odd UTF-16 length.
            Write32(unicode, 8, uint.MaxValue); AssertFallback(form, unicode);
        }

        [TestMethod]
        public void ScalarAlignmentTextFontAndAllSupportedSiteStringsPreserveValues()
        {
            byte[] label = RichLeaf(true), button = RichLeaf(false);
            byte[] objects = Join(label, button);
            byte[] form = RichForm(label.Length, button.Length);
            byte[][] normalized = FormStreamPadding.Normalize(form, objects);
            Assert.AreNotSame(form, normalized[0]); Assert.AreNotSame(objects, normalized[1]);
            Assert.IsTrue(Array.Exists(normalized[0], value => value == 0x65));
            // Fields, names, coordinates and font values survive; only sentinel padding is cleared.
            for (int stream = 0; stream < 2; stream++)
            {
                byte[] input = stream == 0 ? form : objects;
                for (int i = 0; i < input.Length; i++)
                    if (input[i] != normalized[stream][i])
                    {
                        Assert.AreEqual((byte)0xe7, input[i]); Assert.AreEqual((byte)0, normalized[stream][i]);
                    }
            }
            byte[] unsupportedFont = (byte[])form.Clone();
            int guidStart = 4 + BitConverter.ToUInt16(form, 2); unsupportedFont[guidStart] ^= 1;
            AssertFallback(unsupportedFont, objects);
            byte[] formExtension = Join(form, new byte[] { 1 }); AssertFallback(formExtension, objects);
            byte[] singleDepthEntries = (byte[])form.Clone();
            // RichForm uses individual entries, while the native fixture uses a run-length entry.
            byte[][] second = FormStreamPadding.Normalize(singleDepthEntries, objects);
            CollectionAssert.AreEqual(normalized[0], second[0]);
        }

        private static byte[] RichForm(int labelSize, int buttonSize)
        {
            var header = new BlockBuilder(0x400, 0x0fdf3fceu);
            header.Field(4, 0x12345678); header.Field(4, 0x87654321); header.Field(4, 3);
            header.Field(4, 0x8004); // Class table omitted.
            header.Field(1, 1); header.Field(1, 2); header.Field(1, 3);
            header.Field(4, 5); header.Field(1, 1); header.Field(1, 2);
            header.Field(4, 0x7f010203); header.Field(4, 0x80000003); header.Field(2, 0xffff);
            header.Field(4, 120); header.Field(1, 2); header.Field(1, 1);
            header.Field(4, 55); header.Field(4, 1000); header.Align(4);
            header.Bytes(new byte[24]); header.String(new byte[] { 0x65, 0x66, 0x67 });
            using (var stream = new MemoryStream())
            using (var writer = new BinaryWriter(stream))
            {
                writer.Write(header.Finish());
                writer.Write(new Guid("AFC20920-DA4E-11CE-B943-00AA006887B4").ToByteArray());
                writer.Write(RichText()); writer.Write(2u);
                byte[] first = RichSite(21, labelSize), second = RichSite(17, buttonSize);
                writer.Write((uint)(4 + first.Length + second.Length));
                writer.Write(new byte[] { 0, 1, 0, 1 }); writer.Write(first); writer.Write(second);
                return stream.ToArray();
            }
        }
        private static byte[] RichSite(uint type, int size)
        {
            var block = new BlockBuilder(0, 0x7bff);
            block.Field(4, 0x80000003); block.Field(4, 2); block.Field(4, type); block.Field(4, 100);
            block.Field(4, 0x33); block.Field(4, (uint)size); block.Field(2, 2); block.Field(2, type);
            block.Field(2, 7);
            for (int i = 0; i < 4; i++) block.Field(4, 0x80000001);
            block.Align(4);
            block.String(new byte[] { 65, 66, 67 }); block.String(new byte[] { 65, 0 });
            block.Bytes(new byte[] { 1, 2, 3, 4, 5, 6, 7, 8 });
            for (int i = 0; i < 4; i++) block.String(new byte[] { (byte)(70 + i) });
            return block.Finish();
        }
        private static byte[] RichLeaf(bool label)
        {
            var block = new BlockBuilder(0x200, label ? 0xbffu : 0x37fu);
            block.Field(4, 0x112233); block.Field(4, 0x445566); block.Field(4, 0x1234);
            block.Field(4, 0x80000003); block.Field(4, 2); block.Field(1, 3);
            if (label) { block.Field(4, 0x778899); block.Field(2, 1); block.Field(2, 2); }
            block.Field(2, 65); block.Align(4); block.String(new byte[] { 65, 66, 67 });
            block.Bytes(new byte[] { 1, 0, 0, 0, 2, 0, 0, 0 });
            return Join(block.Finish(), RichText());
        }
        private static byte[] RichText()
        {
            var block = new BlockBuilder(0x200, 0xf7);
            block.Field(4, 0x80000003); block.Field(4, 5); block.Field(4, 220);
            block.Field(1, 1); block.Field(1, 2); block.Field(1, 3); block.Field(2, 600);
            block.Align(4); block.String(new byte[] { 88, 89, 90 });
            return block.Finish();
        }
        private static byte[] Join(byte[] first, byte[] second)
        {
            var result = new byte[first.Length + second.Length];
            Array.Copy(first, result, first.Length); Array.Copy(second, 0, result, first.Length, second.Length);
            return result;
        }
        private sealed class BlockBuilder
        {
            private readonly MemoryStream stream = new MemoryStream();
            private readonly BinaryWriter writer;
            internal BlockBuilder(ushort version, uint mask)
            {
                writer = new BinaryWriter(stream); writer.Write(version); writer.Write((ushort)0); writer.Write(mask);
            }
            internal void Align(int size) { while ((stream.Length % size) != 0) writer.Write((byte)0xe7); }
            internal void Field(int size, uint value)
            {
                Align(size);
                if (size == 4) writer.Write(value);
                else if (size == 2) writer.Write((ushort)value);
                else writer.Write((byte)value);
            }
            internal void Bytes(byte[] bytes) { writer.Write(bytes); }
            internal void String(byte[] bytes) { writer.Write(bytes); Align(4); }
            internal byte[] Finish()
            {
                byte[] result = stream.ToArray(); Array.Copy(BitConverter.GetBytes((ushort)(result.Length - 4)), 0, result, 2, 2);
                writer.Dispose(); stream.Dispose(); return result;
            }
        }

        private static byte[] EmptyForm(uint flags, uint count, uint length)
        {
            using (var stream = new MemoryStream())
            using (var writer = new BinaryWriter(stream))
            {
                writer.Write((ushort)0x400); writer.Write((ushort)12); writer.Write(0x08000040u);
                writer.Write(flags); writer.Write(0u);
                if ((flags & 0x8000) == 0) writer.Write((ushort)0);
                writer.Write(count); writer.Write(length);
                return stream.ToArray();
            }
        }
        private static void AssertFallback(byte[] form, byte[] objects)
        {
            byte[][] result = FormStreamPadding.Normalize(form, objects);
            Assert.AreSame(form, result[0]); Assert.AreSame(objects, result[1]);
        }
        private static int[] FormPadding() { return new[] { 67, 81, 128, 129, 181 }; }
        private static int[] ObjectPadding() { return new[] { 38, 39, 66, 67, 74, 75, 131, 138, 139 }; }
        private static byte[] Hex(string value)
        {
            var bytes = new byte[value.Length / 2];
            for (int i = 0; i < bytes.Length; i++) bytes[i] = Convert.ToByte(value.Substring(i * 2, 2), 16);
            return bytes;
        }
        private static void Write32(byte[] bytes, int offset, uint value) { Array.Copy(BitConverter.GetBytes(value), 0, bytes, offset, 4); }
    }
}
