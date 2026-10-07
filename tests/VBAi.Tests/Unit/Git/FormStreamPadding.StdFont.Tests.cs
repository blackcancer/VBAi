using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Text;

namespace VBAi.Tests.Unit
{
    public sealed partial class FormStreamPaddingTests
    {
        [TestMethod]
        public void StandardFormFontAllowsSitePaddingNormalizationAndPreservesEveryFontByte()
        {
            byte[] objects = Join(RichLeaf(true), RichLeaf(false));
            byte[] form = StandardFontForm(14, 1000, 655350000, "Tahoma", out int fontStart, out int fontLength);
            byte[] original = (byte[])form.Clone();
            byte[][] normalized = FormStreamPadding.Normalize(form, objects);
            Assert.AreNotSame(form, normalized[0]);
            CollectionAssert.AreEqual(original, form);
            for (int i = 0; i < fontLength; i++) Assert.AreEqual(form[fontStart + i], normalized[0][fontStart + i]);
            for (int i = fontStart; i < fontStart + fontLength; i++)
            {
                byte[] changed = (byte[])form.Clone(); changed[i] ^= 1;
                byte[][] comparison = FormStreamPadding.Normalize(changed, objects);
                Assert.AreNotEqual(normalized[0][i], comparison[0][i], "Font byte " + i + " must remain significant even if malformed.");
            }
            byte[][] repeated = FormStreamPadding.Normalize(normalized[0], normalized[1]);
            CollectionAssert.AreEqual(normalized[0], repeated[0]);
            CollectionAssert.AreEqual(normalized[1], repeated[1]);
        }

        [TestMethod]
        [DataRow(0, 0, 1, "")]
        [DataRow(2, 1, 82500, "Arial")]
        [DataRow(4, 400, 100000, "Tahoma")]
        [DataRow(8, 1000, 655350000, "1234567890123456789012345678901")]
        public void DocumentedStandardFontBoundariesRemainSupported(int flags, int weight, int height, string name)
        {
            byte[] objects = Join(RichLeaf(true), RichLeaf(false));
            byte[] form = StandardFontForm((byte)flags, (ushort)weight, (uint)height, name, out _, out _);
            Assert.AreNotSame(form, FormStreamPadding.Normalize(form, objects)[0]);
        }

        [TestMethod]
        public void MalformedStandardFontsAndAllTruncationsFallBackWithoutPartialNormalization()
        {
            byte[] objects = Join(RichLeaf(true), RichLeaf(false));
            byte[] form = StandardFontForm(0, 400, 82500, "Tahoma", out int fontStart, out int fontLength);
            foreach (int offset in new[] { 0, 16, 19, 20, 22, 26, 27 })
            {
                byte[] changed = (byte[])form.Clone();
                if (offset == 16) changed[fontStart + offset] = 2;
                else if (offset == 19) changed[fontStart + offset] = 1;
                else if (offset == 20) Array.Copy(BitConverter.GetBytes((ushort)1001), 0, changed, fontStart + offset, 2);
                else if (offset == 22) Array.Clear(changed, fontStart + offset, 4);
                else if (offset == 26) changed[fontStart + offset] = 32;
                else if (offset == 27) changed[fontStart + offset] = 128;
                else changed[fontStart + offset] ^= 1;
                AssertFallback(changed, objects);
            }
            foreach (byte flags in new byte[] { 1, 16, 128, 255 })
                AssertFallback(StandardFontForm(flags, 400, 82500, "Tahoma", out _, out _), objects);
            AssertFallback(StandardFontForm(0, 400, 655350001, "Tahoma", out _, out _), objects);
            for (int i = 0; i < form.Length; i++)
            {
                byte[] truncated = new byte[i]; Array.Copy(form, truncated, i);
                AssertFallback(truncated, objects);
            }
        }

        private static byte[] StandardFontForm(byte flags, ushort weight, uint height, string name, out int start, out int length)
        {
            byte[] textForm = RichForm(RichLeaf(true).Length, RichLeaf(false).Length);
            start = 4 + BitConverter.ToUInt16(textForm, 2);
            int suffix = start + 16 + RichText().Length;
            using (var stream = new MemoryStream())
            using (var writer = new BinaryWriter(stream))
            {
                writer.Write(textForm, 0, start);
                writer.Write(new Guid("0BE35203-8F91-11CE-9DE3-00AA004BB851").ToByteArray());
                writer.Write((byte)1); writer.Write((ushort)0xffff); // Signed charset is preserved.
                writer.Write(flags); writer.Write(weight); writer.Write(height);
                byte[] face = Encoding.ASCII.GetBytes(name);
                writer.Write((byte)face.Length); writer.Write(face);
                length = 27 + face.Length;
                writer.Write(textForm, suffix, textForm.Length - suffix);
                return stream.ToArray();
            }
        }
    }
}
