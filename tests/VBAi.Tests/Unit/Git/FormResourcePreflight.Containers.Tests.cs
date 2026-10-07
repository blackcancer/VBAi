using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace VBAi.Tests.Unit
{
    public sealed partial class FormResourcePreflightTests
    {
        [TestMethod]
        public void WholeResourceFontBindingsRetainExactNestedPrecisionWithoutChangingTransportBytes()
        {
            byte[] resources = FormStreamPaddingTests.ContainerResourceBefore();
            byte[] before = (byte[])resources.Clone();
            var bindings = FormResourcePreflight.ReadFontBindings(resources, 0);
            Assert.AreEqual(2, bindings.Length);
            Assert.AreEqual(82700u, BitConverter.ToUInt32(bindings.Single(binding => binding.OwnerPath == "Controls/QualificationExtra").Descriptor, 6));
            CollectionAssert.AreEqual(before, resources);
            byte[] unknown = (byte[])resources.Clone(); unknown[3048] ^= 1;
            Assert.IsNull(FormResourcePreflight.ReadFontBindings(unknown, 0));
            Assert.ThrowsException<InvalidOperationException>(() => FormResourcePreflight.ReadFontBindings(new byte[] { 0 }, 0));
        }

        [TestMethod]
        public void RetainedWholeFrameMultiPageResourcesHaveEqualComparisonAndExactRawTransport()
        {
            byte[] before = FormStreamPaddingTests.ContainerResourceBefore();
            byte[] after = FormStreamPaddingTests.ContainerResourceAfter();
            byte[] savedBefore = (byte[])before.Clone(), savedAfter = (byte[])after.Clone();
            Assert.IsFalse(before.SequenceEqual(after), "This fixture must contain the observed native export drift.");
            byte[] comparison = FormResourcePreflight.ComparisonBytes(before, new[] { 0 });
            CollectionAssert.AreEqual(comparison, FormResourcePreflight.ComparisonBytes(after, new[] { 0 }));
            var streams = ReadComparisonStreams(comparison);
            CollectionAssert.AreEqual(FormStreamPaddingTests.ContainerStreamsBefore()["/i03/i04/x"], streams["/i03/i04/x"]);
            Assert.AreEqual((byte)0, streams["/i03/i04/i06/f"][50]);
            Assert.AreEqual((byte)0, streams["/i03/i04/i06/f"][51]);
            CollectionAssert.AreEqual(savedBefore, before); CollectionAssert.AreEqual(savedAfter, after);
        }

        [TestMethod]
        public void UnknownFrameMetadataPreventsIndependentNormalizationOfItsKnownPageDescendants()
        {
            byte[] before = FormStreamPaddingTests.ContainerResourceBefore(), after = FormStreamPaddingTests.ContainerResourceAfter();
            // Physical offsets from the retained MS-CFB directory chain; verified against exact documented CLSID bytes.
            byte[] frameClass = new Guid("6E182020-F460-11CE-9BCD-00AA00608E01").ToByteArray();
            CollectionAssert.AreEqual(frameClass, before.Skip(3048).Take(16).ToArray());
            CollectionAssert.AreEqual(frameClass, after.Skip(3048).Take(16).ToArray());
            before[3048] ^= 1; after[3048] ^= 1;
            byte[] first = FormResourcePreflight.ComparisonBytes(before, new[] { 0 });
            byte[] repeated = FormResourcePreflight.ComparisonBytes(after, new[] { 0 });
            Assert.IsFalse(first.SequenceEqual(repeated));
            var originalStreams = FormStreamPaddingTests.ContainerStreamsBefore();
            var repeatedStreams = FormStreamPaddingTests.ContainerStreamsAfter();
            var firstStreams = ReadComparisonStreams(first); var otherStreams = ReadComparisonStreams(repeated);
            foreach (string path in originalStreams.Keys)
            {
                CollectionAssert.AreEqual(originalStreams[path], firstStreams[path], "Partial normalization after parent refusal: " + path);
                CollectionAssert.AreEqual(repeatedStreams[path], otherStreams[path]);
            }
            Assert.AreNotEqual(firstStreams["/i03/i04/i06/f"][51], otherStreams["/i03/i04/i06/f"][51]);
        }

        [TestMethod]
        public void EveryContainerClsidAndStateByteRemainsSignificantInResourceComparison()
        {
            byte[] raw = FormStreamPaddingTests.ContainerResourceBefore();
            byte[] expected = FormResourcePreflight.ComparisonBytes(raw, new[] { 0 });
            foreach (int position in new[] { 2664, 3048, 3944, 5480, 5608 })
                for (int i = 0; i < 20; i++)
                {
                    byte[] changed = (byte[])raw.Clone(); changed[position + i] ^= 1;
                    Assert.IsFalse(expected.SequenceEqual(FormResourcePreflight.ComparisonBytes(changed, new[] { 0 })),
                        "Lost storage identity/state byte " + (position + i));
                }
        }

        // Decode the comparison record envelope independently, without asking the production parser to expose its internals.
        private static Dictionary<string, byte[]> ReadComparisonStreams(byte[] bytes)
        {
            var streams = new Dictionary<string, byte[]>(StringComparer.Ordinal);
            using (var reader = new BinaryReader(new MemoryStream(bytes), Encoding.UTF8))
            {
                Assert.AreEqual(0, reader.ReadInt32()); reader.ReadBytes(20);
                int length = reader.ReadInt32(); long end = reader.BaseStream.Position + length;
                Assert.AreEqual(16, reader.ReadBytes(16).Length);
                while (reader.BaseStream.Position < end)
                {
                    byte kind = reader.ReadByte(); string path = reader.ReadString(); reader.ReadString();
                    Assert.AreEqual(20, reader.ReadBytes(20).Length);
                    int size = reader.ReadInt32(); byte[] data = reader.ReadBytes(size); Assert.AreEqual(size, data.Length);
                    if (kind == 2) streams.Add(path, data);
                }
                Assert.AreEqual(end, reader.BaseStream.Position); Assert.AreEqual(0, reader.ReadInt32());
                Assert.AreEqual(reader.BaseStream.Length, reader.BaseStream.Position);
            }
            return streams;
        }
    }
}
