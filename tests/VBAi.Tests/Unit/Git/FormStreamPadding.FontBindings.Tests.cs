using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;

namespace VBAi.Tests.Unit
{
    public sealed partial class FormStreamPaddingTests
    {
        [TestMethod]
        public void FontBindingsPreserveRootAndNestedFrameDescriptorsWithoutAddingImplicitFonts()
        {
            var streams = ContainerStreamsBefore(); var metadata = ContainerMetadata();
            var originals = CopyStreams(streams);
            var bindings = FormStreamPadding.ReadFontBindings(streams, metadata);
            Assert.IsNotNull(bindings); Assert.AreEqual(2, bindings.Length);
            var root = bindings.Single(binding => binding.OwnerPath == "");
            var frame = bindings.Single(binding => binding.OwnerPath == "Controls/QualificationExtra");
            Assert.AreEqual(7u, root.Type); Assert.AreEqual(14u, frame.Type);
            Assert.AreEqual(82500u, BitConverter.ToUInt32(root.Descriptor, 6));
            Assert.AreEqual(82700u, BitConverter.ToUInt32(frame.Descriptor, 6));
            FormFontRestoration.ValidateDescriptor(root.Descriptor); FormFontRestoration.ValidateDescriptor(frame.Descriptor);
            foreach (string key in originals.Keys) CollectionAssert.AreEqual(originals[key], streams[key]);
            byte expected = root.Descriptor[0]; root.Descriptor[0] ^= 1;
            Assert.AreEqual(expected, FormStreamPadding.ReadFontBindings(streams, metadata).Single(binding => binding.OwnerPath == "").Descriptor[0]);
        }

        [TestMethod]
        public void FontOwnerNamesAreBoundedBeforeBuildingPathsWithoutRestrictingComparison()
        {
            foreach (string name in new[] { "Frame", new string('A', 40), "", "9Frame", "Frame/Child", new string('A', 41), new string('A', 1000) })
            {
                var header = new BlockBuilder(0x400, 0x08000040); header.Field(4, 0x8004); header.Field(4, 1);
                var site = new BlockBuilder(0, 0x95);
                site.Field(4, 0x80000000u | (uint)name.Length); site.Field(4, 1); site.Field(4, 0x40023);
                site.Field(2, 14); site.Align(4); site.String(System.Text.Encoding.ASCII.GetBytes(name));
                byte[] child = site.Finish();
                byte[] root = new[] { header.Finish(), BitConverter.GetBytes(1u), BitConverter.GetBytes((uint)(4 + child.Length)),
                    new byte[] { 0, 1, 0, 0 }, child }.SelectMany(bytes => bytes).ToArray();
                // A minimal no-site FormControl is valid for this Frame storage.
                var frameHeader = new BlockBuilder(0x400, 0x08000040);
                frameHeader.Field(4, 0x8004); frameHeader.Field(4, 1);
                byte[] empty = Join(frameHeader.Finish(), new byte[8]);
                var streams = new Dictionary<string, byte[]> { ["/f"] = root, ["/o"] = new byte[0], ["/i01/f"] = empty, ["/i01/o"] = new byte[0] };
                var metadata = new Dictionary<string, byte[]> { [""] = Metadata("C62A69F0-16DC-11CE-9E98-00AA00574A4F"), ["/i01"] = Metadata("6E182020-F460-11CE-9BCD-00AA00608E01") };
                var original = CopyStreams(streams);
                Assert.AreNotSame(streams, FormStreamPadding.NormalizeGraph(streams, metadata), name);
                var bindings = FormStreamPadding.ReadFontBindings(streams, metadata);
                if (name == "Frame" || name.Length == 40) Assert.IsNotNull(bindings, name);
                else Assert.IsNull(bindings, name);
                foreach (string key in streams.Keys) CollectionAssert.AreEqual(original[key], streams[key]);
            }
        }

        [TestMethod]
        public void UnknownDescendantPreventsPartialFontPlansFromItsKnownAncestors()
        {
            var original = ContainerStreamsBefore(); var metadata = ContainerMetadata();
            foreach (string key in original.Keys.Where(key => key.EndsWith("/f") || key.EndsWith("/o") || key.EndsWith("/x")))
            {
                var streams = CopyStreams(original); streams[key] = Join(streams[key], new byte[] { 1 });
                Assert.IsNull(FormStreamPadding.ReadFontBindings(streams, metadata), key);
            }
            var wrongClass = metadata.ToDictionary(pair => pair.Key, pair => (byte[])pair.Value.Clone());
            wrongClass["/i03"][0] ^= 1;
            Assert.IsNull(FormStreamPadding.ReadFontBindings(original, wrongClass));
            Assert.IsNull(FormStreamPadding.ReadFontBindings(null, metadata));
            Assert.IsNull(FormStreamPadding.ReadFontBindings(original, null));
        }

        [TestMethod]
        public void TextPropsFontRemainsComparableButDoesNotProduceAStandardFontRestorationPlan()
        {
            var streams = new Dictionary<string, byte[]>
            {
                ["/f"] = RichForm(RichLeaf(true).Length, RichLeaf(false).Length),
                ["/o"] = Join(RichLeaf(true), RichLeaf(false))
            };
            var metadata = new Dictionary<string, byte[]> { [""] = ContainerMetadata()[""] };
            Assert.AreNotSame(streams, FormStreamPadding.NormalizeGraph(streams, metadata));
            Assert.IsNull(FormStreamPadding.ReadFontBindings(streams, metadata));
        }
    }
}
