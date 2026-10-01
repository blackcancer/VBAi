using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

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
            var streams = new Dictionary<string, byte[]> {
                ["/f"] = RichForm(RichLeaf(true).Length, RichLeaf(false).Length),
                ["/o"] = Join(RichLeaf(true), RichLeaf(false))
            };
            var metadata = new Dictionary<string, byte[]> { [""] = ContainerMetadata()[""] };
            Assert.AreNotSame(streams, FormStreamPadding.NormalizeGraph(streams, metadata));
            Assert.IsNull(FormStreamPadding.ReadFontBindings(streams, metadata));
        }
    }
}
