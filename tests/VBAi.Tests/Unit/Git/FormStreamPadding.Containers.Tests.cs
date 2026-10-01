using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    public sealed partial class FormStreamPaddingTests
    {
        [TestMethod]
        public void RetainedCompleteContainerGraphNormalizesOnlyProvenPaddingWithoutChangingInputs()
        {
            var before = ContainerStreamsBefore(); var after = ContainerStreamsAfter(); var metadata = ContainerMetadata();
            var originals = before.ToDictionary(x => x.Key, x => (byte[])x.Value.Clone());
            var comparison = FormStreamPadding.NormalizeGraph(before, metadata);
            var other = FormStreamPadding.NormalizeGraph(after, metadata);
            Assert.AreNotSame(before, comparison); Assert.AreNotSame(after, other);
            var padding = ContainerPadding();
            foreach (string path in before.Keys)
            {
                CollectionAssert.AreEqual(originals[path], before[path], "Input mutation: " + path);
                CollectionAssert.AreEqual(comparison[path], other[path], "Repeat drift: " + path);
                if (!padding.TryGetValue(path, out var spans)) spans = new HashSet<int>();
                for (int i = 0; i < before[path].Length; i++)
                    Assert.AreEqual(spans.Contains(i) ? (byte)0 : before[path][i], comparison[path][i], path + ":" + i);
                if (padding.ContainsKey(path)) Assert.AreNotSame(before[path], comparison[path]);
                else Assert.AreSame(before[path], comparison[path], "Opaque/x streams must remain exact and retained.");
                foreach (int offset in spans) before[path][offset] ^= 0xa5;
            }
            var perturbed = FormStreamPadding.NormalizeGraph(before, metadata);
            var repeated = FormStreamPadding.NormalizeGraph(comparison, metadata);
            foreach (string path in before.Keys)
            {
                CollectionAssert.AreEqual(comparison[path], perturbed[path]);
                CollectionAssert.AreEqual(comparison[path], repeated[path]);
            }
        }

        [TestMethod]
        public void EveryContainerStreamNonPaddingByteIncludingIgnoredPageRecordRemainsSignificant()
        {
            var input = ContainerStreamsBefore(); var metadata = ContainerMetadata(); var padding = ContainerPadding();
            var expected = FormStreamPadding.NormalizeGraph(input, metadata);
            foreach (string path in input.Keys)
            {
                if (!padding.TryGetValue(path, out var spans)) spans = new HashSet<int>();
                for (int i = 0; i < input[path].Length; i++)
                {
                    if (spans.Contains(i)) continue;
                    var changed = CopyStreams(input); changed[path][i] ^= 1;
                    var actual = FormStreamPadding.NormalizeGraph(changed, metadata);
                    Assert.AreNotEqual(expected[path][i], actual[path][i], "Lost property/metadata byte " + path + ":" + i);
                }
            }
        }

        [TestMethod]
        public void EveryContainerFormObjectAndXTruncationOrExtensionRefusesWholeGraph()
        {
            var input = ContainerStreamsBefore(); var metadata = ContainerMetadata();
            foreach (string path in input.Keys.Where(x => x.EndsWith("/f") || x.EndsWith("/o") || x.EndsWith("/x")))
            {
                for (int length = 0; length < input[path].Length; length++)
                {
                    var changed = CopyStreams(input); changed[path] = Slice(changed[path], 0, length); AssertGraphFallback(changed, metadata);
                }
                var extra = CopyStreams(input); extra[path] = Join(extra[path], new byte[] { 0x71 }); AssertGraphFallback(extra, metadata);
            }
        }

        [TestMethod]
        public void MissingExtraUnknownClassAndIncorrectContainerMetadataKeepOriginalDictionaryAndArrays()
        {
            var original = ContainerStreamsBefore(); var metadata = ContainerMetadata();
            foreach (string path in original.Keys.Where(x => x.EndsWith("/f") || x.EndsWith("/o") || x.EndsWith("/x")))
            {
                var missing = CopyStreams(original); missing.Remove(path); AssertGraphFallback(missing, metadata);
            }
            foreach (string path in metadata.Keys)
            {
                var missing = CopyStreams(metadata); missing.Remove(path); AssertGraphFallback(original, missing);
                var bad = CopyStreams(metadata); bad[path][0] ^= 1; AssertGraphFallback(original, bad);
                bad = CopyStreams(metadata); bad[path] = Slice(bad[path], 0, 19); AssertGraphFallback(original, bad);
                bad = CopyStreams(metadata); bad[path] = null; AssertGraphFallback(original, bad);
            }
            foreach (string path in new[] { "/unknown", "/i03/x", "/i03/i04/i06/x", "/i03/i04/i08/f" })
            {
                var extra = CopyStreams(original); extra[path] = new byte[0]; AssertGraphFallback(extra, metadata);
            }
            var extraStorage = CopyStreams(metadata); extraStorage["/i99"] = (byte[])metadata[""] .Clone(); AssertGraphFallback(original, extraStorage);
            var nullStream = CopyStreams(original); nullStream["/f"] = null; AssertGraphFallback(nullStream, metadata);
            Assert.AreSame(original, FormStreamPadding.NormalizeGraph(original, null));
            Assert.AreSame(original, FormStreamPadding.NormalizeGraph(original, metadata, null));
            Assert.AreSame(original, FormStreamPadding.NormalizeGraph(original, metadata, "/wrong"));
            Assert.IsNull(FormStreamPadding.NormalizeGraph(null, metadata));
        }

        [TestMethod]
        public void MalformedParentSitesAndUnresolvedIDsNeverConsumeAZeroSizedLeaf()
        {
            var input = ContainerStreamsBefore(); var metadata = ContainerMetadata();
            // Root third parent site starts223: Name/Tag lengths, ID239, Flags243, cached type249.
            Assert.AreEqual(3u, BitConverter.ToUInt32(input["/f"], 239));
            foreach (Action<byte[]> corrupt in new Action<byte[]>[]
            {
                b => Write32(b, 239, 99), // Exact i99 storage does not exist.
                b => Write32(b, 239, 1), // Duplicate ID of Label; cannot alias an existing object.
                b => Write32(b, 239, uint.MaxValue), // Signed negative IDs are unsupported conservatively.
                b => Write32(b, 243, 0x40033), // Parent falsely marked streamed.
                b => Write32(b, 243, 0x23), // Parent without required PromoteControls.
                b => { b[249] = 7; }, // Page appears outside a MultiPage.
                b => { b[249] = 99; },
                b => Write32(b, 227, BitConverter.ToUInt32(b, 227) | 0x20), // ObjectStreamSize cannot be invented for storage parent.
            })
            {
                var changed = CopyStreams(input); corrupt(changed["/f"]); AssertGraphFallback(changed, metadata);
            }
            byte[][] flat = FormStreamPadding.Normalize(input["/f"], input["/o"]);
            Assert.AreSame(input["/f"], flat[0]); Assert.AreSame(input["/o"], flat[1]);
            // A class-table extension or nonzero depth remains unknown, including anywhere in the graph.
            var classTable = CopyStreams(input); classTable["/f"][101] = 1; AssertGraphFallback(classTable, metadata);
            var depth = CopyStreams(input); depth["/i03/i04/i06/f"][48] = 1; AssertGraphFallback(depth, metadata);
        }

        [TestMethod]
        public void MultiPageXCountsPageIDsNamesMasksAndVersionsAreLinkedToTheExactChildren()
        {
            var input = ContainerStreamsBefore(); var metadata = ContainerMetadata(); string key = "/i03/i04/x";
            CollectionAssert.AreEqual(new byte[] { 6, 0, 0, 0, 7, 0, 0, 0 }, Slice(input[key], 40, 8));
            foreach (Action<byte[]> corrupt in new Action<byte[]>[]
            {
                b => b[0] = 1, b => b[24] = 1,
                b => Write32(b, 4, 1), // First PageProperties is semantically ignored, but unknown grammar is not normalized.
                b => Write32(b, 28, 1), b => Write32(b, 32, 1), b => Write32(b, 32, uint.MaxValue),
                b => Write32(b, 40, 7), // Duplicate last PageID.
                b => Write32(b, 40, 8), // No such Page storage in this MultiPage.
            })
            {
                var changed = CopyStreams(input); corrupt(changed[key]); AssertGraphFallback(changed, metadata);
            }
            // Optional x.ID is retained, with no unsupported assertion that it equals outer siteID4.
            var alternativeId = CopyStreams(input); Write32(alternativeId[key], 36, 42);
            var accepted = FormStreamPadding.NormalizeGraph(alternativeId, metadata);
            Assert.AreNotSame(alternativeId, accepted); Assert.AreSame(alternativeId[key], accepted[key]);
            var flags = CopyStreams(input); Write32(flags[key], 28, 14); // Flags bit does not add a scalar.
            Assert.AreNotSame(flags, FormStreamPadding.NormalizeGraph(flags, metadata));
            // TabNames (Tab3/Tab4) are independent from Page site names (Page1/Page2).
            var changedName = CopyStreams(input); changedName["/i03/i04/f"][104] ^= 1;
            var changedResult = FormStreamPadding.NormalizeGraph(changedName, metadata);
            Assert.AreNotSame(changedName, changedResult); Assert.AreNotEqual(input["/i03/i04/f"][104], changedResult["/i03/i04/f"][104]);
            var reordered = CopyStreams(input); Write32(reordered[key], 40, 7); Write32(reordered[key], 44, 6);
            var reorderedResult = FormStreamPadding.NormalizeGraph(reordered, metadata);
            Assert.AreNotSame(reordered, reorderedResult); Assert.AreSame(reordered[key], reorderedResult[key]);
            Assert.AreNotEqual(input[key][40], reorderedResult[key][40]);
            // A second streamed TabStrip cannot substitute for a Page storage.
            var duplicateTab = CopyStreams(input); duplicateTab["/i03/i04/f"][102] = 18; AssertGraphFallback(duplicateTab, metadata);
        }

        [TestMethod]
        public void AllDocumentedPageAndMultiPagePropertyMasksRetainEveryXByte()
        {
            foreach (uint pageMask in new uint[] { 0, 2, 4, 6 })
                foreach (uint multiMask in new uint[] { 2, 6, 10, 14 })
                {
                    var input = ContainerStreamsBefore(); var metadata = ContainerMetadata(); string key = "/i03/i04/x";
                    byte[] pages = new byte[0];
                    for (int i = 0; i < 3; i++)
                    {
                        var page = new BlockBuilder(0x200, pageMask);
                        if ((pageMask & 2) != 0) page.Field(4, 0);
                        if ((pageMask & 4) != 0) page.Field(4, (uint)(100 + i));
                        pages = Join(pages, page.Finish());
                    }
                    var multi = new BlockBuilder(0x200, multiMask); multi.Field(4, 2);
                    if ((multiMask & 4) != 0) multi.Field(4, 321);
                    input[key] = Join(Join(pages, multi.Finish()), new byte[] { 6, 0, 0, 0, 7, 0, 0, 0 });
                    var result = FormStreamPadding.NormalizeGraph(input, metadata);
                    Assert.AreNotSame(input, result, pageMask + ":" + multiMask);
                    Assert.AreSame(input[key], result[key], "No x bytes, including the ignored first PageProperties, may be normalized.");
                    var unknown = CopyStreams(input); Write32(unknown[key], pages.Length + 4, multiMask | 1);
                    AssertGraphFallback(unknown, metadata);
                }
        }

        [TestMethod]
        public void GraphRootPrefixAndStorageStateAreRetainedWithoutLooseningIdentity()
        {
            var input = ContainerStreamsBefore(); var metadata = ContainerMetadata();
            var prefixedStreams = input.ToDictionary(x => "/owned" + x.Key, x => x.Value);
            var prefixedMetadata = metadata.ToDictionary(x => "/owned" + x.Key, x => x.Value);
            Assert.AreNotSame(prefixedStreams, FormStreamPadding.NormalizeGraph(prefixedStreams, prefixedMetadata, "/owned"));
            foreach (string path in metadata.Keys)
                for (int i = 16; i < 20; i++)
                {
                    var changed = CopyStreams(metadata); changed[path][i] ^= 1;
                    var result = FormStreamPadding.NormalizeGraph(input, changed);
                    Assert.AreNotSame(input, result, "Known CLSID state bits are significant metadata, not a new control grammar.");
                    CollectionAssert.AreEqual(input["/i03/i04/x"], result["/i03/i04/x"]);
                    CollectionAssert.AreEqual(metadata[path], ContainerMetadata()[path]);
                }
        }

        [TestMethod]
        public void CompressedUnicodeAndStrictWideNamesRetainBytesAndRejectMalformedEncoding()
        {
            var input = ContainerStreamsBefore(); var metadata = ContainerMetadata();
            var compressed = CopyStreams(input); compressed["/i03/i04/o"][92] = 0x80;
            var result = FormStreamPadding.NormalizeGraph(compressed, metadata);
            Assert.AreNotSame(compressed, result); Assert.AreEqual((byte)0x80, result["/i03/i04/o"][92]);
            // Form site strings use byte counts, while TabNames array strings use character counts.
            var wide = CopyStreams(input); Write32(wide["/f"], 231, 18);
            for (int i = 0; i < 9; i++) { wide["/f"][251 + 2 * i] = (byte)(65 + i); wide["/f"][252 + 2 * i] = 0; }
            Assert.AreNotSame(wide, FormStreamPadding.NormalizeGraph(wide, metadata));
            var odd = CopyStreams(input); Write32(odd["/f"], 231, 17); AssertGraphFallback(odd, metadata);
            wide["/f"][251] = 0; wide["/f"][252] = 0xd8; AssertGraphFallback(wide, metadata);
            var wideTab = CopyStreams(input); Write32(wideTab["/i03/i04/o"], 88, 2);
            wideTab["/i03/i04/o"][92] = 65; wideTab["/i03/i04/o"][93] = 0;
            wideTab["/i03/i04/o"][94] = 66; wideTab["/i03/i04/o"][95] = 0;
            Assert.AreNotSame(wideTab, FormStreamPadding.NormalizeGraph(wideTab, metadata));
            wideTab["/i03/i04/o"][92] = 0; wideTab["/i03/i04/o"][93] = 0xd8; AssertGraphFallback(wideTab, metadata);
        }

        [TestMethod]
        public void EveryDesignExtenderMaskReadsExactPropertiesAndOnlyClearsFinalAlignment()
        {
            for (uint mask = 0; mask < 32; mask++)
            {
                var b = new BlockBuilder(0x200, mask);
                if ((mask & 1) != 0) b.Field(4, 0x3ffff);
                if ((mask & 2) != 0) b.Field(4, 123);
                if ((mask & 4) != 0) b.Field(4, 321);
                if ((mask & 8) != 0) b.Field(1, 0xff);
                if ((mask & 16) != 0) b.Field(1, 0xfe);
                b.Align(4); byte[] design = b.Finish(); byte[] form = Join(EmptyForm(0x4004, 0, 0), design);
                var streams = new Dictionary<string, byte[]> { ["/f"] = form, ["/o"] = new byte[0] };
                var metadata = new Dictionary<string, byte[]> { [""] = Metadata("C62A69F0-16DC-11CE-9E98-00AA00574A4F") };
                var result = FormStreamPadding.NormalizeGraph(streams, metadata); Assert.AreNotSame(streams, result, "Mask " + mask);
                for (int i = 0; i < form.Length; i++)
                    if (form[i] != result["/f"][i]) { Assert.AreEqual((byte)0xe7, form[i]); Assert.AreEqual((byte)0, result["/f"][i]); }
                foreach (int i in Enumerable.Range(0, design.Length))
                {
                    int at = form.Length - design.Length + i;
                    if (form[at] != result["/f"][at]) continue;
                    var changed = CopyStreams(streams); changed["/f"][at] ^= 1;
                    Assert.AreNotEqual(result["/f"][at], FormStreamPadding.NormalizeGraph(changed, metadata)["/f"][at]);
                }
            }
        }

        [TestMethod]
        public void UnknownDesignExtenderFlagsClickModesAndExtraBytesRefuseEntireGraph()
        {
            var input = ContainerStreamsBefore(); var metadata = ContainerMetadata(); string path = "/i03/i04/f";
            Assert.AreEqual(0x19u, BitConverter.ToUInt32(input[path], 164));
            foreach (Action<byte[]> corrupt in new Action<byte[]>[]
            {
                b => b[160] = 1, b => Write32(b, 164, 32), b => Write32(b, 168, 0x40000),
                b => b[172] = 2, b => b[173] = 0xff,
            })
            {
                var changed = CopyStreams(input); corrupt(changed[path]); AssertGraphFallback(changed, metadata);
            }
        }

        [TestMethod]
        [DataRow(64, true)] [DataRow(65, false)]
        public void DeepParentGraphsUseAnExplicitBoundAndCannotExhaustCallStack(int depth, bool accepted)
        {
            var streams = new Dictionary<string, byte[]>(StringComparer.Ordinal);
            var metadata = new Dictionary<string, byte[]>(StringComparer.Ordinal); string path = "";
            for (int i = 0; i <= depth; i++)
            {
                streams[path + "/f"] = i == depth ? EmptyForm(0x8004, 0, 0) : OneFrameParent();
                streams[path + "/o"] = new byte[0];
                metadata[path] = Metadata(i == 0 ? "C62A69F0-16DC-11CE-9E98-00AA00574A4F" : "6E182020-F460-11CE-9BCD-00AA00608E01");
                path += "/i01";
            }
            var result = FormStreamPadding.NormalizeGraph(streams, metadata);
            if (accepted) Assert.AreNotSame(streams, result); else AssertGraphFallback(streams, metadata);
        }

        [TestMethod]
        public void ContainerGraphBudgetAndUnrelatedStorageTreesAreConservative()
        {
            var streams = ContainerStreamsBefore(); var metadata = ContainerMetadata();
            var largeMetadata = CopyStreams(metadata);
            for (int i = 0; i < 4096; i++) largeMetadata["/unknown" + i] = Metadata("C62A69F0-16DC-11CE-9E98-00AA00574A4F");
            AssertGraphFallback(streams, largeMetadata);
            var largeStreams = new Dictionary<string, byte[]>(streams);
            for (int i = 0; i < 16384; i++) largeStreams["/unknown" + i] = new byte[0];
            AssertGraphFallback(largeStreams, metadata);
            var renamed = CopyStreams(metadata); renamed["/i3"] = renamed["/i03"]; renamed.Remove("/i03");
            AssertGraphFallback(streams, renamed); // Required decimal leading zero is part of exact storage identity.
        }

        [TestMethod]
        [DataRow(0u, "i00")] [DataRow(9u, "i09")] [DataRow(10u, "i10")]
        [DataRow(11u, "i11")] [DataRow(16u, "i16")] [DataRow(99u, "i99")] [DataRow(2147483647u, "i2147483647")]
        public void ParentStorageNamesUseDecimalIDsWithExactlyOneLeadingZeroBelowTen(uint id, string child)
        {
            var streams = new Dictionary<string, byte[]> { ["/f"] = OneFrameParent(id), ["/o"] = new byte[0],
                ["/" + child + "/f"] = EmptyForm(0x8004, 0, 0), ["/" + child + "/o"] = new byte[0] };
            var metadata = new Dictionary<string, byte[]> { [""] = Metadata("C62A69F0-16DC-11CE-9E98-00AA00574A4F"),
                ["/" + child] = Metadata("6E182020-F460-11CE-9BCD-00AA00608E01") };
            Assert.AreNotSame(streams, FormStreamPadding.NormalizeGraph(streams, metadata));
            foreach (string wrong in new[] { "i" + id, "i0" + id, "i" + id.ToString("X2"), "i" + id + "0" }.Distinct().Where(x => x != child))
            {
                var wrongStreams = streams.ToDictionary(x => x.Key.Replace("/" + child + "/", "/" + wrong + "/"), x => x.Value);
                var wrongMetadata = new Dictionary<string, byte[]> { [""] = metadata[""], ["/" + wrong] = metadata["/" + child] };
                AssertGraphFallback(wrongStreams, wrongMetadata);
            }
        }

        [TestMethod]
        public void WholeGraphStreamBytesAndAggregateSiteCountsHaveIndependentBounds()
        {
            var metadata = ContainerMetadata();
            var oversize = ContainerStreamsBefore(); oversize["/o"] = new byte[32 * 1024 * 1024 + 1];
            Assert.AreSame(oversize, FormStreamPadding.NormalizeGraph(oversize, metadata));
            var aggregate = ContainerStreamsBefore(); aggregate["/f"] = new byte[16 * 1024 * 1024 + 1];
            aggregate["/o"] = new byte[16 * 1024 * 1024 + 1];
            Assert.AreSame(aggregate, FormStreamPadding.NormalizeGraph(aggregate, metadata));
            byte[] rootObjects, childObjects;
            var many = new Dictionary<string, byte[]> { ["/f"] = ManyLeavesForm(8192, true, out rootObjects),
                ["/i01/f"] = ManyLeavesForm(8192, false, out childObjects) };
            many["/o"] = rootObjects; many["/i01/o"] = childObjects;
            var known = new Dictionary<string, byte[]> { [""] = Metadata("C62A69F0-16DC-11CE-9E98-00AA00574A4F"),
                ["/i01"] = Metadata("6E182020-F460-11CE-9BCD-00AA00608E01") };
            AssertGraphFallback(many, known); // 8193 + 8192 exceeds the whole graph limit although each form fits.
            many["/i01/f"] = ManyLeavesForm(8191, false, out childObjects); many["/i01/o"] = childObjects;
            Assert.AreNotSame(many, FormStreamPadding.NormalizeGraph(many, known),
                "Exactly 16384 valid sites must be accepted, proving the refusal above is the aggregate boundary.");
        }

        private static byte[] ManyLeavesForm(int count, bool parent, out byte[] objects)
        {
            byte[] leaf = RichLeaf(true); int sites = count + (parent ? 1 : 0);
            var site = new BlockBuilder(0, 0xa4); site.Field(4, 2); site.Field(4, (uint)leaf.Length); site.Field(2, 21); site.Align(4);
            byte[] template = site.Finish();
            using (var f = new System.IO.MemoryStream()) using (var writer = new System.IO.BinaryWriter(f))
            using (var o = new System.IO.MemoryStream())
            {
                byte[] depth = new byte[(sites * 2 + 3) & ~3];
                for (int i = 0; i < sites; i++) depth[2 * i + 1] = 1;
                var header = new BlockBuilder(0x400, 0x08000040); header.Field(4, 0x8004); header.Field(4, (uint)(count + 2));
                byte[] frame = Slice(OneFrameParent(), 28, OneFrameParent().Length - 28);
                writer.Write(header.Finish()); writer.Write((uint)sites);
                writer.Write((uint)(depth.Length + count * template.Length + (parent ? frame.Length : 0))); writer.Write(depth);
                for (int i = 0; i < count; i++)
                { Write32(template, 8, (uint)(i + 2)); writer.Write(template); o.Write(leaf, 0, leaf.Length); }
                if (parent) writer.Write(frame);
                objects = o.ToArray(); return f.ToArray();
            }
        }

        private static byte[] OneFrameParent(uint id = 1)
        {
            var header = new BlockBuilder(0x400, 0x08000040); header.Field(4, 0x8004); header.Field(4, 1);
            var site = new BlockBuilder(0, 0x94); site.Field(4, id); site.Field(4, 0x40023); site.Field(2, 14); site.Align(4);
            using (var buffer = new System.IO.MemoryStream()) using (var writer = new System.IO.BinaryWriter(buffer))
            {
                byte[] child = site.Finish(); writer.Write(header.Finish()); writer.Write(1u); writer.Write((uint)(4 + child.Length));
                writer.Write(new byte[] { 0, 1, 0xe7, 0xe7 }); writer.Write(child); return buffer.ToArray();
            }
        }

        private static byte[] Metadata(string guid) { return Join(new Guid(guid).ToByteArray(), new byte[4]); }

        private static Dictionary<string, byte[]> CopyStreams(IReadOnlyDictionary<string, byte[]> streams)
        { return streams.ToDictionary(x => x.Key, x => x.Value == null ? null : (byte[])x.Value.Clone(), StringComparer.Ordinal); }

        private static void AssertGraphFallback(IReadOnlyDictionary<string, byte[]> streams, IReadOnlyDictionary<string, byte[]> metadata)
        {
            var backup = CopyStreams(streams); var result = FormStreamPadding.NormalizeGraph(streams, metadata);
            Assert.AreSame(streams, result, "Unknown/malformed graph must return its exact original dictionary.");
            foreach (var entry in streams)
            {
                Assert.AreSame(entry.Value, result[entry.Key]);
                if (entry.Value != null) CollectionAssert.AreEqual(backup[entry.Key], entry.Value);
            }
        }
    }
}
