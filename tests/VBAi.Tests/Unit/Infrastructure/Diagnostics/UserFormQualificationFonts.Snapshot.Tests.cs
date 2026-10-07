using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Uses complete synthetic CFB resources to distinguish implicit defaults from lost fonts.</summary>
    public sealed partial class UserFormQualificationFontsTests
    {
        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void ExplicitAndImplicitRootRequireCurrentNativeProofWithoutChangingResources(bool implicitRoot)
        {
            WithQualification(() =>
            {
                var snapshot = Snapshot(implicitRoot);
                var original = (byte[])snapshot.Files["Form1.frx"].Clone();
                int calls = 0;
                UserFormQualificationFonts.RequireSnapshot(snapshot, (form, layout) =>
                {
                    calls++;
                    Assert.AreEqual("Form1", form);
                    Assert.AreEqual("FrameMultiPage", layout, "The declared Frame requires native proof even without a layout marker.");
                    return Fonts();
                });
                Assert.AreEqual(1, calls);
                CollectionAssert.AreEqual(original, snapshot.Files["Form1.frx"]);
                Assert.AreEqual(implicitRoot ? 0 : 1, snapshot.FormFonts(snapshot.Manifest.Components[0]).Count(x => x.OwnerPath == ""));
            });
        }

        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void MissingOrChangedCurrentNativeProofCannotPassAnExactResource(bool implicitRoot)
        {
            WithQualification(() =>
            {
                var snapshot = Snapshot(implicitRoot);
                Assert.ThrowsException<AssertFailedException>(() => UserFormQualificationFonts.RequireSnapshot(snapshot, null));
                Assert.ThrowsException<AssertFailedException>(() => UserFormQualificationFonts.RequireSnapshot(snapshot, (form, layout) => null));
                foreach (var change in new Dictionary<string, object>
                {
                    ["Name"] = "Arial",
                    ["Size"] = 9d,
                    ["Bold"] = true,
                    ["Italic"] = true,
                    ["Underline"] = true,
                    ["Strikethrough"] = true,
                    ["Charset"] = 1,
                    ["Weight"] = 700
                })
                {
                    var native = Fonts(); native["Form.Font." + change.Key] = change.Value;
                    Assert.ThrowsException<AssertFailedException>(() => UserFormQualificationFonts.RequireSnapshot(snapshot, (form, layout) => native), change.Key);
                }
                var rounded = Fonts(); rounded["Frame.Font.Size"] = 8.25d;
                Assert.ThrowsException<AssertFailedException>(() => UserFormQualificationFonts.RequireSnapshot(snapshot, (form, layout) => rounded),
                    "A missing layout marker must not hide a rounded native Frame.");
            });
        }

        [DataTestMethod]
        [DataRow("Root")]
        [DataRow("Frame")]
        public void DeclaredInexactFontCannotBeHiddenByExactNativeReadback(string owner)
        {
            WithQualification(() =>
            {
                byte[] resource = ResourceWithChangedDeclaredFont(owner, owner == "Root" ? 9m : 8.25m);
                var snapshot = Snapshot(resource);
                Assert.ThrowsException<AssertFailedException>(() => UserFormQualificationFonts.RequireSnapshot(snapshot, (form, layout) => Fonts()));
            });
        }

        [TestMethod]
        public void OpaqueResourceAndMissingDeclaredFrameCannotProducePartialFontAcceptance()
        {
            WithQualification(() =>
            {
                var opaque = Snapshot(FormStreamPaddingTests.ContainerResourceBefore(), "Picture = \"Form1.frx\":0001\n");
                int calls = 0;
                Assert.ThrowsException<AssertFailedException>(() => UserFormQualificationFonts.RequireSnapshot(opaque, (form, layout) => { calls++; return Fonts(); }));
                Assert.AreEqual(0, calls, "An unsupported complete font plan is refused before native reads.");
                byte[] resource = FormStreamPaddingTests.ContainerResourceBefore();
                int directory = DirectoryStream(resource, FormStreamPaddingTests.ContainerStreamsBefore()["/i03/f"].Length);
                RemoveFont(resource, directory, 8.27m);
                var missing = Snapshot(resource);
                Assert.ThrowsException<AssertFailedException>(() => UserFormQualificationFonts.RequireSnapshot(missing, (form, layout) => Fonts(), "FrameMultiPage"));
            });
        }

        [TestMethod]
        public void DisabledQualificationNeverReadsTheNativeDesigner()
        {
            string old = Environment.GetEnvironmentVariable(UserFormQualificationFonts.OptIn);
            try
            {
                Environment.SetEnvironmentVariable(UserFormQualificationFonts.OptIn, null);
                UserFormQualificationFonts.RequireSnapshot(Snapshot(true), (form, layout) => throw new InvalidOperationException("Must not read native state."));
            }
            finally { Environment.SetEnvironmentVariable(UserFormQualificationFonts.OptIn, old); }
        }

        /// <summary>Changes one declared font through its bounded logical CFB stream, preserving duplicate raw payloads elsewhere.</summary>
        internal static byte[] ResourceWithChangedDeclaredFont(string owner, decimal replacement)
        {
            if (owner != "Root" && owner != "Frame") throw new ArgumentException("A declared fixture font owner is required.", nameof(owner));
            byte[] resource = FormStreamPaddingTests.ContainerResourceBefore();
            int length = FormStreamPaddingTests.ContainerStreamsBefore()[owner == "Root" ? "/f" : "/i03/f"].Length;
            RemoveFont(resource, DirectoryStream(resource, length), owner == "Root" ? 8.25m : 8.27m, replacement);
            var snapshot = Snapshot(resource);
            var updated = snapshot.FormFonts(snapshot.Manifest.Components[0]);
            var original = FormStreamPadding.ReadFontBindings(FormStreamPaddingTests.ContainerStreamsBefore(), FormStreamPaddingTests.ContainerMetadata());
            Assert.IsNotNull(updated, "The changed fixture must retain a supported complete resource graph.");
            Assert.AreEqual(2, updated.Length);
            string changedPath = owner == "Root" ? "" : "Controls/QualificationExtra";
            foreach (var before in original)
            {
                var after = updated.Single(value => value.OwnerPath == before.OwnerPath && value.Type == before.Type);
                byte[] expected = (byte[])before.Descriptor.Clone();
                if (before.OwnerPath == changedPath)
                    Array.Copy(BitConverter.GetBytes(checked((uint)(replacement * 10000m))), 0, expected, 6, 4);
                CollectionAssert.AreEqual(expected, after.Descriptor, "Only the exact declared owner may change.");
                FormFontRestoration.ValidateDescriptor(after.Descriptor);
            }
            return resource;
        }

        private static void WithQualification(Action scenario)
        {
            string old = Environment.GetEnvironmentVariable(UserFormQualificationFonts.OptIn);
            try { Environment.SetEnvironmentVariable(UserFormQualificationFonts.OptIn, "1"); scenario(); }
            finally { Environment.SetEnvironmentVariable(UserFormQualificationFonts.OptIn, old); }
        }

        private static VbaGitSnapshot Snapshot(bool implicitRoot)
        {
            byte[] resource = FormStreamPaddingTests.ContainerResourceBefore();
            if (implicitRoot)
            {
                byte[] root = FormStreamPaddingTests.ContainerStreamsBefore()["/f"];
                RemoveFont(resource, DirectoryStream(resource, root.Length), 8.25m);
            }
            return Snapshot(resource);
        }

        private static VbaGitSnapshot Snapshot(byte[] resource, string metadata = "")
        {
            var component = new VbaGitComponent { Name = "Form1", Type = 3, HasResources = true };
            return new VbaGitSnapshot(new VbaGitManifest { References = "", Components = new[] { component } },
                new Dictionary<string, byte[]>
                {
                    ["Form1.frm"] = VbaGitSnapshot.Utf8.GetBytes("OleObjectBlob = \"Form1.frx\":0000\n" + metadata + "Attribute VB_Name = \"Form1\"\nOption Explicit\n"),
                    ["Form1.frx"] = resource
                });
        }

        /// <summary>Removes only the declared Form font and updates the existing bounded CFB stream extent.</summary>
        private static void RemoveFont(byte[] resource, int directory, decimal size, decimal? replacementSize = null)
        {
            int length = checked((int)BitConverter.ToUInt32(resource, directory + 120));
            int origin = 24, sectorSize = 512, miniSize = 64;
            int directoryRoot = origin + sectorSize * (checked((int)BitConverter.ToUInt32(resource, origin + 48)) + 1);
            uint miniRoot = BitConverter.ToUInt32(resource, directoryRoot + 116);
            uint fat = BitConverter.ToUInt32(resource, origin + 76);
            uint miniFat = BitConverter.ToUInt32(resource, origin + 60);
            Func<uint, int> sector = value => checked(origin + sectorSize * ((int)value + 1));
            Func<uint, uint> nextSector = value => BitConverter.ToUInt32(resource, sector(fat) + checked((int)value * 4));
            var rootSectors = new List<uint>();
            for (uint current = miniRoot; current != 0xfffffffe; current = nextSector(current))
            {
                Assert.IsTrue(rootSectors.Count < resource.Length / sectorSize && !rootSectors.Contains(current));
                rootSectors.Add(current);
            }
            var addresses = new List<int>();
            uint mini = BitConverter.ToUInt32(resource, directory + 116);
            var seen = new HashSet<uint>();
            while (mini != 0xfffffffe)
            {
                Assert.IsTrue(seen.Add(mini) && mini / 8 < rootSectors.Count);
                addresses.Add(sector(rootSectors[checked((int)(mini / 8))]) + checked((int)(mini % 8) * miniSize));
                mini = BitConverter.ToUInt32(resource, sector(miniFat) + checked((int)mini * 4));
            }
            Assert.AreEqual((length + miniSize - 1) / miniSize, addresses.Count);
            byte[] stream = new byte[length];
            for (int i = 0; i < length; i++) stream[i] = resource[addresses[i / miniSize] + i % miniSize];
            byte[] descriptor = UserFormQualificationFonts.Descriptor(size);
            int payload = Find(stream, descriptor) - 16;
            Assert.IsTrue(payload >= 8);
            byte[] omitted;
            if (replacementSize.HasValue)
            {
                omitted = stream;
                Buffer.BlockCopy(BitConverter.GetBytes(checked((uint)(replacementSize.Value * 10000m))),
                    0, omitted, payload + 16 + 6, 4);
            }
            else
            {
                omitted = stream.Take(payload).Concat(stream.Skip(payload + 16 + descriptor.Length)).ToArray();
                // fFont has a 0xffff scalar plus two alignment bytes in FormDataBlock.
                int marker = Find(omitted, new byte[] { 0xff, 0xff, 0, 0 });
                omitted = omitted.Take(marker).Concat(omitted.Skip(marker + 4)).ToArray();
                ushort blockLength = BitConverter.ToUInt16(omitted, 2);
                Buffer.BlockCopy(BitConverter.GetBytes(checked((ushort)(blockLength - 4))), 0, omitted, 2, 2);
                uint mask = BitConverter.ToUInt32(omitted, 4);
                Assert.IsTrue((mask & (1u << 20)) != 0);
                Buffer.BlockCopy(BitConverter.GetBytes(mask & ~(1u << 20)), 0, omitted, 4, 4);
            }
            Assert.AreEqual((stream.Length + miniSize - 1) / miniSize, (omitted.Length + miniSize - 1) / miniSize);
            // Preserve the actual mini-FAT graph: native streams can use noncontiguous mini sectors.
            for (int i = 0; i < addresses.Count * miniSize; i++)
                resource[addresses[i / miniSize] + i % miniSize] = i < omitted.Length ? omitted[i] : (byte)0;
            Buffer.BlockCopy(BitConverter.GetBytes((uint)omitted.Length), 0, resource, directory + 120, 4);
        }

        private static int DirectoryStream(byte[] resource, int length)
        {
            var matches = Enumerable.Range(0, resource.Length - 128).Where(i => resource[i] == (byte)'f' && resource[i + 1] == 0 &&
                resource[i + 2] == 0 && resource[i + 3] == 0 && resource[i + 64] == 4 && resource[i + 66] == 2 &&
                BitConverter.ToUInt32(resource, i + 120) == length).ToArray();
            Assert.AreEqual(1, matches.Length, "The fixture must identify one exact native f stream.");
            return matches[0];
        }

        private static int Find(byte[] bytes, byte[] needle)
        {
            var matches = Enumerable.Range(0, bytes.Length - needle.Length + 1)
                .Where(i => needle.Select((value, offset) => bytes[i + offset] == value).All(equal => equal)).ToArray();
            Assert.AreEqual(1, matches.Length, "One exact declared synthetic resource span is required.");
            return matches[0];
        }
    }
}
