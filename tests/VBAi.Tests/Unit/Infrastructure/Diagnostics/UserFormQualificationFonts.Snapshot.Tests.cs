using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
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
            WithQualification(() => {
                var snapshot = Snapshot(implicitRoot);
                var original = (byte[])snapshot.Files["Form1.frx"].Clone();
                int calls = 0;
                UserFormQualificationFonts.RequireSnapshot(snapshot, (form, layout) => {
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
            WithQualification(() => {
                var snapshot = Snapshot(implicitRoot);
                Assert.ThrowsException<AssertFailedException>(() => UserFormQualificationFonts.RequireSnapshot(snapshot, null));
                Assert.ThrowsException<AssertFailedException>(() => UserFormQualificationFonts.RequireSnapshot(snapshot, (form, layout) => null));
                foreach (var change in new Dictionary<string, object> {
                    ["Name"] = "Arial", ["Size"] = 9d, ["Bold"] = true, ["Italic"] = true,
                    ["Underline"] = true, ["Strikethrough"] = true, ["Charset"] = 1, ["Weight"] = 700 })
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
            WithQualification(() => {
                byte[] resource = FormStreamPaddingTests.ContainerResourceBefore();
                byte[] descriptor = UserFormQualificationFonts.Descriptor(owner == "Root" ? 8.25m : 8.27m);
                int location = Find(resource, descriptor);
                Buffer.BlockCopy(BitConverter.GetBytes(owner == "Root" ? 90000u : 82500u), 0, resource, location + 6, 4);
                var snapshot = Snapshot(resource);
                Assert.ThrowsException<AssertFailedException>(() => UserFormQualificationFonts.RequireSnapshot(snapshot, (form, layout) => Fonts()));
            });
        }

        [TestMethod]
        public void OpaqueResourceAndMissingDeclaredFrameCannotProducePartialFontAcceptance()
        {
            WithQualification(() => {
                var opaque = Snapshot(FormStreamPaddingTests.ContainerResourceBefore(), "Picture = \"Form1.frx\":0001\n");
                int calls = 0;
                Assert.ThrowsException<AssertFailedException>(() => UserFormQualificationFonts.RequireSnapshot(opaque, (form, layout) => { calls++; return Fonts(); }));
                Assert.AreEqual(0, calls, "An unsupported complete font plan is refused before native reads.");
                byte[] resource = FormStreamPaddingTests.ContainerResourceBefore();
                int directory = DirectoryStream(resource, FormStreamPaddingTests.ContainerStreamsBefore()["/i03/f"].Length);
                int start = Find(resource, FormStreamPaddingTests.ContainerStreamsBefore()["/i03/f"]);
                RemoveFont(resource, start, directory, FormStreamPaddingTests.ContainerStreamsBefore()["/i03/f"], 8.27m);
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
                RemoveFont(resource, Find(resource, root), DirectoryStream(resource, root.Length), root, 8.25m);
            }
            return Snapshot(resource);
        }

        private static VbaGitSnapshot Snapshot(byte[] resource, string metadata = "")
        {
            var component = new VbaGitComponent { Name = "Form1", Type = 3, HasResources = true };
            return new VbaGitSnapshot(new VbaGitManifest { References = "", Components = new[] { component } },
                new Dictionary<string, byte[]> {
                    ["Form1.frm"] = VbaGitSnapshot.Utf8.GetBytes("OleObjectBlob = \"Form1.frx\":0000\n" + metadata + "Attribute VB_Name = \"Form1\"\nOption Explicit\n"),
                    ["Form1.frx"] = resource });
        }

        /// <summary>Removes only the declared Form font and updates the existing bounded CFB stream extent.</summary>
        private static void RemoveFont(byte[] resource, int streamStart, int directory, byte[] stream, decimal size)
        {
            byte[] descriptor = UserFormQualificationFonts.Descriptor(size);
            int payload = Find(stream, descriptor) - 16;
            Assert.IsTrue(payload >= 8);
            int length = 16 + descriptor.Length;
            byte[] omitted = stream.Take(payload).Concat(stream.Skip(payload + length)).ToArray();
            uint mask = BitConverter.ToUInt32(omitted, 4);
            Assert.IsTrue((mask & (1u << 20)) != 0);
            Buffer.BlockCopy(BitConverter.GetBytes(mask & ~(1u << 20)), 0, omitted, 4, 4);
            // Keep the same mini-FAT chain valid: both lengths occupy the same number of mini sectors.
            Assert.AreEqual((stream.Length + 63) / 64, (omitted.Length + 63) / 64);
            Array.Clear(resource, streamStart, stream.Length);
            Buffer.BlockCopy(omitted, 0, resource, streamStart, omitted.Length);
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
