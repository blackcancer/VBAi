namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net;
    using System.Net.Http;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Web.Script.Serialization;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>Vérifie la sélection partielle de fichiers avec ressources et références cohérentes.</summary>
    public sealed partial class GitReviewTests
    {
        /// <summary>Conserve les ressources du formulaire sélectionné et les références de la base appropriée.</summary>
        [TestMethod]
        public void PartialSnapshotsKeepFormResourcesAndReferencesConsistent()
        {
            VbaGitSnapshot make(byte resource, string reference)
            {
                return new VbaGitSnapshot(new VbaGitManifest { References = reference, Components = new[] { new VbaGitComponent { Name = "Form1", Type = 3, HasResources = true }, new VbaGitComponent { Name = "Module1", Type = 1 } } }, new Dictionary<string, byte[]> { { "Form1.frm", VbaGitSnapshot.Utf8.GetBytes("Attribute VB_Name = \"Form1\"\n") }, { "Form1.frx", new[] { resource } }, { "Module1.bas", VbaGitSnapshot.Utf8.GetBytes("Attribute VB_Name = \"Module1\"\n") } });
            }

            var old = make(1, "a");
            var current = make(2, "b");
            var selected = VbaGitSnapshot.Select(old, current, new[] { "Form1" });
            Assert.AreEqual((byte)2, selected.Files["Form1.frx"][0]);
            Assert.AreEqual("a", selected.Manifest.References);
            Assert.AreEqual("b", VbaGitSnapshot.Select(old, current, new string[0], true).Manifest.References);
            Assert.ThrowsException<ArgumentException>(() => VbaGitSnapshot.Select(old, current, new[] { "Unknown" }));
        }
    }
}

namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>Vérifie les invariants de manifeste, fichiers, sélection et résumé des snapshots VBA.</summary>
    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbaGitSnapshotCoverageTests
    {
        [TestMethod]
        public void FormComparisonAndChangesUseLogicalDataWhileSerializedFrxRemainsExact()
        {
            byte[] firstBytes = FormResourcePreflightTests.Resource();
            byte[] secondBytes = (byte[])firstBytes.Clone(); secondBytes[24 + 1024 + 108] = 42;
            var first = LogicalForm(firstBytes); var second = LogicalForm(secondBytes);
            Assert.IsTrue(first.SameAs(second)); Assert.IsTrue(first.SameFile(second, "Form1.frx"));
            Assert.AreEqual(0, second.Changes(first).Length);
            CollectionAssert.AreEqual(secondBytes, second.Serialize()["Form1.frx"]);
            secondBytes[24 + 2048]++;
            var changed = LogicalForm(secondBytes);
            Assert.IsFalse(first.SameAs(changed)); CollectionAssert.AreEqual(new[] { "~ Form1.frx" }, changed.Changes(first));
        }

        [TestMethod]
        public void OverlappingUnknownResourcesRetainExactFrxComparison()
        {
            byte[] firstBytes = FormResourcePreflightTests.Resource();
            byte[] secondBytes = (byte[])firstBytes.Clone(); secondBytes[24 + 1024 + 108] = 42;
            Assert.IsFalse(LogicalForm(firstBytes, " Picture = \"Form1.frx\":0003\n").SameAs(
                LogicalForm(secondBytes, " Picture = \"Form1.frx\":0003\n")));
        }

        [TestMethod]
        public void OpaqueResourceReferencesKeepRawComparisonAcrossOffsetsAndWhitespace()
        {
            foreach (string declaration in new[] {
                " Picture = \"Form1.frx\":0000\n",
                " Picture =\n \"Form1.frx\":0007\n",
                " Picture = \"Form1.frx\":0007\n" })
            {
                byte[] original = FormResourcePreflightTests.Resource();
                byte[] first = new byte[original.Length + 4];
                Array.Copy(original, 0, first, 4, original.Length);
                byte[] second = (byte[])first.Clone(); second[4 + 24 + 1024 + 108] = 42;
                // Build directly so the OLE blob begins after the opaque reference.
                VbaGitSnapshot shifted(byte[] bytes)
                {
                    var files = new Dictionary<string, byte[]> {
                        ["Form1.frm"] = VbaGitSnapshot.Utf8.GetBytes("VERSION 5.00\nBegin SyntheticForm\n OleObjectBlob = \"Form1.frx\":0004\n" + declaration + "End\nAttribute VB_Name = \"Form1\"\n"),
                        ["Form1.frx"] = bytes };
                    return new VbaGitSnapshot(Manifest(3, "Form1", true), files);
                }
                var left = shifted(first); var right = shifted(second);
                Assert.IsFalse(left.SameAs(right), declaration);
                Assert.IsFalse(left.SameFile(right, "Form1.frx"), declaration);
                CollectionAssert.AreEqual(new[] { "~ Form1.frx" }, right.Changes(left), declaration);
                CollectionAssert.AreEqual(second, right.ComparisonFiles()["Form1.frx"], declaration);
            }
        }

        internal static VbaGitSnapshot LogicalForm(byte[] resources, string extra = "")
        {
            return new VbaGitSnapshot(new VbaGitManifest { References = "", Components = new[] {
                new VbaGitComponent { Name = "Form1", Type = 3, HasResources = true } } }, new Dictionary<string, byte[]> {
                ["Form1.frm"] = VbaGitSnapshot.Utf8.GetBytes("VERSION 5.00\nBegin SyntheticForm\n OleObjectBlob = \"Form1.frx\":0000\n" + extra + "End\nAttribute VB_Name = \"Form1\"\n"),
                ["Form1.frx"] = resources });
        }

        [TestMethod]
        public void FormResourcePreflightRejectsEmptyInvalidAndOutOfRangeOffsets()
        {
            foreach (var invalid in new[] {
                new { Length = 0, Offset = "0000" },
                new { Length = 4, Offset = "0004" },
                new { Length = 4, Offset = "FFFF" },
                new { Length = 4, Offset = "FFFFFFFFFFFFFFFFFFFFFFFF" },
                new { Length = 4, Offset = "000G" },
                new { Length = 4, Offset = "-1" },
                new { Length = 4, Offset = "" }
            })
            {
                var files = Files("VERSION 5.00\nBegin VB.UserForm Form1\n" +
                    "   OleObjectBlob = \"Form1.frx\":" + invalid.Offset +
                    "\nEnd\nAttribute VB_Name = \"Form1\"\n", "Form1.frm");
                files["Form1.frx"] = new byte[invalid.Length];
                Assert.ThrowsException<InvalidOperationException>(() =>
                    new VbaGitSnapshot(Manifest(3, "Form1", true), files), invalid.Offset);
            }
            var empty = Files("Attribute VB_Name = \"Form1\"\n", "Form1.frm");
            empty["Form1.frx"] = new byte[0];
            Assert.ThrowsException<InvalidOperationException>(() => new VbaGitSnapshot(Manifest(3, "Form1", true), empty));
        }

        [TestMethod]
        public void FormResourcePreflightChecksEveryDesignerOffsetWithoutParsingVbaComments()
        {
            string header = "VERSION 5.00\nBegin VB.UserForm Form1\n" +
                " OleObjectBlob = \"Form1.frx\":0000\n Picture = \"Form1.frx\":0003\nEnd\n";
            string code = "Attribute VB_Name = \"Form1\"\n' Example: \"Form1.frx\":FFFF\n";
            var files = Files(header + code, "Form1.frm");
            files["Form1.frx"] = FormResourcePreflightTests.Resource();
            var accepted = new VbaGitSnapshot(Manifest(3, "Form1", true), files);
            CollectionAssert.AreEqual(files["Form1.frx"], accepted.Files["Form1.frx"]);
            files["Form1.frm"] = VbaGitSnapshot.Utf8.GetBytes(header.Replace(":0003", ":" + files["Form1.frx"].Length.ToString("X4")) + code);
            Assert.ThrowsException<InvalidOperationException>(() => new VbaGitSnapshot(Manifest(3, "Form1", true), files));
        }

        [TestMethod]
        public void FormResourcePreflightRequiresExactCompanionCasing()
        {
            var files = Files("Attribute VB_Name = \"Form1\"\n", "Form1.frm");
            files["Form1.FRX"] = new byte[] { 1 };
            Assert.ThrowsException<InvalidOperationException>(() => new VbaGitSnapshot(Manifest(3, "Form1", true), files));
        }

        [TestMethod]
        public void DuplicateNamesAcrossComponentTypesAreRejectedBeforeImport()
        {
            foreach (string secondName in new[] { "Module1", "module1" })
            foreach (int secondType in new[] { 2, 3, 100 })
            {
                var manifest = new VbaGitManifest { References = "", Components = new[] {
                    new VbaGitComponent { Name = "Module1", Type = 1 },
                    new VbaGitComponent { Name = secondName, Type = secondType } } };
                var files = Files();
                files.Add(manifest.Components[1].FileName, VbaGitSnapshot.Utf8.GetBytes(secondType == 100
                    ? "Option Explicit\n" : "Attribute VB_Name = \"" + secondName + "\"\n"));
                Assert.ThrowsException<InvalidOperationException>(() => new VbaGitSnapshot(manifest, files));
            }
        }

        /// <summary>Crée un manifeste minimal comportant un seul composant.</summary>
        /// <param name="type">Type VBE du composant.</param>
        /// <param name="name">Nom du composant.</param>
        /// <param name="resources">Indique si le formulaire possède des ressources.</param>
        /// <returns>Manifeste de test avec une référence vide.</returns>
        private static VbaGitManifest Manifest(int type = 1, string name = "Module1", bool resources = false)
        { return new VbaGitManifest { References = "", Components = new[] { new VbaGitComponent { Name = name, Type = type, HasResources = resources } } }; }

        /// <summary>Crée un dictionnaire de test contenant un fichier source encodé en UTF-8 strict.</summary>
        /// <param name="text">Texte du fichier source.</param>
        /// <param name="filename">Nom du fichier à ajouter.</param>
        /// <returns>Dictionnaire contenant l’unique fichier de test.</returns>
        private static Dictionary<string, byte[]> Files(string text = "Attribute VB_Name = \"Module1\"\n", string filename = "Module1.bas")
        { return new Dictionary<string, byte[]> { [filename] = VbaGitSnapshot.Utf8.GetBytes(text) }; }

        /// <summary>Vérifie les rejets de manifestes incomplets, identités invalides, ressources et paquets mal formés.</summary>
        [TestMethod]
        public void ManifestComponentIdentityResourceAndPackageValidationMatrix()
        {
            foreach (var manifest in new[] { null, new VbaGitManifest { Format = 2 }, new VbaGitManifest { Components = null },
                new VbaGitManifest { Components = new VbaGitComponent[1025], References = "" },
                new VbaGitManifest { Components = new VbaGitComponent[0], References = null } })
                Assert.ThrowsException<InvalidOperationException>(() => new VbaGitSnapshot(manifest, new Dictionary<string, byte[]>()));
            var missing = Manifest(); missing.Components[0] = null;
            Assert.ThrowsException<InvalidOperationException>(() => new VbaGitSnapshot(missing, Files()));
            foreach (int type in new[] { 1, 2, 3, 100 }) Assert.IsTrue(VbaGitComponent.Extension(type).StartsWith("."));
            Assert.ThrowsException<InvalidOperationException>(() => VbaGitComponent.Extension(0));
            foreach (string name in new[] { null, "", "CON", "bad-name", new string('A', 41) })
                Assert.ThrowsException<InvalidOperationException>(() => new VbaGitSnapshot(Manifest(name: name), Files()));
            var duplicate = Manifest(); duplicate.Components = new[] { duplicate.Components[0], new VbaGitComponent { Name = "module1", Type = 1 } };
            Assert.ThrowsException<InvalidOperationException>(() => new VbaGitSnapshot(duplicate, Files()));
            Assert.ThrowsException<InvalidOperationException>(() => new VbaGitSnapshot(Manifest(resources: true), Files()));
            foreach (var files in new[] { new Dictionary<string, byte[]>(), new Dictionary<string, byte[]> { ["Other.bas"] = new byte[0] },
                new Dictionary<string, byte[]> { ["Module1.bas"] = null }, new Dictionary<string, byte[]> { ["Module1.bas"] = new byte[32 * 1024 * 1024 + 1] } })
                Assert.ThrowsException<InvalidOperationException>(() => new VbaGitSnapshot(Manifest(), files));
            Assert.ThrowsException<InvalidOperationException>(() => new VbaGitSnapshot(Manifest(), Files(filename: "module1.bas")));
            foreach (string text in new[] { "\0", "<<<<<<< local", ">>>>>>> incoming", "Attribute VB_Name = \"Other\"\n" })
                Assert.ThrowsException<InvalidOperationException>(() => new VbaGitSnapshot(Manifest(), Files(text)));
            Assert.ThrowsException<System.Text.DecoderFallbackException>(() => new VbaGitSnapshot(Manifest(), new Dictionary<string, byte[]> { ["Module1.bas"] = new byte[] { 255 } }));
            Assert.ThrowsException<InvalidOperationException>(() => new VbaGitSnapshot(Manifest(100, "Host"), Files("Attribute VB_Name = \"Host\"", "Host.vba")));
            var host = new VbaGitSnapshot(Manifest(100, "Host"), Files("Option Explicit", "Host.vba"));
            Assert.AreEqual("Host.vba", host.Manifest.Components[0].FileName);
            Assert.ThrowsException<InvalidOperationException>(() => VbaGitSnapshot.Read(Files()));
            Assert.IsTrue(host.SameAs(VbaGitSnapshot.Read(host.Serialize())));
        }

        /// <summary>Vérifie l’appartenance du FRX au formulaire et la cohérence entre ressources et blobs texte.</summary>
        [TestMethod]
        public void FormResourceOwnershipAndTextBlobReferencesMatrix()
        {
            foreach (bool resources in new[] { false, true })
            {
                foreach (string reference in new[] { "\"Other.frx\"", "\"Form1.frx\"", "\"Other.bin\":0000", "\"Form1.FRX\":0000" })
                {
                    var files = Files("Attribute VB_Name = \"Form1\"\nPicture = " + reference, "Form1.frm");
                    if (resources) files["Form1.frx"] = new byte[] { 1, 2 };
                    if (resources && reference == "\"Form1.frx\"")
                        Assert.AreEqual(2, new VbaGitSnapshot(Manifest(3, "Form1", true), files).Files.Count);
                    else Assert.ThrowsException<InvalidOperationException>(() => new VbaGitSnapshot(Manifest(3, "Form1", resources), files), reference);
                }
            }
            var valid = Files("Attribute VB_Name = \"Form1\"\nOleObjectBlob = \"Form1.frx\":0000", "Form1.frm"); valid["Form1.frx"] = new byte[] { 1 };
            Assert.AreEqual(2, new VbaGitSnapshot(Manifest(3, "Form1", true), valid).Files.Count);
            Assert.AreEqual(1, new VbaGitSnapshot(Manifest(3, "Form1"), Files("Attribute VB_Name = \"Form1\"", "Form1.frm")).Files.Count);
        }

        /// <summary>Vérifie égalité, classification des changements, sélection partielle et résumés de références.</summary>
        [TestMethod]
        public void SelectionEqualityChangeClassificationAndReferenceSummaryMatrix()
        {
            var one = new VbaGitSnapshot(Manifest(), Files());
            var two = new VbaGitSnapshot(Manifest(), Files("Attribute VB_Name = \"Module1\"\n'changed\n"));
            var empty = new VbaGitSnapshot(new VbaGitManifest { References = "", Components = new VbaGitComponent[0] }, new Dictionary<string, byte[]>());
            Assert.IsFalse(one.SameAs(null)); Assert.IsFalse(one.SameAs(empty)); Assert.IsFalse(one.SameAs(two));
            var other = new VbaGitSnapshot(Manifest(name: "Other"), Files("Attribute VB_Name = \"Other\"", "Other.bas"));
            Assert.IsFalse(one.SameAs(other));
            Assert.IsTrue(one.Changes(null).Length > 0); StringAssert.StartsWith(one.Changes(two)[0], "~ ");
            StringAssert.StartsWith(empty.Changes(one)[0], "− ");
            Assert.ThrowsException<ArgumentException>(() => VbaGitSnapshot.Select(one, null, null));
            Assert.IsTrue(one.SameAs(VbaGitSnapshot.Select(one, two, null)));
            Assert.IsTrue(empty.SameAs(VbaGitSnapshot.Select(null, one, null)));
            Assert.IsTrue(two.SameAs(VbaGitSnapshot.Select(null, two, new[] { "Module1" })));
            var changedReferences = Manifest(); changedReferences.References = "Different";
            var refs = new VbaGitSnapshot(changedReferences, Files());
            StringAssert.Contains(refs.ImportSummary(one), UiText.Get("References differ: align them in the VBE before importing."));
            StringAssert.Contains(refs.ImportSummary(null), UiText.Get("A checkpoint protects this import."));
            StringAssert.Contains(one.ImportSummary(one), UiText.Get("A checkpoint protects this import."));
        }
    }
}
