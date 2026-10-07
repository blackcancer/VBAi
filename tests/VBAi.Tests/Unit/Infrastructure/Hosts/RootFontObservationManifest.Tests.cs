using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Managed manifest qualification only: these tests never instantiate an Office object or font setter.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class RootFontObservationManifestTests
    {
        [DataTestMethod]
        [DataRow(null)]
        [DataRow("")]
        [DataRow("0")]
        [DataRow("true")]
        public void DisabledDiagnosticReadsOnlyOptInAndDoesNotInspectPaths(string flag)
        {
            int reads = 0;
            Assert.IsNull(RootFontObservationManifest.Prepare(key =>
            {
                Assert.AreEqual(RootFontObservationManifest.OptIn, key); reads++; return flag;
            }, null, true, path => { Assert.Fail("Disabled diagnostics must not inspect metadata."); return 0; }));
            Assert.AreEqual(1, reads);
        }

        [DataTestMethod]
        [DataRow(null, false)]
        [DataRow("LabelButton", true)]
        [DataRow("unknown", false)]
        public void CapturePersistenceAndUndeclaredLayoutsRefuseBeforeAnyMetadata(string layout, bool persistence)
        {
            Assert.ThrowsException<InvalidOperationException>(() => RootFontObservationManifest.Prepare(key =>
            {
                Assert.AreEqual(RootFontObservationManifest.OptIn, key); return "1";
            }, layout, persistence, path => { Assert.Fail("No metadata after an invalid scenario."); return 0; }));
        }

        [DataTestMethod]
        [DataRow(null)]
        [DataRow("")]
        [DataRow("observewrites")]
        [DataRow("DistinctName")]
        public void ModeMustBeExplicitAndExact(string mode)
        {
            Assert.ThrowsException<InvalidOperationException>(() => RootFontObservationManifest.Prepare(
                key => key == RootFontObservationManifest.OptIn ? "1" : mode, "LabelButton", false,
                path => { Assert.Fail("Invalid modes must not inspect metadata."); return 0; }));
        }

        [TestMethod]
        public void ExplicitSeedIsBoundBeforeBootstrapOnlyForDeferredDiagnostic()
        {
            WithCase((root, path, project) =>
            {
                var armed = RootFontObservationManifest.Prepare(
                    EnvironmentFor(path, "AfterInitialCapture", RootFontObservationManifest.SyntheticExplicitArial9),
                    "LabelButton", false);
                Assert.AreEqual(RootFontObservationManifest.SyntheticExplicitArial9, armed.SeedProfile);
                Assert.IsNull(RootFontObservationManifest.Prepare(
                    EnvironmentFor(path, "AfterInitialCapture"), "LabelButton", false).SeedProfile);
                foreach (string mode in new[] { "ObserveWrites", "DistinctChildName" })
                    Assert.ThrowsException<InvalidOperationException>(() => RootFontObservationManifest.Prepare(
                        EnvironmentFor(path, mode, RootFontObservationManifest.SyntheticExplicitArial9), "LabelButton", false));
                Assert.ThrowsException<InvalidOperationException>(() => RootFontObservationManifest.Prepare(
                    EnvironmentFor(path, "AfterInitialCapture", "arial"), "LabelButton", false));
                Assert.IsFalse(File.Exists(path));
            });
        }

        [TestMethod]
        public void RetainedSourceRequiresItsExactDeferredProfileAndPinnedBytesBeforeBootstrap()
        {
            WithCase((root, path, project) =>
            {
                string source = Path.Combine(root, "source.xlsm");
                File.WriteAllText(source, "inert, but not the pinned workbook");
                Func<string, string> retained = key => key == RootFontObservationManifest.SourceWorkbookVariable ? source :
                    EnvironmentFor(path, "AfterInitialCapture", RootFontObservationManifest.RetainedSyntheticTahoma825)(key);
                Assert.ThrowsException<InvalidOperationException>(() => RootFontObservationManifest.Prepare(
                    retained, "LabelButton", false));
                Assert.ThrowsException<InvalidOperationException>(() => RootFontObservationManifest.Prepare(
                    retained, "TextBox", false));
                Assert.ThrowsException<InvalidOperationException>(() => RootFontObservationManifest.Prepare(
                    key => key == RootFontObservationManifest.SourceWorkbookVariable ? null : retained(key),
                    "LabelButton", false));
                Assert.ThrowsException<InvalidOperationException>(() => RootFontObservationManifest.Prepare(
                    key => key == RootFontObservationManifest.SeedProfileVariable ? null : retained(key),
                    "LabelButton", false));
                Assert.ThrowsException<InvalidOperationException>(() => RootFontObservationManifest.Prepare(
                    key => key == RootFontObservationManifest.ModeVariable ? "ObserveWrites" : retained(key),
                    "LabelButton", false));
                Assert.IsFalse(File.Exists(path));
            });
        }

        [TestMethod]
        public void SyntheticSeedRequiresExactReopenedRootDescriptorAndNeverInventsOne()
        {
            byte[] expected = { 1, 0, 0, 0, 144, 1, 144, 95, 1, 0, 5, 65, 114, 105, 97, 108 };
            CollectionAssert.AreEqual(expected, RootFontObservationManifest.SyntheticArial9Descriptor());
            var values = RootFontObservationManifest.SyntheticArial9Values(RootFontObservationManifest.SyntheticExplicitArial9);
            Assert.AreEqual("Arial", values["Form.Font.Name"]); Assert.AreEqual(9.00m, values["Form.Font.Size"]);
            Assert.AreEqual((short)400, values["Form.Font.Weight"]); Assert.AreEqual((short)0, values["Form.Font.Charset"]);
            Assert.IsFalse((bool)values["Form.Font.Italic"]); Assert.IsFalse((bool)values["Form.Font.Underline"]);
            Assert.IsFalse((bool)values["Form.Font.Strikethrough"]);
            Assert.ThrowsException<InvalidOperationException>(() => RootFontObservationManifest.SyntheticArial9Values("TahomaDefault"));
            CollectionAssert.AreEqual(expected, RootFontObservationManifest.RequireRoot(
                new[] { Binding("", 7, expected) }, "AfterInitialCapture", RootFontObservationManifest.SyntheticExplicitArial9));
            foreach (var invalid in new[] { null, new FormStreamPadding.FormFontBinding[0],
                new[] { Binding("", 7, Descriptor("Tahoma")) },
                new[] { Binding("", 7, expected), Binding("", 7, expected) } })
                Assert.ThrowsException<InvalidOperationException>(() => RootFontObservationManifest.RequireRoot(
                    invalid, "AfterInitialCapture", RootFontObservationManifest.SyntheticExplicitArial9));
            WithCase((root, path, project) =>
            {
                Guid candidate = Guid.NewGuid();
                Assert.ThrowsException<InvalidOperationException>(() => RootFontObservationManifest.Build(
                    Config(path, "AfterInitialCapture", RootFontObservationManifest.SyntheticExplicitArial9),
                    root, project, Baseline(), candidate, candidate, Guid.NewGuid()));
                Assert.IsFalse(File.Exists(path));
            });
        }

        [DataTestMethod]
        [DataRow("relative.json")]
        [DataRow("C:relative.json")]
        [DataRow("\\\\server\\share\\manifest.json")]
        [DataRow("C:\\manifest.json:stream")]
        [DataRow("C:\\folder\\..\\manifest.json")]
        public void NonExactLocalPathsRefuse(string path)
        {
            Assert.ThrowsException<InvalidOperationException>(() => RootFontObservationManifest.Prepare(
                EnvironmentFor(path, "ObserveWrites"), "LabelButton", false));
        }

        [TestMethod]
        public void EachDeclaredLayoutCanArmBeforeBootstrapWithoutCreatingTheManifest()
        {
            WithCase((root, path, project) =>
            {
                foreach (string layout in new[] { "LabelButton", "TextBox", "ComboBox", "ListBox", "CheckBox", "OptionButton",
                    "ToggleButton", "ScrollBar", "SpinButton", "TabStrip", "Image", "FrameMultiPage" })
                    Assert.AreEqual(path, RootFontObservationManifest.Prepare(EnvironmentFor(path, "ObserveWrites"), layout, false).Path);
                Assert.IsFalse(File.Exists(path)); Assert.AreEqual(1, Directory.GetFiles(root).Length);
            });
        }

        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void ExistingManifestFileOrDirectoryIsNotOverwritten(bool directory)
        {
            WithCase((root, path, project) =>
            {
                if (directory) Directory.CreateDirectory(path); else File.WriteAllText(path, "original");
                Assert.ThrowsException<InvalidOperationException>(() => RootFontObservationManifest.Prepare(
                    EnvironmentFor(path, "ObserveWrites"), "LabelButton", false));
                if (!directory) Assert.AreEqual("original", File.ReadAllText(path));
            });
        }

        [DataTestMethod]
        [DataRow("reparse-leaf")]
        [DataRow("reparse-parent")]
        [DataRow("file-parent")]
        [DataRow("missing-parent")]
        public void UnsafeOrMissingManifestAncestorsRefuseWithoutClaim(string scenario)
        {
            WithCase((root, path, project) =>
            {
                Func<string, FileAttributes> inspect = item =>
                {
                    if (item == path && scenario == "reparse-leaf") return FileAttributes.ReparsePoint;
                    if (item == root && scenario == "reparse-parent") return FileAttributes.Directory | FileAttributes.ReparsePoint;
                    if (item == root && scenario == "file-parent") return FileAttributes.Normal;
                    if (item == root && scenario == "missing-parent") throw new DirectoryNotFoundException();
                    return File.GetAttributes(item);
                };
                Assert.ThrowsException<InvalidOperationException>(() => RootFontObservationManifest.Prepare(
                    EnvironmentFor(path, "ObserveWrites"), "LabelButton", false, inspect));
                Assert.IsFalse(File.Exists(path));
            });
        }

        [DataTestMethod]
        [DataRow("access")]
        [DataRow("io")]
        public void MetadataErrorsRemainErrorsRatherThanBeingTreatedAsAbsence(string failure)
        {
            WithCase((root, path, project) =>
            {
                Exception expected = failure == "access" ? (Exception)new UnauthorizedAccessException("owned synthetic probe") : new IOException("owned synthetic probe");
                try
                {
                    RootFontObservationManifest.Prepare(EnvironmentFor(path, "ObserveWrites"), "LabelButton", false,
                    item => { throw expected; }); Assert.Fail("Metadata uncertainty must refuse.");
                }
                catch (Exception actual) { Assert.AreSame(expected, actual); }
                Assert.IsFalse(File.Exists(path));
            });
        }

        [DataTestMethod]
        [DataRow("ObserveWrites", "")]
        [DataRow("DistinctChildName", "Arial")]
        [DataRow("AfterInitialCapture", "")]
        public void ActualBoundedRootFixtureBindsAllExactContractFieldsWithoutChangingBaseline(string mode, string temporaryFace)
        {
            WithCase((root, path, project) =>
            {
                var baseline = Baseline(); var original = baseline.Files.ToDictionary(item => item.Key, item => (byte[])item.Value.Clone());
                var candidate = Guid.NewGuid(); var nonce = Guid.NewGuid();
                var manifest = RootFontObservationManifest.Build(Config(path, mode), root, project, baseline, candidate, candidate, nonce);
                CollectionAssert.AreEquivalent(new[] { "ProjectPath", "FormName", "TargetFormSha256", "TargetDescriptorHex",
                    "CandidateMvid", "OutputRoot", "Nonce", "Mode", "TemporaryName" }, manifest.Keys.ToArray());
                Assert.AreEqual(project, manifest["ProjectPath"]); Assert.AreEqual("EmbeddedForm", manifest["FormName"]);
                Assert.AreEqual(candidate.ToString("D"), manifest["CandidateMvid"]); Assert.AreEqual(nonce.ToString("N"), manifest["Nonce"]);
                Assert.AreEqual(Path.Combine(root, nonce.ToString("N")), manifest["OutputRoot"]);
                Assert.AreEqual(mode, manifest["Mode"]); Assert.AreEqual(temporaryFace, manifest["TemporaryName"]);
                using (var hash = SHA256.Create()) Assert.AreEqual(Hex(hash.ComputeHash(baseline.Files["EmbeddedForm.frm"])), manifest["TargetFormSha256"]);
                var rootBinding = baseline.FormFonts(baseline.Manifest.Components.Single()).Single(item => item.OwnerPath == "" && item.Type == 7);
                Assert.AreEqual(Hex(rootBinding.Descriptor), manifest["TargetDescriptorHex"]);
                foreach (var file in original) CollectionAssert.AreEqual(file.Value, baseline.Files[file.Key]);
                Assert.IsFalse(File.Exists(path)); Assert.IsFalse(Directory.Exists(Convert.ToString(manifest["OutputRoot"])));
            });
        }

        [DataTestMethod]
        [DataRow("empty-candidate")]
        [DataRow("wrong-candidate")]
        [DataRow("empty-nonce")]
        public void CandidateMismatchOrMissingNonceCannotBuildAClaim(string scenario)
        {
            WithCase((root, path, project) =>
            {
                Guid expected = scenario == "empty-candidate" ? Guid.Empty : Guid.NewGuid();
                Guid loaded = scenario == "wrong-candidate" ? Guid.NewGuid() : expected;
                Guid nonce = scenario == "empty-nonce" ? Guid.Empty : Guid.NewGuid();
                Assert.ThrowsException<InvalidOperationException>(() => RootFontObservationManifest.Build(
                    Config(path), root, project, Baseline(), expected, loaded, nonce));
                Assert.IsFalse(File.Exists(path));
            });
        }

        [DataTestMethod]
        [DataRow("missing-form")]
        [DataRow("ambiguous-form")]
        [DataRow("no-resources")]
        [DataRow("missing-source")]
        [DataRow("empty-source")]
        [DataRow("missing-resource")]
        [DataRow("opaque-resource")]
        [DataRow("ambiguous-blob")]
        public void MissingOrAmbiguousBaselineCannotArmTheDiagnostic(string scenario)
        {
            WithCase((root, path, project) =>
            {
                var baseline = Baseline(); var form = baseline.Manifest.Components[0];
                if (scenario == "missing-form") form.Name = "AnotherForm";
                if (scenario == "ambiguous-form") baseline.Manifest.Components = new[] { form, form };
                if (scenario == "no-resources") form.HasResources = false;
                if (scenario == "missing-source") baseline.Files.Remove("EmbeddedForm.frm");
                if (scenario == "empty-source") baseline.Files["EmbeddedForm.frm"] = new byte[0];
                if (scenario == "missing-resource") baseline.Files.Remove("EmbeddedForm.frx");
                if (scenario == "opaque-resource") baseline.Files["EmbeddedForm.frx"] = new byte[] { 1, 2, 3 };
                if (scenario == "ambiguous-blob") baseline.Files["EmbeddedForm.frm"] = VbaGitSnapshot.Utf8.GetBytes(
                    "OleObjectBlob = \"EmbeddedForm.frx\":0000\nOleObjectBlob = \"EmbeddedForm.frx\":0000\nAttribute VB_Name = \"EmbeddedForm\"\n");
                Guid candidate = Guid.NewGuid();
                Assert.ThrowsException<InvalidOperationException>(() => RootFontObservationManifest.Build(
                    Config(path), root, project, baseline, candidate, candidate, Guid.NewGuid()));
                Assert.IsFalse(File.Exists(path));
            });
        }

        [TestMethod]
        public void MissingDuplicateNonRootAndWrongTypeBindingsNeverInferARoot()
        {
            byte[] descriptor = Descriptor("Tahoma");
            foreach (var bindings in new[] { null, new FormStreamPadding.FormFontBinding[0],
                new[] { Binding("Controls/Frame", 14, descriptor) }, new[] { Binding("", 14, descriptor) },
                new[] { Binding("", 7, descriptor), Binding("", 7, descriptor) }, new[] { Binding("", 7, new byte[11]) } })
                Assert.ThrowsException<InvalidOperationException>(() => RootFontObservationManifest.RequireRoot(bindings));
            var root = Binding("", 7, descriptor);
            byte[] copy = RootFontObservationManifest.RequireRoot(new[] { Binding("Controls/Frame", 14, descriptor), root });
            CollectionAssert.AreEqual(descriptor, copy); copy[0] ^= 1; CollectionAssert.AreEqual(descriptor, root.Descriptor);
        }

        [DataTestMethod]
        [DataRow("Arial")]
        [DataRow("arial")]
        public void DistinctModeRefusesTheFixedFaceWhenItIsAlreadyTheTarget(string face)
        {
            var bindings = new[] { Binding("", 7, Descriptor(face)) };
            Assert.ThrowsException<InvalidOperationException>(() => RootFontObservationManifest.RequireRoot(bindings, "DistinctChildName"));
            CollectionAssert.AreEqual(bindings[0].Descriptor, RootFontObservationManifest.RequireRoot(bindings, "ObserveWrites"));
        }

        [TestMethod]
        public void FreshPublicationCreatesOnlyOneExactUtf8ManifestAndNeverTheOutputDirectory()
        {
            WithCase((root, path, project) =>
            {
                var configuration = Config(path); Guid candidate = Guid.NewGuid();
                var manifest = RootFontObservationManifest.Build(configuration, root, project, Baseline(), candidate, candidate, Guid.NewGuid());
                RootFontObservationManifest.Publish(configuration, manifest, EnvironmentFor(path, configuration.Mode));
                byte[] bytes = File.ReadAllBytes(path); Assert.IsTrue(bytes.Length < 16384); Assert.AreEqual((byte)'{', bytes[0]);
                var parsed = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(new UTF8Encoding(false, true).GetString(bytes));
                CollectionAssert.AreEquivalent(manifest.Keys.ToArray(), parsed.Keys.ToArray());
                foreach (var item in manifest) Assert.AreEqual(item.Value, parsed[item.Key]);
                Assert.IsFalse(Directory.Exists(Convert.ToString(manifest["OutputRoot"])));
                Assert.ThrowsException<InvalidOperationException>(() => RootFontObservationManifest.Publish(configuration, manifest, EnvironmentFor(path, configuration.Mode)));
                CollectionAssert.AreEqual(bytes, File.ReadAllBytes(path));
            });
        }

        [DataTestMethod]
        [DataRow("flag")]
        [DataRow("mode")]
        [DataRow("path")]
        [DataRow("seed")]
        public void ChangedInheritedConfigurationRefusesBeforePublication(string changed)
        {
            WithCase((root, path, project) =>
            {
                var configuration = Config(path); Guid candidate = Guid.NewGuid();
                var manifest = RootFontObservationManifest.Build(configuration, root, project, Baseline(), candidate, candidate, Guid.NewGuid());
                var environment = EnvironmentFor(path, configuration.Mode);
                Assert.ThrowsException<InvalidOperationException>(() => RootFontObservationManifest.Publish(configuration, manifest, key =>
                    key == (changed == "flag" ? RootFontObservationManifest.OptIn : changed == "mode" ? RootFontObservationManifest.ModeVariable :
                        changed == "seed" ? RootFontObservationManifest.SeedProfileVariable :
                        RootFontObservationManifest.ManifestVariable) ? "changed" : environment(key)));
                Assert.IsFalse(File.Exists(path));
            });
        }

        [DataTestMethod]
        [DataRow("existing-output")]
        [DataRow("project-directory")]
        [DataRow("project-missing")]
        [DataRow("reparse-project")]
        [DataRow("output-nonce")]
        [DataRow("output-mode")]
        public void PublicationRevalidatesFreshOutputAndOwnedProjectBeforeClaim(string changed)
        {
            WithCase((root, path, project) =>
            {
                var configuration = Config(path); Guid candidate = Guid.NewGuid();
                var manifest = RootFontObservationManifest.Build(configuration, root, project, Baseline(), candidate, candidate, Guid.NewGuid());
                if (changed == "existing-output") Directory.CreateDirectory(Convert.ToString(manifest["OutputRoot"]));
                if (changed == "project-directory") { File.Delete(project); Directory.CreateDirectory(project); }
                if (changed == "project-missing") File.Delete(project);
                if (changed == "output-nonce") manifest["Nonce"] = Guid.NewGuid().ToString("N");
                if (changed == "output-mode") manifest["Mode"] = "DistinctChildName";
                Assert.ThrowsException<InvalidOperationException>(() => RootFontObservationManifest.Publish(configuration, manifest,
                    EnvironmentFor(path, configuration.Mode), item => item == project && changed == "reparse-project" ? FileAttributes.ReparsePoint : File.GetAttributes(item)));
                Assert.IsFalse(File.Exists(path));
            });
        }

        [TestMethod]
        public void FileCreatedAfterMetadataCheckCannotBeOverwrittenByCreateNew()
        {
            WithCase((root, path, project) =>
            {
                var configuration = Config(path); Guid candidate = Guid.NewGuid();
                var manifest = RootFontObservationManifest.Build(configuration, root, project, Baseline(), candidate, candidate, Guid.NewGuid());
                File.WriteAllText(path, "race claim");
                Assert.ThrowsException<IOException>(() => RootFontObservationManifest.Publish(configuration, manifest, EnvironmentFor(path, configuration.Mode),
                    item => { if (item == path) throw new FileNotFoundException(); return File.GetAttributes(item); }));
                Assert.AreEqual("race claim", File.ReadAllText(path));
            });
        }

        private static RootFontObservationManifest.Configuration Config(string path, string mode = "ObserveWrites", string seed = null)
        { return RootFontObservationManifest.Prepare(EnvironmentFor(path, mode, seed), "LabelButton", false); }
        private static Func<string, string> EnvironmentFor(string path, string mode, string seed = null)
        {
            return key => key == RootFontObservationManifest.OptIn ? "1" : key == RootFontObservationManifest.ModeVariable ? mode :
            key == RootFontObservationManifest.SeedProfileVariable ? seed : key == RootFontObservationManifest.ManifestVariable ? path : null;
        }
        private static VbaGitSnapshot Baseline()
        {
            var form = new VbaGitComponent { Name = "EmbeddedForm", Type = 3, HasResources = true };
            return new VbaGitSnapshot(new VbaGitManifest { References = "", Components = new[] { form } },
                new Dictionary<string, byte[]>
                {
                    ["EmbeddedForm.frm"] = VbaGitSnapshot.Utf8.GetBytes(
                    "OleObjectBlob = \"EmbeddedForm.frx\":0000\r\nAttribute VB_Name = \"EmbeddedForm\"\r\nOption Explicit\r\n"),
                    ["EmbeddedForm.frx"] = FormStreamPaddingTests.ContainerResourceBefore()
                });
        }
        private static FormStreamPadding.FormFontBinding Binding(string owner, uint type, byte[] descriptor)
        { return new FormStreamPadding.FormFontBinding(owner, descriptor, type); }
        private static byte[] Descriptor(string face)
        { return new byte[] { 1, 0, 0, 0, 144, 1, 68, 66, 1, 0, (byte)face.Length }.Concat(Encoding.ASCII.GetBytes(face)).ToArray(); }
        private static string Hex(byte[] bytes) { return BitConverter.ToString(bytes).Replace("-", ""); }
        private static void WithCase(Action<string, string, string> action)
        {
            string root = Path.Combine(Path.GetTempPath(), "VBAi-root-font-manifest-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            string project = Path.Combine(root, "Owned.xlsm"); File.WriteAllText(project, "synthetic bytes, not an Office file");
            try { action(root, Path.Combine(root, "observation-manifest.json"), project); }
            finally { Directory.Delete(root, true); }
        }
    }
}
