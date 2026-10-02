using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class FormFontObservationTests
    {
        [TestMethod]
        public void NoManifestReturnsBeforeProjectOrNativeAccess()
        {
            Assert.IsNull(FormFontObservation.TryBeginAtPath(null, null, null, null, null));
            Assert.IsNull(FormFontObservation.TryBeginAtPath("", null, null, null, null));
        }

        [TestMethod]
        public void ExactDisposableTargetManifestAcceptsBothPredeclaredModesWithoutCreatingOutput()
        {
            var target = Target();
            foreach (string mode in new[] { FormFontObservation.ObserveWrites, FormFontObservation.DistinctChildName })
            {
                var manifest = Valid(target, mode);
                FormFontObservation.ValidateManifest(manifest, manifest.ProjectPath, target,
                    new HashSet<string>(StringComparer.Ordinal) { "Form1" });
                Assert.IsFalse(Directory.Exists(manifest.OutputRoot));
            }
        }

        [TestMethod]
        public void ManifestRejectsWrongProjectSourceDescriptorCandidateModeAndReusedOutput()
        {
            var target = Target();
            foreach (string field in new[] { "project", "source", "descriptor", "mvid", "mode", "changed", "temporary", "output" })
            {
                var manifest = Valid(target, FormFontObservation.DistinctChildName);
                var changed = new HashSet<string>(StringComparer.Ordinal) { "Form1" };
                if (field == "project") manifest.ProjectPath = Path.Combine(Path.GetTempPath(), "other.xlsm");
                if (field == "source") manifest.TargetFormSha256 = new string('0', 64);
                if (field == "descriptor") manifest.TargetDescriptorHex = "01000000900144420100065461686F6D62";
                if (field == "mvid") manifest.CandidateMvid = Guid.NewGuid().ToString();
                if (field == "mode") manifest.Mode = "Replay";
                if (field == "changed") changed.Add("Module1");
                if (field == "temporary") manifest.TemporaryName = "Tahoma";
                if (field == "output") manifest.OutputRoot = Path.Combine(Path.GetTempPath(), "wrong-" + manifest.Nonce);
                Assert.ThrowsException<InvalidOperationException>(() =>
                    FormFontObservation.ValidateManifest(manifest, ProjectPath(), target, changed), field);
            }
        }

        [TestMethod]
        public void ObserveWritesRejectsAnyTemporaryName()
        {
            var target = Target();
            var manifest = Valid(target, FormFontObservation.ObserveWrites);
            manifest.TemporaryName = "Arial";
            Assert.ThrowsException<InvalidOperationException>(() =>
                FormFontObservation.ValidateManifest(manifest, manifest.ProjectPath, target,
                    new HashSet<string>(StringComparer.Ordinal) { "Form1" }));
        }

        [TestMethod]
        public void ManifestPathGuardRejectsAlternateStreamAndNonCanonicalOutputBeforeMutation()
        {
            string file = Path.GetTempFileName();
            try
            {
                FormFontObservation.RequireSafeLocalPath(file, true);
                Assert.ThrowsException<InvalidOperationException>(() =>
                    FormFontObservation.RequireSafeLocalPath(file + ":other", true));
                Assert.ThrowsException<InvalidOperationException>(() =>
                    FormFontObservation.RequireSafeLocalPath(Path.Combine(Path.GetDirectoryName(file), ".", "trial"), false));
            }
            finally { File.Delete(file); }
        }

        private static string ProjectPath() { return Path.Combine(Path.GetTempPath(), "owned-root-font-observation.xlsm"); }
        private static VbaGitSnapshot Target()
        {
            var form = new VbaGitComponent { Name = "Form1", Type = 3, HasResources = true };
            var files = new Dictionary<string, byte[]>(StringComparer.Ordinal) {
                ["Form1.frm"] = VbaGitSnapshot.Utf8.GetBytes(
                    "OleObjectBlob = \"Form1.frx\":0000\nAttribute VB_Name = \"Form1\"\nOption Explicit\n"),
                ["Form1.frx"] = FormStreamPaddingTests.ContainerResourceBefore()
            };
            return new VbaGitSnapshot(new VbaGitManifest { References = "", Components = new[] { form } }, files);
        }
        private static FormFontObservation.Manifest Valid(VbaGitSnapshot target, string mode)
        {
            string nonce = Guid.NewGuid().ToString("N");
            byte[] descriptor = target.FormFonts(target.Manifest.Components[0]).Single(item => item.OwnerPath == "").Descriptor;
            return new FormFontObservation.Manifest {
                ProjectPath = ProjectPath(), FormName = "Form1", Mode = mode,
                TargetFormSha256 = Hash(target.Files["Form1.frm"]),
                TargetDescriptorHex = BitConverter.ToString(descriptor).Replace("-", ""),
                CandidateMvid = typeof(FormFontObservation).Module.ModuleVersionId.ToString(),
                OutputRoot = Path.Combine(Path.GetTempPath(), nonce), Nonce = nonce,
                TemporaryName = mode == FormFontObservation.DistinctChildName ? "Arial" : null
            };
        }
        private static string Hash(byte[] bytes)
        {
            using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", "");
        }
    }
}
