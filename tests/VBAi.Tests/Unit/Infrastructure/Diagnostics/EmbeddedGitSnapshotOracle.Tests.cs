using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class EmbeddedGitSnapshotOracleTests
    {
        [DataTestMethod, DataRow("code whitespace"), DataRow("class code"), DataRow("caption"), DataRow("logical resource"), DataRow("references")]
        public void IndependentExportOracleRefusesEveryChangedSourceMetadataOrLogicalResource(string change)
        {
            byte[] resource = FormResourcePreflightTests.Resource();
            var baseline = Snapshot(resource, "Caption", "Option Explicit\n", "' Class body\n", "");
            byte[] other = (byte[])resource.Clone(); if (change == "logical resource") other[24 + 2048] ^= 1;
            var actual = Snapshot(other, change == "caption" ? "Changed caption" : "Caption",
                change == "code whitespace" ? "Option Explicit\n \n" : "Option Explicit\n",
                change == "class code" ? "' Changed body\n" : "' Class body\n",
                change == "references" ? "{00000000-0000-0000-0000-000000000000}:1:0" : "");
            Assert.ThrowsException<AssertFailedException>(() => EmbeddedGitSnapshotOracle.Verify(baseline, actual));
        }

        [TestMethod]
        public void ComparisonUsesTheRealNonzeroFrmOffsetAndPreservesBothRawResourceArrays()
        {
            byte[] resource = FormResourcePreflightTests.Resource(); var other = (byte[])resource.Clone();
            other[24 + 1024 + 100] = 31; other[24 + 2048 + 7] = 99;
            var before = (byte[])resource.Clone(); var after = (byte[])other.Clone();
            var baseline = Snapshot(resource, "Caption", "Option Explicit\n", "' Class body\n", "");
            var actual = Snapshot(other, "Caption", "Option Explicit\n", "' Class body\n", "");
            EmbeddedGitSnapshotOracle.Verify(baseline, actual);
            CollectionAssert.AreEqual(new[] { 13 }, EmbeddedGitSnapshotOracle.Offsets(VbaGitSnapshot.Utf8.GetString(baseline.Files["EmbeddedForm.frm"])));
            CollectionAssert.AreEqual(before, resource); CollectionAssert.AreEqual(after, other);
            Assert.AreNotEqual(Convert.ToBase64String(baseline.Files["EmbeddedForm.frx"]), Convert.ToBase64String(actual.Files["EmbeddedForm.frx"]));
        }

        [TestMethod]
        public void OffsetOracleIgnoresProcedureTextAndOtherResourceDeclarations()
        {
            string text = "Caption = \"other.frx\":0000\n OleObjectBlob = \"F.frx\":000D\nAttribute VB_Name = \"F\"\nOleObjectBlob = \"F.frx\":001A\n";
            CollectionAssert.AreEqual(new[] { 13 }, EmbeddedGitSnapshotOracle.Offsets(text));
        }

        [TestMethod]
        public void SnapshotOracleRefusesAnEntireMissingComponent()
        {
            var baseline = Snapshot(FormResourcePreflightTests.Resource(), "Caption", "Option Explicit\n", "' Class body\n", "");
            var files = new Dictionary<string, byte[]>(baseline.Files); files.Remove("EmbeddedClass.cls");
            var actual = new VbaGitSnapshot(new VbaGitManifest
            {
                References = baseline.Manifest.References,
                Components = new[] { baseline.Manifest.Components[0], baseline.Manifest.Components[2] }
            }, files);
            Assert.ThrowsException<AssertFailedException>(() => EmbeddedGitSnapshotOracle.Verify(baseline, actual));
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void RetainedAuthorizedUserFormBranchRefusesNearNamesAndChangedRevision(bool changedRevision)
        {
            var manifest = Manifest();
            manifest["embeddedBranch"] = "qualification-userform-20260929225354-93ed53dc";
            manifest["embeddedBranchCommit"] = "f5fb1a004dcb287c4c820b4bf308c673ce3b64e6";
            Assert.AreEqual(manifest["embeddedBranch"], EmbeddedGitWindowTests.ValidateManifest(manifest).Branch);
            manifest[changedRevision ? "embeddedBranchCommit" : "embeddedBranch"] = changedRevision
                ? new string('b', 40) : "qualification-userform-20260929225354-93ed53dc-other";
            Assert.ThrowsException<InvalidOperationException>(() => EmbeddedGitWindowTests.ValidateManifest(manifest));
        }

        [DataTestMethod, DataRow("repo"), DataRow("id"), DataRow("main"), DataRow("revision"), DataRow("tab")]
        public void ManifestRefusesDifferentRepositoryBranchRevisionAndUnobservedTab(string change)
        {
            var manifest = Manifest();
            if (change == "repo") manifest["repositoryUrl"] = "https://github.com/another/repository";
            if (change == "id") manifest["repositoryId"] = "other";
            if (change == "main") manifest["embeddedBranch"] = "main";
            if (change == "revision") manifest["embeddedBranchCommit"] = "HEAD";
            if (change == "tab") manifest["checkpointTabName"] = "";
            Assert.ThrowsException<InvalidOperationException>(() => EmbeddedGitWindowTests.ValidateManifest(manifest));
        }

        [TestMethod]
        public void ManifestHasExactlyTheExplicitSyntheticScopeWithoutFallback()
        {
            var plan = EmbeddedGitWindowTests.ValidateManifest(Manifest());
            Assert.AreEqual("qualification-embedded-ui-fixture", plan.Branch);
            Assert.AreEqual(new string('a', 40), plan.Commit);
        }
        private static Dictionary<string, object> Manifest() => new Dictionary<string, object>
        {
            ["repositoryUrl"] = "https://github.com/blackcancer/vbai-qualification-20260929203712-7267b1e6",
            ["repositoryId"] = "1396566119",
            ["embeddedBranch"] = "qualification-embedded-ui-fixture",
            ["embeddedBranchCommit"] = new string('a', 40),
            ["checkpointTabName"] = "Checkpoints"
        };

        private static VbaGitSnapshot Snapshot(byte[] resource, string caption, string module, string classCode, string references)
        {
            var bytes = new byte[resource.Length + 13]; Buffer.BlockCopy(resource, 0, bytes, 13, resource.Length);
            return new VbaGitSnapshot(new VbaGitManifest
            {
                References = references,
                Components = new[] {
                new VbaGitComponent { Name = "EmbeddedModule", Type = 1 }, new VbaGitComponent { Name = "EmbeddedClass", Type = 2 },
                new VbaGitComponent { Name = "EmbeddedForm", Type = 3, HasResources = true } }
            }, new Dictionary<string, byte[]>
            {
                ["EmbeddedModule.bas"] = VbaGitSnapshot.Utf8.GetBytes("Attribute VB_Name = \"EmbeddedModule\"\n" + module),
                ["EmbeddedClass.cls"] = VbaGitSnapshot.Utf8.GetBytes("Attribute VB_Name = \"EmbeddedClass\"\n" + classCode),
                ["EmbeddedForm.frm"] = VbaGitSnapshot.Utf8.GetBytes("VERSION 5.00\nBegin VB.UserForm EmbeddedForm\n Caption = \"" + caption + "\"\n OleObjectBlob = \"EmbeddedForm.frx\":000D\nEnd\nAttribute VB_Name = \"EmbeddedForm\"\nOption Explicit\n"),
                ["EmbeddedForm.frx"] = bytes
            });
        }
    }
}
