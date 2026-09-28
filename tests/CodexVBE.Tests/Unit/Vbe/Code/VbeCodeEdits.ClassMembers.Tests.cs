using System;
using System.IO;
using System.Linq;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    /// <summary>Matrice du snapshot complet, export caché, écriture gardée et historique.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class VbeClassMemberRenameTests
    {
        [TestMethod]
        public void PreviewApplyAndHistoryPreserveExactProjectSelectorAndExportCleanup()
        {
            var fixture = new ClassMemberRenameFixture(); var request = ClassMemberRenameFixture.Request();
            request.Project = @"C:\Work\Book.xlsm";
            dynamic preview = fixture.Service.PreviewClassMemberRename(request);
            Assert.AreEqual(0, fixture.Writes); Assert.AreEqual(1, fixture.Exports); Assert.AreEqual(request.Project, (string)preview.Project);
            request.ExpectedProjectVersion = (string)preview.ExpectedProjectVersion;
            dynamic result = fixture.Service.ApplyClassMemberRename(request);
            Assert.IsTrue((bool)result.ReadbackVerified); Assert.IsFalse((bool)result.Atomic); Assert.AreEqual(1, fixture.Writes);
            StringAssert.Contains(fixture.Sources["CalcClass"], "Me.Compute(2)"); Assert.IsTrue(fixture.Selectors.All(x => x == request.Project));
            Assert.IsTrue(fixture.ExportPaths.All(x => !File.Exists(x)));
            fixture.Service.Replay(new Request { Project = request.Project, Module = request.Module,
                ExpectedSha256 = VbaProcedureRename.Digest(fixture.Sources[request.Module]) }, false);
            Assert.AreEqual(ClassMemberRenameFixture.Target, fixture.Sources[request.Module]);
        }

        [TestMethod]
        public void SourceCatalogueMetadataModeAndPreparationRacesRefuseAllWrites()
        {
            foreach (string scenario in new[] { "source", "other", "component", "metadata", "mode", "race", "expected mode", "missing version" })
            {
                var fixture = new ClassMemberRenameFixture(); var request = fixture.PreviewRequest();
                if (scenario == "source") fixture.Sources["CalcClass"] += "\r\n' changed";
                if (scenario == "other") fixture.Sources["Other"] += "\r\n' changed";
                if (scenario == "component") fixture.Sources.Add("New", "Option Explicit");
                if (scenario == "metadata") fixture.MetadataVersion = "meta-v2";
                if (scenario == "mode") fixture.Mode = 1;
                if (scenario == "race") fixture.BeforeCatalogue = () => { if (fixture.Captures == 3) fixture.Sources["Other"] += "\r\n' race"; };
                if (scenario == "expected mode") request.ExpectedMode = 1;
                if (scenario == "missing version") request.ExpectedProjectVersion = null;
                if (scenario == "expected mode" || scenario == "missing version")
                    Assert.ThrowsException<ArgumentException>(() => fixture.Service.ApplyClassMemberRename(request), scenario);
                else Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.ApplyClassMemberRename(request), scenario);
                Assert.AreEqual(0, fixture.Writes); Assert.IsTrue(fixture.ExportPaths.All(x => !File.Exists(x)));
            }
        }

        [TestMethod]
        public void ReadAndExportFailuresRefuseMutationAndRemoveOwnedTemporaryFiles()
        {
            foreach (string command in new[] { "project_properties", "list_modules", "read_module", "component_properties", "export_component" })
            {
                var fixture = new ClassMemberRenameFixture { FailCommand = command };
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.PreviewClassMemberRename(ClassMemberRenameFixture.Request()));
                Assert.AreEqual(0, fixture.Writes); Assert.IsTrue(fixture.ExportPaths.All(x => !File.Exists(x)));
            }
            foreach (int scenario in Enumerable.Range(0, 5))
            {
                var fixture = new ClassMemberRenameFixture();
                if (scenario == 0) fixture.MissingExport = true;
                if (scenario == 1) fixture.WrongExportBody = true;
                if (scenario == 2) fixture.HiddenAttributes = "Attribute Calculate.VB_UserMemId = 0\r\n";
                if (scenario == 3) fixture.HeaderOverride = "Attribute VB_Exposed = True";
                if (scenario == 4) fixture.HeaderOverride = "Attribute VB_PredeclaredId = True";
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.PreviewClassMemberRename(ClassMemberRenameFixture.Request()));
                Assert.AreEqual(0, fixture.Writes); Assert.IsTrue(fixture.ExportPaths.All(x => !File.Exists(x)));
            }
        }

        [TestMethod]
        public void WriteOrPostWriteAttributeFailureReportsUncertaintyWithoutRetry()
        {
            foreach (int scenario in Enumerable.Range(0, 3))
            {
                var fixture = new ClassMemberRenameFixture(); var request = fixture.PreviewRequest();
                if (scenario == 0) fixture.FailCommand = "replace_lines";
                if (scenario == 1) fixture.WrongReadback = true;
                if (scenario == 2) fixture.HiddenAfterWrite = true;
                var failure = Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.ApplyClassMemberRename(request));
                StringAssert.Contains(failure.Message, "failing module may also have changed");
                StringAssert.Contains(failure.Message, "No automatic rollback or retry"); Assert.IsTrue(fixture.Writes <= 1);
                Assert.IsTrue(fixture.ExportPaths.All(x => !File.Exists(x)));
            }
        }

        [TestMethod]
        public void ExportGuardRejectsHiddenUnknownDuplicateAttributesAndMismatchingBodies()
        {
            string export = ClassMemberRenameFixture.Export("CalcClass", ClassMemberRenameFixture.Target);
            VbeCodeEdits.ValidateClassMemberExport(export, "CalcClass", ClassMemberRenameFixture.Target);
            foreach (int scenario in Enumerable.Range(0, 9))
            {
                string invalid = export;
                if (scenario == 0) invalid = null;
                if (scenario == 1) invalid = "";
                if (scenario == 2) invalid = invalid.Replace("VERSION 1.0 CLASS", "VERSION 2.0 CLASS");
                if (scenario == 3) invalid = invalid.Replace("MultiUse = -1", "MultiUse = 0");
                if (scenario == 4) invalid = invalid.Replace("Attribute VB_Exposed = False", "Attribute VB_Exposed = True");
                if (scenario == 5) invalid = invalid.Replace("Attribute VB_Exposed = False", "Attribute VB_Unknown = False");
                if (scenario == 6) invalid = invalid.Replace("Attribute VB_Exposed = False", "Attribute VB_Name = \"CalcClass\"");
                if (scenario == 7) invalid += "Attribute Calculate.VB_Description = \"hidden\"\r\n";
                if (scenario == 8) invalid = invalid.Replace("Calculate = x", "Calculate = 2");
                Assert.ThrowsException<InvalidOperationException>(() => VbeCodeEdits.ValidateClassMemberExport(invalid, "CalcClass", ClassMemberRenameFixture.Target));
            }
        }
        [TestMethod]
        public void ExactClassAndNativeVersionGuardsRefuseBeforeExportOrWrite()
        {
            Assert.ThrowsException<ArgumentException>(() => new ClassMemberRenameFixture().Service.ApplyClassMemberRename(null));
            foreach (int scenario in Enumerable.Range(0, 4))
            {
                var fixture = new ClassMemberRenameFixture();
                if (scenario == 0) fixture.Sources.Remove("CalcClass");
                if (scenario == 1) fixture.TargetType = 1;
                if (scenario == 2) fixture.NativeComponentType = 1;
                if (scenario == 3) fixture.NativeComponentVersion = " ";
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.PreviewClassMemberRename(ClassMemberRenameFixture.Request()), "scenario " + scenario);
                Assert.AreEqual(0, fixture.Exports); Assert.AreEqual(0, fixture.Writes);
            }
        }

        [TestMethod]
        public void OwnedExportPathCollisionPreservesFileAndNullFactoryIsRejected()
        {
            Assert.ThrowsException<ArgumentNullException>(() => new VbeCodeEdits(_ => Response.Success(null), null));
            string path = Path.Combine(Path.GetTempPath(), "CodexVBE-ClassMemberFixture-" + Guid.NewGuid().ToString("N") + ".cls");
            try
            {
                File.WriteAllText(path, "existing file must survive");
                var fixture = new ClassMemberRenameFixture(() => path);
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.PreviewClassMemberRename(ClassMemberRenameFixture.Request()));
                Assert.AreEqual("existing file must survive", File.ReadAllText(path)); Assert.AreEqual(0, fixture.Exports); Assert.AreEqual(0, fixture.Writes);
            }
            finally { if (File.Exists(path)) File.Delete(path); }
        }

        [TestMethod]
        public void OversizedExportRefusesAndRemovesOnlyOwnedTemporaryFile()
        {
            var fixture = new ClassMemberRenameFixture { OversizedExport = true };
            var failure = Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.PreviewClassMemberRename(ClassMemberRenameFixture.Request()));
            StringAssert.Contains(failure.Message, "exceeds its bound"); Assert.AreEqual(1, fixture.Exports); Assert.AreEqual(0, fixture.Writes);
            Assert.IsTrue(fixture.ExportPaths.All(path => !File.Exists(path)));
        }

        [TestMethod]
        public void ExportHeaderNullSourceMissingAttributeWrongNameAndEofAreHandled()
        {
            string export = ClassMemberRenameFixture.Export("CalcClass", ClassMemberRenameFixture.Target);
            Assert.ThrowsException<InvalidOperationException>(() => VbeCodeEdits.ValidateClassMemberExport(export, "CalcClass", null));
            foreach (string invalid in new[] { export.Replace("BEGIN", "BEGIN OTHER"), export.Replace("\r\nEND\r\n", "\r\nEND OTHER\r\n"),
                export.Replace("Attribute VB_GlobalNameSpace = False\r\n", ""), export.Replace("VB_Name = \"CalcClass\"", "VB_Name = \"Wrong\"") })
                Assert.ThrowsException<InvalidOperationException>(() => VbeCodeEdits.ValidateClassMemberExport(invalid, "CalcClass", ClassMemberRenameFixture.Target));
            string headerOnly = export.Substring(0, export.IndexOf("Option Explicit", StringComparison.Ordinal)).TrimEnd('\r', '\n');
            VbeCodeEdits.ValidateClassMemberExport(headerOnly, "CalcClass", "");
        }

        [TestMethod]
        public void SameMemberNamePreviewAndApplyPerformNoWrites()
        {
            var fixture = new ClassMemberRenameFixture(); var request = ClassMemberRenameFixture.Request(); request.NewName = request.Query;
            dynamic preview = fixture.Service.PreviewClassMemberRename(request);
            Assert.IsFalse((bool)preview.Changed); request.ExpectedProjectVersion = (string)preview.ExpectedProjectVersion;
            dynamic result = fixture.Service.ApplyClassMemberRename(request);
            Assert.IsTrue((bool)result.ReadbackVerified); Assert.AreEqual(0, fixture.Writes); Assert.AreEqual(ClassMemberRenameFixture.Target, fixture.Sources["CalcClass"]);
            Assert.IsTrue(fixture.ExportPaths.All(path => !File.Exists(path)));
        }
    }
}