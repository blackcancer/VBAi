using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace VBAi.Tests.Unit
{
    /// <summary>Exerce la prévisualisation, la double vérification du projet et les mutations partielles déclarées.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class VbeProcedureRenameWorkflowTests
    {
        [TestMethod]
        public void IncompleteMetadataAndComponentCataloguesAreRejectedBeforeSourceMutation()
        {
            var fixture = new ProcedureRenameWorkflowFixture();
            foreach (Request invalid in new[] { null, new Request { Module = "MathModule" }, new Request { Project = "P" } })
                Assert.ThrowsException<ArgumentException>(() => fixture.Service.PreviewProcedureRename(invalid));
            Assert.ThrowsException<ArgumentException>(() => fixture.Service.ApplyProcedureRename(null));
            foreach (object metadata in new object[] { null, new object[0], new { Mode = 2, Version = "v" }, new { Project = "P", Version = "v" }, new { Project = "P", Mode = 2 }, new { Project = "P", Mode = 2, Version = " " }, new { Project = "", Mode = 2, Version = "v" } })
            {
                fixture = new ProcedureRenameWorkflowFixture(); fixture.Override = request => request.Command == "project_properties" ? Response.Success(metadata) : null;
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.PreviewProcedureRename(ProcedureRenameMatrix.Request())); Assert.AreEqual(0, fixture.Writes);
            }
            foreach (object catalogue in new object[] { null, new { Name = "MathModule" }, new object[0], new object[1001], new object[] { 1 }, new object[] { new { Type = 1 } }, new object[] { new { Name = "MathModule" } }, new object[] { new { Name = " ", Type = 1 } }, new object[] { new { Name = "MathModule", Type = 1 }, new { Name = "MATHMODULE", Type = 1 } } })
            {
                fixture = new ProcedureRenameWorkflowFixture(); fixture.Override = request => request.Command == "list_modules" ? Response.Success(catalogue) : null;
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.PreviewProcedureRename(ProcedureRenameMatrix.Request())); Assert.AreEqual(0, fixture.Writes);
            }
            fixture = new ProcedureRenameWorkflowFixture(); int reads = 0;
            fixture.Override = request => request.Command == "project_properties" ? Response.Success(new { Project = ++reads == 1 ? "P" : "Different", Mode = 2, Version = "v" }) : null;
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.PreviewProcedureRename(ProcedureRenameMatrix.Request())); Assert.AreEqual(0, fixture.Writes);
            fixture = new ProcedureRenameWorkflowFixture(); reads = 0;
            fixture.Override = request => request.Command == "project_properties" ? Response.Success(new { Project = "P", Mode = 2, Version = ++reads == 1 ? "v1" : "v2" }) : null;
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.PreviewProcedureRename(ProcedureRenameMatrix.Request())); Assert.AreEqual(0, fixture.Writes);
            fixture = new ProcedureRenameWorkflowFixture(); var apply = fixture.PreviewRequest();
            reads = 0;
            fixture.Override = request => { if (request.Command == "project_properties" && ++reads == 3) fixture.Mode = 1; return null; };
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.ApplyProcedureRename(apply)); Assert.AreEqual(0, fixture.Writes);
        }
        /// <summary>Le chemin sélectionne le projet natif; son nom VBA canonique résout les qualifications du code.</summary>
        [TestMethod]
        public void AbsoluteProjectSelectorsArePreservedWhileCanonicalNameResolvesCalls()
        {
            const string selector = @"C:\Work\Native workbook.xlsm";
            var fixture = new ProcedureRenameWorkflowFixture(); var request = ProcedureRenameMatrix.Request(); request.Project = selector;
            dynamic preview = fixture.Service.PreviewProcedureRename(request);
            Assert.AreEqual(selector, (string)preview.Project); Assert.AreEqual("P", (string)preview.CanonicalProjectName);
            request.ExpectedProjectVersion = (string)preview.ExpectedProjectVersion;
            fixture.Service.ApplyProcedureRename(request);
            StringAssert.Contains(fixture.Sources["Caller"], "P.MathModule.Compute(3)");
            Assert.IsTrue(fixture.Selectors.TrueForAll(value => value == selector));
        }

        /// <summary>Prévisualisation sans écriture, application relue, puis undo indépendant avec son SHA exact.</summary>
        [TestMethod]
        public void PreviewApplyAndPerModuleHistoryUseObservedVersionsAndSources()
        {
            var fixture = new ProcedureRenameWorkflowFixture(); var request = fixture.PreviewRequest();
            Assert.AreEqual(0, fixture.Writes); Assert.AreEqual(ProcedureRenameMatrix.Target, fixture.Sources["MathModule"]);
            dynamic result = fixture.Service.ApplyProcedureRename(request);
            Assert.IsTrue((bool)result.ReadbackVerified); Assert.IsFalse((bool)result.Atomic); Assert.IsFalse((bool)result.MultiModuleUndoAvailable);
            Assert.AreEqual(2, fixture.Writes); StringAssert.Contains(fixture.Sources["Caller"], "MathModule.Compute(2)");
            fixture.Service.Replay(new Request { Project = "P", Module = "Caller", ExpectedSha256 = VbaProcedureRename.Digest(fixture.Sources["Caller"]) }, false);
            Assert.AreEqual(ProcedureRenameMatrix.Caller, fixture.Sources["Caller"]);
            fixture.Service.Replay(new Request { Project = "P", Module = "MathModule", ExpectedSha256 = VbaProcedureRename.Digest(fixture.Sources["MathModule"]) }, false);
            Assert.AreEqual(ProcedureRenameMatrix.Target, fixture.Sources["MathModule"]);
        }

        /// <summary>Un appelant, composant supplémentaire, référence ou mode changé invalide le projet avant toute écriture.</summary>
        [TestMethod]
        public void AllProjectChangesAndPreparationRaceRefuseEveryWrite()
        {
            foreach (string scenario in new[] { "caller", "component", "metadata", "mode", "race", "expected mode", "missing version" })
            {
                var fixture = new ProcedureRenameWorkflowFixture(); var request = fixture.PreviewRequest();
                if (scenario == "caller") fixture.Sources["Caller"] += "\r\n' changed";
                if (scenario == "component") fixture.Sources.Add("NewSheet", "");
                if (scenario == "metadata") fixture.MetadataVersion = "metadata-v2";
                if (scenario == "mode") fixture.Mode = 1;
                if (scenario == "race") fixture.BeforeCatalogue = () => { if (fixture.Captures == 3) fixture.Sources["Caller"] += "\r\n' changed during preparation"; };
                if (scenario == "expected mode") request.ExpectedMode = 1;
                if (scenario == "missing version") request.ExpectedProjectVersion = null;
                if (scenario == "expected mode" || scenario == "missing version")
                    Assert.ThrowsException<ArgumentException>(() => fixture.Service.ApplyProcedureRename(request), scenario);
                else Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.ApplyProcedureRename(request), scenario);
                Assert.AreEqual(0, fixture.Writes, scenario);
            }
        }

        /// <summary>Les lectures refusées n'écrivent rien; un échec après la première mutation révèle le résultat partiel.</summary>
        [TestMethod]
        public void ReadFailuresAndPartialWriteFailuresReportTheirActualBoundary()
        {
            foreach (string command in new[] { "project_properties", "list_modules", "read_module" })
            {
                var fixture = new ProcedureRenameWorkflowFixture { FailCommand = command };
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.PreviewProcedureRename(ProcedureRenameMatrix.Request()));
                Assert.AreEqual(0, fixture.Writes);
            }
            var partial = new ProcedureRenameWorkflowFixture(); var request = partial.PreviewRequest(); partial.FailedModule = "Caller";
            var failure = Assert.ThrowsException<InvalidOperationException>(() => partial.Service.ApplyProcedureRename(request));
            StringAssert.Contains(failure.Message, "1 module(s) were already verified");
            StringAssert.Contains(partial.Sources["MathModule"], "Function Compute"); Assert.AreEqual(ProcedureRenameMatrix.Caller, partial.Sources["Caller"]);
            var mismatch = new ProcedureRenameWorkflowFixture(); request = mismatch.PreviewRequest(); mismatch.WrongReadback = true;
            failure = Assert.ThrowsException<InvalidOperationException>(() => mismatch.Service.ApplyProcedureRename(request));
            StringAssert.Contains(failure.Message, "failing module may also have changed"); Assert.AreEqual(1, mismatch.Writes);
        }
    }
}
