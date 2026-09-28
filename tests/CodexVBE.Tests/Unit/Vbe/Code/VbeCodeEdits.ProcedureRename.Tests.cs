using System;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    /// <summary>Exerce la prévisualisation, la double vérification du projet et les mutations partielles déclarées.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class VbeProcedureRenameWorkflowTests
    {
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
