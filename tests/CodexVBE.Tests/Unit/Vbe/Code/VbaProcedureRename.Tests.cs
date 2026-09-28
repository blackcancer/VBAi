using System;
using System.Linq;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    /// <summary>Vérifie le plan immutable du renommage standard sans moteur COM ni mutations.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class VbaProcedureRenameTests
    {
        /// <summary>Résout les appels du projet et conserve les homonymes membres, littéraux, labels et noms d'arguments.</summary>
        [TestMethod]
        public void ProjectFunctionCallersRecursionAndReturnValueRenameOnlyResolvedBindings()
        {
            var modules = ProcedureRenameMatrix.Project(); var original = modules[0].Source;
            var plan = VbaProcedureRename.Prepare("P", modules, ProcedureRenameMatrix.Request());
            Assert.AreEqual(2, plan.Edits.Count); Assert.AreEqual(original, modules[0].Source);
            var target = plan.Edits.Single(x => x.Module == "MathModule");
            StringAssert.Contains(target.After, "Public Function Compute(ByVal value As Long) As Long");
            StringAssert.Contains(target.After, "Compute = Compute(value - 1)");
            StringAssert.Contains(target.After, "Debug.Print \"Calc\" ' Calc");
            Assert.AreEqual(3, target.Replacements);
            var caller = plan.Edits.Single(x => x.Module == "Caller");
            StringAssert.Contains(caller.After, "Compute(1), MathModule.Compute(2), P.MathModule.Compute(3), obj.Calc(4)");
            StringAssert.Contains(caller.After, "Other Calc:=1"); StringAssert.Contains(caller.After, ".Calc(5)");
            StringAssert.Contains(caller.After, "GoTo Calc\r\nCalc: Debug.Print \"Calc\"");
            Assert.AreEqual(3, caller.Replacements);
        }

        /// <summary>Une procédure privée reste limitée au module; Sub et continuations préservent les colonnes physiques.</summary>
        [TestMethod]
        public void PrivateSubDirectAndQualifiedCallsAndCaseChangesAreSupported()
        {
            string source = "Option Explicit\nPrivate Sub Calc()\nCall Calc\nMathModule.Calc\nIf True Then Calc Else Calc\nCall P.MathModule. _\nCalc\nEnd Sub";
            var request = ProcedureRenameMatrix.Request(source);
            var plan = VbaProcedureRename.Prepare("P", new[] { new VbaProcedureRename.ModuleSnapshot("MathModule", 1, source),
                new VbaProcedureRename.ModuleSnapshot("EmptySheet", 100, "") }, request);
            StringAssert.Contains(plan.Edits[0].After, "Private Sub Compute()");
            StringAssert.Contains(plan.Edits[0].After, "Call P.MathModule. _\nCompute");
            Assert.AreEqual(6, plan.Edits[0].Replacements);
            request.NewName = "calc";
            Assert.IsTrue(VbaProcedureRename.Prepare("P", new[] { new VbaProcedureRename.ModuleSnapshot("MathModule", 1, source) }, request).Edits[0].After.Contains("Private Sub calc()"));
            request.NewName = "Calc";
            Assert.AreEqual(0, VbaProcedureRename.Prepare("P", new[] { new VbaProcedureRename.ModuleSnapshot("MathModule", 1, source) }, request).Edits.Count);
        }

        /// <summary>Chaque risque de résolution de la matrice refuse la préparation avant tout résultat éditable.</summary>
        [TestMethod]
        public void WholeProjectAmbiguityDynamicDispatchCallbacksAndCollisionsAreRejected()
        {
            foreach (string scenario in ProcedureRenameMatrix.Refusals)
            {
                string target = ProcedureRenameMatrix.Target, caller = ProcedureRenameMatrix.Caller; int type = 1;
                if (scenario == "class target") type = 2;
                if (scenario == "procedure collision") caller += "\r\nPublic Function Calc()\r\nEnd Function";
                if (scenario == "new procedure collision") caller += "\r\nPublic Sub Compute()\r\nEnd Sub";
                if (scenario == "local shadow") caller = caller.Replace("Public Sub Caller()", "Public Sub Caller(Calc As Long)");
                if (scenario == "qualifier shadow") caller += "\r\nDim MathModule As Object";
                if (scenario == "new implicit binding") caller = caller.Replace("Other Calc:=1", "Debug.Print Compute");
                if (scenario == "no explicit") caller = caller.Replace("Option Explicit", "");
                if (scenario == "conditional") caller += "\r\n#Const Flag = True";
                if (scenario == "dynamic call") caller += "\r\nSub Dispatch()\r\nApplication.Run \"Calc\"\r\nEnd Sub";
                if (scenario == "callback") caller += "\r\nSub Dispatch()\r\nOther AddressOf Calc\r\nEnd Sub";
                if (scenario == "interface") caller += "\r\nImplements IWorker";
                if (scenario == "bracket expression") caller = caller.Replace("Other Calc:=1", "Debug.Print [Calc]");
                if (scenario == "unterminated") target = target.Replace("End Function", "");
                if (scenario == "external qualification") caller = caller.Replace("P.MathModule.Calc", "OtherProject.MathModule.Calc");
                if (scenario == "private caller") target = target.Replace("Public Function", "Private Function");
                var request = ProcedureRenameMatrix.Request(target); var modules = ProcedureRenameMatrix.Project(target, caller, type);
                if (scenario == "missing target") request.Module = "Missing";
                if (scenario == "wrong position") request.StartColumn++;
                if (scenario == "stale sha") request.ExpectedSha256 = "stale";
                if (scenario == "property kind") request.ProcKind = 1;
                if (scenario == "duplicate module") modules = new[] { modules[0], modules[0] };
                if (scenario == "new callback") request.NewName = "Auto_Open";
                if (scenario == "reserved name") request.NewName = "Byte";
                if (scenario == "reserved name") Assert.ThrowsException<ArgumentException>(() => VbaProcedureRename.Prepare("P", modules, request), scenario);
                else Assert.ThrowsException<InvalidOperationException>(() => VbaProcedureRename.Prepare("P", modules, request), scenario);
            }
        }

        /// <summary>Les sources sans édition, types et identités font partie du SHA; l'ordre de collecte ne change pas la version.</summary>
        [TestMethod]
        public void CompleteSourceVersionIncludesEveryModuleAndComponentType()
        {
            var modules = ProcedureRenameMatrix.Project(); string version = VbaProcedureRename.Version("P", modules);
            Assert.AreEqual(version, VbaProcedureRename.Version("P", modules.Reverse()));
            Assert.AreNotEqual(version, VbaProcedureRename.Version("P", new[] { modules[0], new VbaProcedureRename.ModuleSnapshot("Caller", 2, modules[1].Source + " ' changed") }));
            Assert.AreNotEqual(version, VbaProcedureRename.Version("P", new[] { modules[0], new VbaProcedureRename.ModuleSnapshot("Caller", 1, modules[1].Source) }));
            Assert.AreNotEqual(version, VbaProcedureRename.Version("Other", modules));
            Assert.ThrowsException<ArgumentException>(() => VbaProcedureRename.Prepare("P", modules, null));
            Assert.ThrowsException<InvalidOperationException>(() => VbaProcedureRename.Prepare("P", new VbaProcedureRename.ModuleSnapshot[0], ProcedureRenameMatrix.Request()));
        }
    }
}
