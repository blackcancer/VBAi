using System;
using System.Linq;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    /// <summary>Matrice des références privées prouvées et des ambiguïtés refusées.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class VbaClassMemberRenameTests
    {
        [TestMethod]
        public void PrivateFunctionRenamesDeclarationReturnRecursionAndMeCalls()
        {
            string source = ClassMemberRenameFixture.Target.Replace("Calculate = x", "Calculate = Calculate(x - 1)") + "\r\n' Calculate remains comment\r\n";
            var plan = ClassMemberRenameFixture.Plan(source);
            Assert.AreEqual(1, plan.Edits.Count); Assert.AreEqual(5, plan.Edits[0].Replacements);
            StringAssert.Contains(plan.Edits[0].After, "Function Compute"); StringAssert.Contains(plan.Edits[0].After, "Compute = Compute(x - 1)");
            StringAssert.Contains(plan.Edits[0].After, "Me.Compute(2)"); StringAssert.Contains(plan.Edits[0].After, "' Calculate remains comment");
            var noop = ClassMemberRenameFixture.Request(source); noop.NewName = noop.Query;
            Assert.AreEqual(0, ClassMemberRenameFixture.Plan(source, noop).Edits.Count);
        }

        [TestMethod]
        public void PrivateSubRenamesDirectAndMeStatementCallsAndContinuations()
        {
            string source = "Option Explicit\r\nPrivate Sub Calculate()\r\nEnd Sub\r\nPublic Sub UseIt()\r\nCalculate\r\nCall Me.Calculate()\r\nIf True Then Calculate Else Me.Calculate\r\nCall _\r\n Calculate\r\nEnd Sub";
            var request = ClassMemberRenameFixture.Request(source); request.StartColumn = 13;
            var edit = ClassMemberRenameFixture.Plan(source, request).Edits.Single();
            Assert.AreEqual(6, edit.Replacements); StringAssert.Contains(edit.After, "If True Then Compute Else Me.Compute");
            StringAssert.Contains(edit.After, "Call _\r\n Compute");
        }

        [TestMethod]
        public void PrivatePropertyFamilyIsRenamedTogetherForEverySelectedAccessor()
        {
            string source = "Option Explicit\r\nPrivate Property Get Calculate() As Object\r\nSet Calculate = Nothing\r\nEnd Property\r\nPrivate Property Let Calculate(ByVal value As Object)\r\nEnd Property\r\nPrivate Property Set Calculate(ByVal value As Object)\r\nEnd Property\r\nPublic Sub UseIt()\r\nSet Me.Calculate = Nothing\r\nCalculate = Nothing\r\nEnd Sub";
            foreach (int kind in new[] { 3, 1, 2 })
            {
                var request = ClassMemberRenameFixture.Request(source); request.StartColumn = 22; request.ProcKind = kind;
                request.StartLine = kind == 3 ? 2 : kind == 1 ? 5 : 7;
                var edit = ClassMemberRenameFixture.Plan(source, request).Edits.Single(); Assert.AreEqual(6, edit.Replacements);
                Assert.IsFalse(edit.After.Contains("Calculate")); StringAssert.Contains(edit.After, "Property Set Compute");
            }
        }

        [TestMethod]
        public void LabelsNamedArgumentsLiteralsAndUnrelatedOtherComponentBindingsRemain()
        {
            string source = ClassMemberRenameFixture.Target.Replace("result = Calculate(1) + Me.Calculate(2)", "Calculate: result = Calculate(1)\r\nGoTo Calculate\r\nOther Calculate:=1\r\nDebug.Print \"Calculate\"");
            var edit = ClassMemberRenameFixture.Plan(source).Edits.Single();
            StringAssert.Contains(edit.After, "Calculate: result = Compute(1)"); StringAssert.Contains(edit.After, "GoTo Calculate");
            StringAssert.Contains(edit.After, "Other Calculate:=1"); StringAssert.Contains(edit.After, "\"Calculate\"");
            var other = "Option Explicit\r\nPublic Function Calculate() As Long\r\nCalculate = 1\r\nEnd Function";
            Assert.AreEqual(1, ClassMemberRenameFixture.Plan(ClassMemberRenameFixture.Target, null, other).Edits.Count);
        }

        [TestMethod]
        public void UnresolvedConstructsAndConsumersAreRefusedBeforeAnyEdit()
        {
            foreach (string construct in new[] { "Implements IFace", "Attribute Calculate.VB_UserMemId = 0", "#If VBA7 Then", "CallByName Me, \"Calculate\", VbMethod", "Application.Run \"Calculate\"", "With Me", "WithEvents sink As Object", "AddressOf Calculate", "DefInt A-Z", "Debug.Print [Calculate]", "Me.other.Calculate", "other.Calculate", ".Calculate", "other!Calculate", "Calculate = 1" })
            {
                string source = ClassMemberRenameFixture.Target.Replace("result = Calculate(1) + Me.Calculate(2)", construct);
                Assert.ThrowsException<InvalidOperationException>(() => ClassMemberRenameFixture.Plan(source), construct);
            }
            Assert.ThrowsException<InvalidOperationException>(() => ClassMemberRenameFixture.Plan(ClassMemberRenameFixture.Target, null,
                "Option Explicit\r\nPublic Sub UseIt()\r\nother.Calculate\r\nEnd Sub"));
        }

        [TestMethod]
        public void AmbiguousVisibilityDeclarationsAndSelectionAreRefused()
        {
            foreach (int scenario in Enumerable.Range(0, 12))
            {
                string source = ClassMemberRenameFixture.Target; var request = ClassMemberRenameFixture.Request();
                if (scenario == 0) source = source.Replace("Private Function", "Public Function");
                if (scenario == 1) source = source.Replace("Private Function", "Friend Function");
                if (scenario == 2) source = source.Replace("Private Function", "Function");
                if (scenario == 3) source = source.Replace("Option Explicit", "Option Compare Text");
                if (scenario == 4) source = source.Replace("Dim result", "Dim Calculate");
                if (scenario == 5) source = source.Replace("Dim result", "Dim Compute");
                if (scenario == 6) source += "\r\nPrivate Sub Compute()\r\nEnd Sub";
                if (scenario == 7) source += "\r\nPrivate Function Calculate()\r\nEnd Function";
                if (scenario == 8) source = source.Replace("End Function", "End Sub");
                if (scenario == 9) request.StartColumn++;
                if (scenario == 10) request.ProcKind = 3;
                if (scenario == 11) request.ExpectedSha256 = "stale";
                if (scenario != 11) request.ExpectedSha256 = VbaProcedureRename.Digest(source);
                Assert.ThrowsException<InvalidOperationException>(() => ClassMemberRenameFixture.Plan(source, request), "scenario " + scenario);
            }
        }

        [TestMethod]
        public void PropertyFamiliesRejectDuplicateKindsMixedVisibilityAndProcedureKinds()
        {
            string property = "Option Explicit\r\nPrivate Property Get Calculate() As Long\r\nCalculate = 1\r\nEnd Property";
            foreach (int scenario in Enumerable.Range(0, 4))
            {
                string source = property;
                if (scenario == 0) source += "\r\nPrivate Property Get Calculate() As Long\r\nEnd Property";
                if (scenario == 1) source += "\r\nPublic Property Let Calculate(ByVal value As Long)\r\nEnd Property";
                if (scenario == 2) source += "\r\nPrivate Sub Calculate()\r\nEnd Sub";
                if (scenario == 3) source = source.Replace("Property Get", "Property Unknown");
                var request = ClassMemberRenameFixture.Request(source); request.StartColumn = 22; request.ProcKind = 3;
                Assert.ThrowsException<InvalidOperationException>(() => ClassMemberRenameFixture.Plan(source, request));
            }
        }
        [TestMethod]
        public void InvalidCataloguesCallbacksAndReplacementNamesAreRefused()
        {
            var request = ClassMemberRenameFixture.Request();
            Assert.ThrowsException<ArgumentException>(() => VbaClassMemberRename.Prepare("P", null, request));
            Assert.ThrowsException<ArgumentException>(() => ClassMemberRenameFixture.Plan(ClassMemberRenameFixture.Target, new Request { Query = "Calculate", NewName = "Type" }));
            foreach (int scenario in Enumerable.Range(0, 6))
            {
                request = ClassMemberRenameFixture.Request();
                if (scenario == 0) request.Query = "Class_Initialize";
                if (scenario == 1) request.NewName = "Next_Value";
                if (scenario == 2) request.NewName = "Other";
                if (scenario == 3) request.NewName = "P";
                if (scenario == 4) request.Module = "absent";
                if (scenario == 5) request.NewName = "Main";
                Assert.ThrowsException<InvalidOperationException>(() => ClassMemberRenameFixture.Plan(ClassMemberRenameFixture.Target, request));
            }
            var module = new VbaProcedureRename.ModuleSnapshot("CalcClass", 1, ClassMemberRenameFixture.Target);
            Assert.ThrowsException<InvalidOperationException>(() => VbaClassMemberRename.Prepare("P", new[] { module }, ClassMemberRenameFixture.Request()));
            Assert.ThrowsException<InvalidOperationException>(() => VbaClassMemberRename.Prepare("P", new VbaProcedureRename.ModuleSnapshot[0], ClassMemberRenameFixture.Request()));
            Assert.ThrowsException<InvalidOperationException>(() => VbaClassMemberRename.Prepare("P", new[] { module, module }, ClassMemberRenameFixture.Request()));
        }
        [TestMethod]
        public void CompleteCatalogueNullAndBoundGuardsRejectBeforeAnalysis()
        {
            var request = ClassMemberRenameFixture.Request();
            var valid = new VbaProcedureRename.ModuleSnapshot("CalcClass", 2, ClassMemberRenameFixture.Target);
            Assert.ThrowsException<ArgumentException>(() => VbaClassMemberRename.Prepare("P", new[] { valid }, null));
            Assert.ThrowsException<ArgumentException>(() => VbaClassMemberRename.Prepare(" ", new[] { valid }, request));
            foreach (var invalid in new[] { new VbaProcedureRename.ModuleSnapshot[] { null },
                new[] { new VbaProcedureRename.ModuleSnapshot(" ", 2, "") },
                new[] { new VbaProcedureRename.ModuleSnapshot("CalcClass", 2, null) },
                Enumerable.Range(0, 1001).Select(i => new VbaProcedureRename.ModuleSnapshot("C" + i, 2, "")).ToArray() })
                Assert.ThrowsException<InvalidOperationException>(() => VbaClassMemberRename.Prepare("P", invalid, request));
            foreach (string name in new[] { "Main", "AutoOpen", "AutoClose", "P", "Other" })
            {
                request = ClassMemberRenameFixture.Request(); request.Query = name;
                Assert.ThrowsException<InvalidOperationException>(() => ClassMemberRenameFixture.Plan(ClassMemberRenameFixture.Target, request));
            }
            Assert.AreEqual(1, ClassMemberRenameFixture.Plan(ClassMemberRenameFixture.Target, null, "").Edits.Count);
        }

        [TestMethod]
        public void UnboundReplacementSuffixOutsideProcedureAndFunctionAssignmentsReject()
        {
            foreach (string body in new[] { "Debug.Print Compute", "Debug.Print Calculate%", "Me.Calculate = 1", "other.Me.Calculate", "Call Me!Calculate" })
            {
                string source = ClassMemberRenameFixture.Target.Replace("result = Calculate(1) + Me.Calculate(2)", body);
                Assert.ThrowsException<InvalidOperationException>(() => ClassMemberRenameFixture.Plan(source), body);
            }
            string outside = ClassMemberRenameFixture.Target + "\r\nCalculate";
            Assert.ThrowsException<InvalidOperationException>(() => ClassMemberRenameFixture.Plan(outside));
            foreach (string call in new[] { "result = Calculate", "Calculate = 1", "Me.Calculate = 1" })
            {
                string source = "Option Explicit\r\nPrivate Sub Calculate()\r\nEnd Sub\r\nPublic Sub UseIt()\r\n" + call + "\r\nEnd Sub";
                var request = ClassMemberRenameFixture.Request(source); request.StartColumn = 13;
                Assert.ThrowsException<InvalidOperationException>(() => ClassMemberRenameFixture.Plan(source, request), call);
            }
        }

        [TestMethod]
        public void ClosedMemberReaderRejectsUnmatchedNestedAndIncompleteSignatures()
        {
            foreach (string source in new[] { "Option Explicit\r\nEnd Function\r\n" + ClassMemberRenameFixture.Target,
                ClassMemberRenameFixture.Target.Replace("End Function", "Private Sub Nested()\r\nEnd Sub"),
                ClassMemberRenameFixture.Target.Replace("End Sub", ""),
                "Option Explicit\r\nPrivate Sub", "Option Explicit\r\nPrivate Property Get" })
                Assert.ThrowsException<InvalidOperationException>(() => ClassMemberRenameFixture.Plan(source));
            string modifiers = ClassMemberRenameFixture.Target.Replace("Private Function", "Private Static Function") + "\r\nPrivate Static\r\nDoEvents";
            var request = ClassMemberRenameFixture.Request(modifiers); request.StartColumn = 25;
            Assert.AreEqual(1, ClassMemberRenameFixture.Plan(modifiers, request).Edits.Count);
        }

        [TestMethod]
        public void ExcludedTypesTerminalLabelsWhitespaceAndUnrelatedTokensStayExact()
        {
            string source = ClassMemberRenameFixture.Target.Replace("result = Calculate(1) + Me.Calculate(2)",
                "GoSub Calculate\r\nResume Calculate\r\nDebug.Print New Calculate\r\nOther Calculate \t:=1\r\nDebug.Print Calculate: Debug.Print 2");
            var edit = ClassMemberRenameFixture.Plan(source, null, "Option Explicit\r\nCalculate\r\nDebug.Print Calculate:").Edits.Single();
            StringAssert.Contains(edit.After, "GoSub Calculate"); StringAssert.Contains(edit.After, "Resume Calculate");
            StringAssert.Contains(edit.After, "New Calculate"); StringAssert.Contains(edit.After, "Other Calculate \t:=1");
            StringAssert.Contains(edit.After, "Debug.Print Compute: Debug.Print 2");
            string label = ClassMemberRenameFixture.Target + "\r\nCalculate:";
            Assert.AreEqual(1, ClassMemberRenameFixture.Plan(label).Edits.Count);
            Assert.AreEqual(1, ClassMemberRenameFixture.Plan(source, null, "Option Explicit\r\nCalculate").Edits.Count);
        }
    }
}