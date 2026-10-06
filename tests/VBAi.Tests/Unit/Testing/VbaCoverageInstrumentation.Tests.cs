using System;
using System.Linq;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbaCoverageInstrumentationTests
    {
        /// <summary>Checks that copied coverage support is MIT licensed without modifying the user's original module.</summary>
        [TestMethod]
        public void CoverageRuntimeCarriesMitNoticeAndPreservesUserCode()
        {
            const string source = "Public Sub Work()\nEnd Sub";
            var plan = VbaCoverageInstrumentation.Create(Project(Module("Production", 1, source)));
            Assert.IsTrue(plan.CanInstrument);
            StringAssert.Contains(plan.RuntimeSource, VbaRuntimeLicense.Comments);
            Assert.AreEqual(source, plan.Modules.Single().OriginalSource);
            Assert.IsFalse(plan.Modules.Single().InstrumentedSource.Contains("SPDX-License-Identifier"));
        }

        [TestMethod]
        public void InstrumentsAllProductionComponentKindsAndPropertyAccessorsButExcludesTestFramework()
        {
            var snapshot = Project(
                Module("Production", 1, "Public Sub Work()\nEnd Sub"),
                Module("Controller", 2, "Private Function Compute(ByVal value As Long) As Long\nCompute = value\nEnd Function\n"
                    + "Property Get Value() As Long\nValue = 1\nEnd Property\nProperty Let Value(ByVal nextValue As Long)\nEnd Property\n"
                    + "Property Set Model(ByVal nextModel As Object)\nEnd Property"),
                Module("Sheet1", 100, "Private Sub Worksheet_Change(ByVal Target As Object)\nEnd Sub"),
                Module("Editor", 3, "Private Sub UserForm_Initialize()\nEnd Sub"),
                Module("Tests", 1, "'@TestModule\n'@TestMethod\nPublic Sub VerifiesWork()\nEnd Sub"),
                Module(VbaTestRuntimeSource.ModuleName, 1, "Public Sub Framework()\nEnd Sub"));
            var plan = VbaCoverageInstrumentation.Create(snapshot);
            Assert.IsTrue(plan.CanInstrument, string.Join("\n", plan.Diagnostics));
            Assert.AreEqual(7, plan.EligibleProcedureCount);
            Assert.AreEqual(7, plan.Probes.Count);
            Assert.AreEqual(4, plan.Modules.Count);
            Assert.AreEqual(2, plan.Exclusions.Count);
            Assert.IsTrue(plan.Exclusions.All(exclusion => exclusion.Intentional));
            CollectionAssert.AreEquivalent(new[] { "Property Get", "Property Let", "Property Set" },
                plan.Probes.Where(probe => probe.Kind.StartsWith("Property", StringComparison.Ordinal)).Select(probe => probe.Kind).ToArray());
            Assert.IsTrue(plan.Probes.Any(probe => probe.Module == "Sheet1"));
            Assert.IsTrue(plan.Probes.Any(probe => probe.Module == "Editor"));
            Assert.IsFalse(plan.StatementCoverageAvailable);
        }

        [TestMethod]
        public void PlansAreDeterministicPreserveOriginalsAndMapProbesToOriginalPhysicalPositions()
        {
            const string original = "Option Explicit\r\nPublic Function Compute( _\r\n    ByVal value As Long) As Long ' header comment\r\n"
                + "    Compute = value\r\nEnd Function\r\n";
            var snapshot = Project(Module("Production", 1, original));
            var first = VbaCoverageInstrumentation.Create(snapshot);
            var second = VbaCoverageInstrumentation.Create(snapshot);
            Assert.AreEqual(original, snapshot.Modules[0].Source);
            Assert.AreEqual(original, first.Modules.Single().OriginalSource);
            Assert.AreEqual(second.Modules.Single().InstrumentedSource, first.Modules.Single().InstrumentedSource);
            Assert.AreEqual(second.Probes.Single().Id, first.Probes.Single().Id);
            Assert.AreEqual(2, first.Probes.Single().OriginalLine);
            Assert.AreEqual(17, first.Probes.Single().OriginalColumn);
            StringAssert.Contains(first.Modules.Single().InstrumentedSource,
                "ByVal value As Long) As Long ' header comment\r\n    VBAiCoverageSupport.VBAiProcedureCoverageHits(1) = True\r\n    Compute = value");
            Assert.IsFalse(first.Modules.Single().InstrumentedSource.Contains("\r\r\n"));
            snapshot.Modules[0].Source = "\r\n" + original;
            var shifted = VbaCoverageInstrumentation.Create(snapshot);
            Assert.AreEqual(first.Probes.Single().Id, shifted.Probes.Single().Id);
            Assert.AreEqual(3, shifted.Probes.Single().OriginalLine);
        }

        [TestMethod]
        public void InlineHeadersAreExplicitlyBlockedUntilHiddenAttributePreservationIsQualified()
        {
            const string source = "'Public Sub Fake()\nPublic Sub Inline(): Debug.Print \"x:y Public Sub Fake()\": End Sub\n"
                + "Public Function Text$(): Text$ = \"ok\": End Function\n"
                + "Public Sub WithDefault(Optional value As String = \"a:b\"): End Sub";
            var plan = VbaCoverageInstrumentation.Create(Project(Module("Production", 1, source)));
            Assert.IsFalse(plan.CanInstrument);
            Assert.IsTrue(plan.DenominatorKnown);
            Assert.AreEqual(3, plan.Probes.Count);
            Assert.AreEqual(3, plan.EligibleProcedureCount);
            Assert.AreEqual(source, plan.Modules.Single().InstrumentedSource);
            Assert.AreEqual(0, plan.Modules.Single().Edits.Count);
            Assert.AreEqual(3, plan.Exclusions.Count(exclusion => !exclusion.Intentional));
            Assert.IsTrue(plan.Diagnostics.All(message => message.Contains("hidden member attributes")));
            Assert.AreEqual("Text", plan.Probes[1].Procedure);
        }

        [TestMethod]
        public void ProbeDoesNotInsertCallsErrorHandlersClearsOrRewriteOriginalLabels()
        {
            const string source = "Public Function ObserveError() As Long\n10 On Error GoTo Failed\n20 ObserveError = Err.Number\n30 Exit Function\n"
                + "Failed:\n40 ObserveError = Erl\n50 Resume Next\nEnd Function";
            var plan = VbaCoverageInstrumentation.Create(Project(Module("Production", 1, source)));
            string rewritten = plan.Modules.Single().InstrumentedSource;
            Assert.AreEqual(source, rewritten.Replace("    VBAiCoverageSupport.VBAiProcedureCoverageHits(1) = True\n", ""));
            Assert.IsFalse(plan.RuntimeSource.Contains("On Error"));
            Assert.IsFalse(plan.RuntimeSource.Contains("Err.Clear"));
            Assert.IsFalse(rewritten.Contains("Call VBAiCoverageSupport"));
            StringAssert.Contains(plan.RuntimeSource, "Public VBAiProcedureCoverageHits(1 To 1) As Boolean");
            StringAssert.Contains(plan.RuntimeSource, "Public Function " + VbaCoverageInstrumentation.ResetProcedure + "(Optional ByVal ignoredHostArgument1 As Variant, Optional ByVal ignoredHostArgument2 As Variant) As Boolean");
            StringAssert.Contains(plan.RuntimeSource, "Public Function " + VbaCoverageInstrumentation.SnapshotProcedure + "(Optional ByVal ignoredHostArgument1 As Variant, Optional ByVal ignoredHostArgument2 As Variant) As Variant");
            Assert.IsFalse(plan.RuntimeSource.Contains("Function Reset("));
        }

        [DataTestMethod]
        [DataRow(0)]
        [DataRow(2)]
        public void CoverageRuntimeAcceptsOptionalWordArgumentsWithoutUsingThemInProbeState(int count)
        {
            string source = string.Join("\n", Enumerable.Range(0, count)
                .Select(index => "Public Sub Work" + index + "()\nEnd Sub"));
            var plan = VbaCoverageInstrumentation.Create(Project(Module("Production", 1, source)));
            Assert.IsTrue(plan.CanInstrument);
            Assert.AreEqual(count, plan.Probes.Count);
            const string arguments = "(Optional ByVal ignoredHostArgument1 As Variant, Optional ByVal ignoredHostArgument2 As Variant)";
            StringAssert.Contains(plan.RuntimeSource, "Public Function " + VbaCoverageInstrumentation.ResetProcedure + arguments + " As Boolean\r\n");
            StringAssert.Contains(plan.RuntimeSource, "Public Function " + VbaCoverageInstrumentation.SnapshotProcedure + arguments + " As Variant\r\n");
            // Each ignored name occurs only in the two helper declarations, never in their executable bodies.
            Assert.AreEqual(2, System.Text.RegularExpressions.Regex.Matches(plan.RuntimeSource, @"\bignoredHostArgument1\b").Count);
            Assert.AreEqual(2, System.Text.RegularExpressions.Regex.Matches(plan.RuntimeSource, @"\bignoredHostArgument2\b").Count);
            StringAssert.Contains(plan.RuntimeSource, "Public " + VbaCoverageInstrumentation.HitsVariable
                + "(1 To " + Math.Max(1, count) + ") As Boolean\r\n");
            var report = VbaCoverageInstrumentation.Read(plan, Native(new bool[Math.Max(1, count)]));
            Assert.AreEqual(0, report.Hit);
            Assert.AreEqual(count, report.Eligible);
        }

        [TestMethod]
        public void ConditionalProcedureDeclarationsBlockCoverageInsteadOfReducingDenominator()
        {
            var snapshot = Project(Module("Production", 1,
                "#If VBA7 Then\nPublic Sub Work()\nEnd Sub\n#Else\nPublic Sub Work()\nEnd Sub\n#End If"));
            var plan = VbaCoverageInstrumentation.Create(snapshot);
            Assert.IsFalse(plan.CanInstrument);
            Assert.IsFalse(plan.DenominatorKnown);
            Assert.AreEqual(2, plan.EligibleProcedureCount);
            Assert.IsTrue(plan.Exclusions.Any(exclusion => !exclusion.Intentional));
            var report = VbaCoverageInstrumentation.Unavailable(plan, "No clone was run.");
            Assert.IsFalse(report.Available);
            Assert.IsNull(report.Eligible);
            Assert.IsNull(report.Hit);
            Assert.IsNull(report.Percent);
            Assert.ThrowsException<InvalidOperationException>(() => VbaCoverageInstrumentation.Decode(plan, Native(false, false)));
        }

        [TestMethod]
        public void MinimalPhysicalEditsReproduceInstrumentedSourceWithoutReplacingDeclarations()
        {
            foreach (string newline in new[] { "\n", "\r\n" })
            {
                string source = string.Join(newline, new[] { "Option Explicit", "Public Function Continued( _",
                    "    ByVal value As Long) As Long ' comment", "    Continued = value", "End Function",
                    "Public Sub Regular()", "Debug.Print 1: Debug.Print 2", "End Sub", "Public Sub Last()", "End Sub", "" });
                var plan = VbaCoverageInstrumentation.Create(Project(Module("Production", 1, source)));
                Assert.IsTrue(plan.CanInstrument, string.Join("\n", plan.Diagnostics));
                var module = plan.Modules.Single();
                var lines = source.Replace("\r\n", "\n").Split('\n').ToList();
                foreach (var edit in module.Edits.OrderByDescending(edit => edit.OriginalLine).ThenByDescending(edit => edit.OriginalColumn))
                {
                    Assert.IsFalse(edit.Text.Contains("\n"));
                    if (edit.IsWholeLine)
                    {
                        Assert.AreEqual(1, edit.OriginalColumn);
                        lines.Insert(edit.OriginalLine - 1, edit.Text);
                    }
                    else lines[edit.OriginalLine - 1] = lines[edit.OriginalLine - 1].Insert(edit.OriginalColumn - 1, edit.Text);
                }
                Assert.AreEqual(module.InstrumentedSource, string.Join(newline, lines));
                Assert.AreEqual(3, module.Edits.Count);
                Assert.AreEqual(4, module.Edits.Single(edit => edit.IsWholeLine && edit.OriginalLine == 4).OriginalLine);
                Assert.IsTrue(module.Edits.All(edit => edit.IsWholeLine));
            }
        }

        [TestMethod]
        public void RuntimeProcedureCollisionsInExcludedTestsStillBlockCoverage()
        {
            foreach (string name in new[] { VbaCoverageInstrumentation.ResetProcedure, VbaCoverageInstrumentation.SnapshotProcedure, VbaCoverageInstrumentation.HitsVariable })
            {
                var plan = VbaCoverageInstrumentation.Create(Project(Module("Production", 1, "Sub Work()\nEnd Sub"),
                    Module("Tests", 1, "'@TestModule\nPublic Sub " + name + "()\nEnd Sub")));
                Assert.IsFalse(plan.CanInstrument, name);
                Assert.IsTrue(plan.DenominatorKnown);
                Assert.IsTrue(plan.Diagnostics.Any(diagnostic => diagnostic.Contains(name)));
            }
        }

        [TestMethod]
        public void ReservedRuntimeTokensBlockImplicitVariablesAndExternalReferencesButNotCommentsOrStrings()
        {
            foreach (string reserved in new[] { VbaCoverageInstrumentation.ModuleName, VbaCoverageInstrumentation.ResetProcedure,
                VbaCoverageInstrumentation.SnapshotProcedure, VbaCoverageInstrumentation.HitsVariable })
            {
                foreach (string body in new[] { reserved + " = 1", "Debug.Print ExternalLibrary." + reserved,
                    "Dim value As String\nvalue = " + reserved + "$", "Debug.Print [" + reserved + "()]" })
                {
                    var blocked = VbaCoverageInstrumentation.Create(Project(Module("Production", 1, "Sub Work()\n" + body + "\nEnd Sub")));
                    Assert.IsFalse(blocked.CanInstrument, body);
                    Assert.IsTrue(blocked.DenominatorKnown, body);
                    Assert.IsTrue(blocked.Diagnostics.Any(message => message.Contains(reserved)), body);
                }
                var safe = VbaCoverageInstrumentation.Create(Project(Module("Production", 1,
                    "' " + reserved + "\nSub Work()\nRem " + reserved + "\nDebug.Print \"" + reserved + "\"\nEnd Sub")));
                Assert.IsTrue(safe.CanInstrument, string.Join("\n", safe.Diagnostics));
            }
        }

        [TestMethod]
        public void ConditionalBodiesLeaveProcedureDenominatorKnownAndDeclarationsRemainUntouched()
        {
            const string source = "#If VBA7 Then\nPrivate Declare PtrSafe Sub External Lib \"x\" ()\n#End If\n"
                + "Public Sub Work()\n#If VBA7 Then\nDebug.Print 1\n#Else\nDebug.Print 2\n#End If\nEnd Sub";
            var plan = VbaCoverageInstrumentation.Create(Project(Module("Production", 1, source)));
            Assert.IsTrue(plan.CanInstrument, string.Join("\n", plan.Diagnostics));
            Assert.AreEqual(1, plan.Probes.Count);
            Assert.AreEqual("Work", plan.Probes.Single().Procedure);
            Assert.AreEqual(source, plan.Modules.Single().InstrumentedSource.Replace("    VBAiCoverageSupport.VBAiProcedureCoverageHits(1) = True\n", ""));
        }

        [TestMethod]
        public void AmbiguousStructuresAndShadowedSupportIdentifiersHaveExplicitBlockingDiagnostics()
        {
            foreach (var module in new[]
            {
                Module("Production", 1, "Public Sub Broken()\nDebug.Print 1"),
                Module("Production", 1, "Public Sub First()\nPublic Sub Second()\nEnd Sub\nEnd Sub"),
                Module("Production", 1, "Public Sub Duplicate()\nEnd Sub\nPublic Function Duplicate() As Long\nEnd Function"),
                Module("Production", 1, "Public Sub Broken(ByVal value As Long\nEnd Sub"),
                Module("Production", 1, "Public Sub Work(ByVal VBAiCoverageSupport As Object)\nEnd Sub"),
                Module("Production", 1, "Public Function VBAiCoverageSupport() As Long\nEnd Function"),
                Module("Production", 1, "Attribute VB_Name = \"Production\"\nPublic Sub Work()\nEnd Sub"),
                Module("UnknownType", 99, "Public Sub Work()\nEnd Sub"),
                Module(VbaCoverageInstrumentation.ModuleName, 1, "Public Sub Work()\nEnd Sub")
            })
            {
                var plan = VbaCoverageInstrumentation.Create(Project(module));
                Assert.IsFalse(plan.CanInstrument, module.Source);
                Assert.IsTrue(plan.Diagnostics.Count > 0, module.Source);
            }
            var duplicatedModule = VbaCoverageInstrumentation.Create(Project(Module("Same", 1, "Sub One()\nEnd Sub"), Module("SAME", 1, "Sub Two()\nEnd Sub")));
            Assert.IsFalse(duplicatedModule.CanInstrument);
            Assert.IsFalse(duplicatedModule.DenominatorKnown);
        }

        [TestMethod]
        public void DecodeMeasuresRealOneBasedBooleanHitsAndMapsEachResultToOriginalProcedure()
        {
            var plan = VbaCoverageInstrumentation.Create(Project(Module("Production", 1, "Public Sub One()\nEnd Sub\nPublic Sub Two()\nEnd Sub")));
            var report = VbaCoverageInstrumentation.Decode(plan, Native(true, false));
            Assert.IsTrue(report.Available);
            Assert.IsTrue(report.Complete);
            Assert.AreEqual(2, report.Eligible);
            Assert.AreEqual(1, report.Hit);
            Assert.AreEqual(50d, report.Percent.Value);
            Assert.AreEqual("Procedure", report.Metric);
            Assert.AreEqual("One", report.Hits.Single(hit => hit.Entered).Probe.Procedure);
            Assert.AreEqual("original-revision", report.Revision);
            Assert.AreEqual(@"C:\Temp\Original.xlsm", report.Original);
            var json = new JavaScriptSerializer().Serialize(report);
            StringAssert.Contains(json, "\"Metric\":\"Procedure\"");
            Assert.IsFalse(json.Contains("OriginalSource"));
            Assert.IsFalse(json.Contains("InstrumentedSource"));
        }

        [TestMethod]
        public void PartialExecutionHasMeasuredHitsButCannotClaimACompleteCoveragePercentage()
        {
            var plan = VbaCoverageInstrumentation.Create(Project(Module("Production", 1, "Sub One()\nEnd Sub")));
            var report = VbaCoverageInstrumentation.Decode(plan, Native(true), false);
            Assert.IsTrue(report.Available);
            Assert.IsFalse(report.Complete);
            Assert.AreEqual(1, report.Hit);
            Assert.IsNull(report.Percent);
            Assert.IsTrue(report.Diagnostics.Any(message => message.Contains("incomplete")));
            var unavailable = VbaCoverageInstrumentation.Unavailable(plan, "No valid snapshot returned.");
            Assert.IsFalse(unavailable.Available);
            Assert.AreEqual(1, unavailable.Eligible);
            Assert.IsNull(unavailable.Hit);
            Assert.IsNull(unavailable.Percent);
        }

        [TestMethod]
        public void DecoderRejectsWrongBoundsRanksLengthsAndCoercedNativeValues()
        {
            var plan = VbaCoverageInstrumentation.Create(Project(Module("Production", 1, "Sub One()\nEnd Sub")));
            var wrongElement = Array.CreateInstance(typeof(object), new[] { 1 }, new[] { 1 });
            wrongElement.SetValue(-1, 1);
            foreach (var invalid in new object[] { null, true, new[] { true }, new bool[1, 1], Native(true, false), wrongElement })
                Assert.ThrowsException<InvalidOperationException>(() => VbaCoverageInstrumentation.Decode(plan, invalid));
        }

        [TestMethod]
        public void InstrumentationDoesNotRewriteAnInlinePhysicalLineEvenWhenNearTheVbaLimit()
        {
            string source = "Public Sub Inline(): Debug.Print \"" + new string('x', 960) + "\": End Sub";
            Assert.IsTrue(source.Length <= 1023);
            var plan = VbaCoverageInstrumentation.Create(Project(Module("Production", 1, source)));
            Assert.IsFalse(plan.CanInstrument);
            Assert.IsTrue(plan.DenominatorKnown);
            Assert.IsTrue(plan.Diagnostics.Any(message => message.Contains("inline procedure header")));
            Assert.AreEqual(source, plan.Modules.Single().InstrumentedSource);
        }

        [TestMethod]
        public void EmptyProductionScopeIsValidButHasNoInventedZeroOrHundredPercent()
        {
            var plan = VbaCoverageInstrumentation.Create(Project(Module("Constants", 1, "Public Const Limit As Long = 4")));
            Assert.IsTrue(plan.CanInstrument);
            Assert.AreEqual(0, plan.EligibleProcedureCount);
            var report = VbaCoverageInstrumentation.Decode(plan, Native(false));
            Assert.AreEqual(0, report.Hit);
            Assert.AreEqual(0, report.Eligible);
            Assert.IsNull(report.Percent);
            Assert.ThrowsException<InvalidOperationException>(() => VbaCoverageInstrumentation.Decode(plan, Native(true)));
        }

        private static VbaTestProjectSnapshot Project(params VbaTestModuleSnapshot[] modules) =>
            new VbaTestProjectSnapshot { Id = "original-project", Selector = @"C:\Temp\Original.xlsm", Name = "Original",
                Revision = "original-revision", Modules = modules };
        private static VbaTestModuleSnapshot Module(string name, int type, string source) =>
            new VbaTestModuleSnapshot { Name = name, ComponentType = type, Source = source };
        private static Array Native(params bool[] values)
        {
            var result = Array.CreateInstance(typeof(bool), new[] { values.Length }, new[] { 1 });
            for (int index = 0; index < values.Length; index++) result.SetValue(values[index], index + 1);
            return result;
        }
        [TestMethod]
        public void MissingSnapshotsAndInvalidModuleNamesRefuseWithoutInventingAnEligibleDenominator()
        {
            Assert.ThrowsException<ArgumentNullException>(() => VbaCoverageInstrumentation.Create(null));
            Assert.ThrowsException<ArgumentNullException>(() => VbaCoverageInstrumentation.Read(null,null));
            Assert.ThrowsException<ArgumentNullException>(() => VbaCoverageInstrumentation.Unavailable(null,null));
            var project = Project(null, Module(null,1,null), Module(null,1,"")); project.Selector = null; project.Id = null;
            var plan = VbaCoverageInstrumentation.Create(project);
            Assert.AreEqual(project.Name,plan.Original); Assert.IsFalse(plan.DenominatorKnown);
            StringAssert.Contains(string.Join("\n",plan.Diagnostics),"missing");
            StringAssert.Contains(string.Join("\n",plan.Diagnostics),"module identifier");
            var empty = VbaCoverageInstrumentation.Create(new VbaTestProjectSnapshot { Modules = null });
            Assert.IsTrue(empty.CanInstrument); Assert.AreEqual(0,empty.Probes.Count);
            Assert.AreEqual(0,VbaCoverageInstrumentation.Unavailable(empty,null).Diagnostics.Count);
            plan = VbaCoverageInstrumentation.Create(Project(Module(VbaCoverageInstrumentation.ResetProcedure,1,"")));
            Assert.IsTrue(plan.Diagnostics.Any(message => message.Contains("module name conflicts")));
        }

        [DataTestMethod]
        [DataRow("#Else\n", "Unmatched conditional compilation branch")]
        [DataRow("#ElseIf True Then\n", "Unmatched conditional compilation branch")]
        [DataRow("#Unsupported\n", "Unsupported conditional compilation directive")]
        [DataRow("#If VBA7 Then\n", "Unterminated conditional compilation block")]
        [DataRow("Public Property Strange()\nEnd Property\n", "accessor is ambiguous")]
        [DataRow("Public Sub\nEnd Sub\n", "identifier is ambiguous")]
        [DataRow("Public Sub 3Bad()\nEnd Sub\n", "identifier is ambiguous")]
        [DataRow("Public Sub Good()", "header terminator")]
        public void UnresolvedSyntaxBlocksInstrumentationWithExplicitReason(string source,string reason)
        {
            var plan = VbaCoverageInstrumentation.Create(Project(Module("Production",1,source)));
            Assert.IsFalse(plan.CanInstrument); Assert.IsFalse(plan.DenominatorKnown);
            StringAssert.Contains(string.Join("\n",plan.Diagnostics),reason);
        }

        [TestMethod]
        public void LongPhysicalLinesAndUnclosedBracketNamesBlockOrPreserveTheirExactSource()
        {
            var source = "Public Sub Good()\nDebug.Print \"" + new string('x',1024) + "\"\nEnd Sub\n";
            var plan = VbaCoverageInstrumentation.Create(Project(Module("Production",1,source)));
            Assert.IsTrue(plan.Diagnostics.Any(message => message.Contains("physical line-length"))); Assert.IsTrue(plan.DenominatorKnown);
            plan = VbaCoverageInstrumentation.Create(Project(Module("Production",1,"Public Sub Good()\nDebug.Print [VBAiReadCoverageHits\nEnd Sub\n")));
            Assert.IsTrue(plan.Diagnostics.Any(message => message.Contains("reserved coverage runtime identifier")));
            plan = VbaCoverageInstrumentation.Create(Project(Module("Production",1,"Public Sub Good() ' trailing header comment\nEnd Sub\n")));
            Assert.IsTrue(plan.CanInstrument); Assert.AreEqual(1,plan.Probes.Count);
            Assert.AreEqual(2,plan.Modules.Single().Edits.Single().OriginalLine);
        }

        [TestMethod]
        public void ProbeCapacityRefusesOversizedProjectsWithoutDroppingSourceOrProcedures()
        {
            var source = new System.Text.StringBuilder();
            for (int index=0;index<=VbaCoverageInstrumentation.MaximumProbes;index++) source.Append("Sub P").Append(index).Append("()\nEnd Sub\n");
            var plan = VbaCoverageInstrumentation.Create(Project(Module("Production",1,source.ToString())));
            Assert.IsFalse(plan.CanInstrument); Assert.AreEqual(VbaCoverageInstrumentation.MaximumProbes+1,plan.Probes.Count);
            Assert.IsTrue(plan.Diagnostics.Any(message => message.Contains("probe capacity")));
        }
        [TestMethod]
        public void PhysicalOffsetsMapBothLineStartsAndInteriorPositionsWithoutChangingTheSource()
        {
            Assert.AreEqual(0,VbaCoverageInstrumentation.PhysicalLine(new[] { 0,10,20 },0));
            Assert.AreEqual(1,VbaCoverageInstrumentation.PhysicalLine(new[] { 0,10,20 },10));
            Assert.AreEqual(1,VbaCoverageInstrumentation.PhysicalLine(new[] { 0,10,20 },19));
            Assert.AreEqual(2,VbaCoverageInstrumentation.PhysicalLine(new[] { 0,10,20 },25));
        }

        [TestMethod]
        public void UnmatchedEndsIncompleteHeadersAndModifiersDoNotProduceQualifiedCoverage()
        {
            foreach(var source in new[] { "#End If\n", "#\n", "Public\n", "Public Sub Good() \"literal\"\n", "Public Sub Good() ' comment without newline" })
            {
                var plan=VbaCoverageInstrumentation.Create(Project(Module("Production",1,source)));
                if(source=="Public\n") Assert.AreEqual(0,plan.Probes.Count);
                else Assert.IsFalse(plan.CanInstrument,source);
            }
        }

        [TestMethod]
        public void ConservativeTokenMappingRefusesUnsafeTerminatorsAndMalformedLiteralPositions()
        {
            var flags=System.Reflection.BindingFlags.Static|System.Reflection.BindingFlags.NonPublic;
            var header=typeof(VbaCoverageInstrumentation).GetMethod("HeaderEndOffset",flags);
            var tokens=new System.Collections.Generic.List<VbaDeclarationIndex.Token> {new VbaDeclarationIndex.Token {Text="X",Line=1,Column=1}};
            Assert.AreEqual(-1,header.Invoke(null,new object[] { "X?",new[] {0},tokens }));
            Assert.AreEqual(-1,header.Invoke(null,new object[] { "X  ",new[] {0},tokens }));
            Assert.AreEqual(4,header.Invoke(null,new object[] { "X'c\n",new[] {0,4},tokens }));
            var identifier=typeof(VbaCoverageInstrumentation).GetMethod("IdentifierTokens",flags);
            var invalid=new VbaDeclarationIndex.Token {Text="<literal>",Line=1,Column=2};
            var mapped=(System.Collections.Generic.IEnumerable<VbaDeclarationIndex.Token>)identifier.Invoke(null,new object[] { invalid,new[] { "X" } });
            Assert.AreEqual(0,mapped.Count());
            var hash=typeof(VbaCoverageInstrumentation).GetMethod("Hash",flags);
            Assert.AreEqual(hash.Invoke(null,new object[] { "" }),hash.Invoke(null,new object[] { null }));
        }
    }
}
