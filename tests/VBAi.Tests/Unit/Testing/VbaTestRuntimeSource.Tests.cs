using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbaTestRuntimeSourceTests
    {
        /// <summary>Reviewed dispatch source must be unchanged by the VBE's trailing-space normalization.</summary>
        [TestMethod]
        public void GeneratedSupportSurvivesNativeTrailingWhitespaceNormalization()
        {
            string source = VbaTestRuntimeSource.Generate(Catalog(Descriptor("TestsMath", "Good")));
            string observed = string.Join("\r\n", source.Split(new[] { "\r\n" }, StringSplitOptions.None)
                .Select(line => line.TrimEnd(' ', '\t')));
            Assert.AreEqual(source, observed);
            StringAssert.Contains(source, "\r\n'\r\n");
            StringAssert.Contains(source, "THE SOFTWARE IS PROVIDED");
        }
        /// <summary>Checks the MIT grant travels with dispatch source while legacy ownership headers remain recognized.</summary>
        [TestMethod]
        public void GeneratedUserProjectRuntimeCarriesMitNoticeWithoutChangingOwnershipProtocol()
        {
            string source = VbaTestRuntimeSource.Generate(Catalog(Descriptor("TestsMath", "Good")));
            StringAssert.Contains(source, VbaRuntimeLicense.Comments);
            Assert.IsTrue(VbaTestRuntimeSource.IsOwned(source));
            string beforeLicensing = source.Replace(VbaRuntimeLicense.Comments, "");
            Assert.IsTrue(VbaTestRuntimeSource.IsOwned(beforeLicensing));
            Assert.AreEqual("3", VbaTestRuntimeSource.Version);
        }

        [TestMethod]
        public void GenerateCreatesStableQualifiedDispatchAndBooleanFailure()
        {
            var alpha = Descriptor("TestsMath", "Add", "Function");
            var beta = Descriptor("TestsMath", "Subtract");
            var catalog = Catalog(beta, alpha);
            catalog.Modules[0].TestInitialize = Descriptor("TestsMath", "Setup");
            var source = VbaTestRuntimeSource.Generate(catalog);
            catalog.Modules[0].Tests.Reverse();
            Assert.AreEqual(source, VbaTestRuntimeSource.Generate(catalog));
            StringAssert.Contains(source, "Option Explicit\r\n' VBAi test support version " + VbaTestRuntimeSource.Version);
            StringAssert.Contains(source, "Case LCase$(\"TestsMath.Add\")\r\n            If Not TestsMath.Add() Then");
            StringAssert.Contains(source, "Call TestsMath.Setup");
            StringAssert.Contains(source, "Call TestsMath.Subtract");
            Assert.IsTrue(source.IndexOf("Case LCase$(\"TestsMath.Add\")", StringComparison.Ordinal)
                < source.IndexOf("Case LCase$(\"TestsMath.Subtract\")", StringComparison.Ordinal));
            Assert.IsTrue(source.Split('\n').All(line => line.Length <= 1023), "VBA physical lines must fit its limit.");
            Assert.IsFalse(source.Contains("Application.Run"));
            Assert.IsFalse(source.Contains("Public Function Run("), "The dispatcher must not shadow the host's global Run method.");
            StringAssert.Contains(source, "Public Function " + VbaTestRuntimeSource.DispatcherProcedure + "(");
            StringAssert.Contains(source, "Public Sub " + VbaTestRuntimeSource.PendingProcedure + "()");
        }

        [TestMethod]
        public void GenerateOmitsIgnoredBlockedUnsafeAndAmbiguousCalls()
        {
            var ignored = Descriptor("TestsMath", "Ignored");
            ignored.IgnoreReason = "Requires network";
            var blocked = Descriptor("TestsMath", "Blocked");
            blocked.Diagnostic = "Parameterized test";
            var source = VbaTestRuntimeSource.Generate(Catalog(
                Descriptor("TestsMath", "Good"), ignored, blocked,
                Descriptor("TestsMath", "bad\"\r\nKill"),
                Descriptor("TestsMath", "FinalNewline\n"),
                Descriptor("TestsMath", "Duplicate"), Descriptor("TestsMath", "DUPLICATE")));
            StringAssert.Contains(source, "Call TestsMath.Good");
            Assert.IsFalse(source.Contains("Call TestsMath.Ignored"));
            Assert.IsFalse(source.Contains("Call TestsMath.Blocked"));
            Assert.IsFalse(source.Contains("Kill"));
            Assert.IsFalse(source.Contains("FinalNewline"));
            Assert.IsFalse(source.Contains("Call TestsMath.Duplicate"));
            Assert.IsFalse(source.Contains("Call TestsMath.DUPLICATE"));
        }

        [TestMethod]
        public void GenerateOmitsConflictingModulesAndFrameworkModule()
        {
            var catalog = Catalog(Descriptor("TestsMath", "Good"));
            catalog.Modules.Add(new VbaTestModule { Name = "TESTSMATH" });
            catalog.Modules.Add(new VbaTestModule
            {
                Name = VbaTestRuntimeSource.ModuleName,
                Tests = { Descriptor(VbaTestRuntimeSource.ModuleName, "Run") }
            });
            catalog.Modules.Add(new VbaTestModule
            {
                Name = "BlockedModule",
                Diagnostic = "Invalid fixtures",
                Tests = { Descriptor("BlockedModule", "DoTest") }
            });
            var source = VbaTestRuntimeSource.Generate(catalog);
            Assert.IsFalse(source.Contains("Call TestsMath.Good"));
            Assert.IsFalse(source.Contains("Call VBAiTestSupport.Run"));
            Assert.IsFalse(source.Contains("Call BlockedModule.DoTest"));
            StringAssert.Contains(source, "Case Else");
        }

        [TestMethod]
        public void GenerateRequiresOwnedVersionedSupportBeforeReplacement()
        {
            var catalog = Catalog(Descriptor("TestsMath", "Good"));
            var source = VbaTestRuntimeSource.Generate(catalog);
            Assert.IsTrue(VbaTestRuntimeSource.IsOwned(source));
            Assert.IsTrue(VbaTestRuntimeSource.IsOwned(source.Replace("\r\n", "\n")));
            string marker = "support version " + VbaTestRuntimeSource.Version;
            string previous = source.Replace(marker, "support version 1");
            Assert.IsTrue(VbaTestRuntimeSource.IsOwned(previous), "The previous owned runtime can migrate to the current version.");
            Assert.IsTrue(VbaTestRuntimeSource.IsOwned(source.Replace(marker, "support version 2")));
            Assert.IsFalse(VbaTestRuntimeSource.IsOwned(source.Replace(marker, "support version 999")));
            Assert.IsFalse(VbaTestRuntimeSource.IsOwned("Option Explicit\n' VBAi test support version 1\n"));
            Assert.IsFalse(VbaTestRuntimeSource.IsOwned("'prefix\n" + source));
            Assert.IsFalse(VbaTestRuntimeSource.IsOwned(null));
            catalog.Project = new VbaTestProjectSnapshot
            {
                Modules = new[] { new VbaTestModuleSnapshot { Name = VbaTestRuntimeSource.ModuleName, Source = "Option Explicit" } }
            };
            Assert.ThrowsException<InvalidOperationException>(() => VbaTestRuntimeSource.Generate(catalog));
            catalog.Project.Modules[0].Source = source;
            Assert.AreEqual(source, VbaTestRuntimeSource.Generate(catalog));
            catalog.Project.Modules[0].Source = previous;
            Assert.AreEqual(source, VbaTestRuntimeSource.Generate(catalog), "Replacement must generate the current runtime, never preserve the legacy one.");
        }

        [TestMethod]
        public void GeneratePreservesUnicodeIdentifiersAndUsesHostCaseMappingForBothOperands()
        {
            var catalog = new VbaTestCatalog();
            catalog.Modules.Add(new VbaTestModule
            {
                Name = "Échangeur",
                Tests = { Descriptor("Échangeur", "Température") }
            });
            var source = VbaTestRuntimeSource.Generate(catalog);
            StringAssert.Contains(source, "Case LCase$(\"Échangeur.Température\")");
            StringAssert.Contains(source, "Call Échangeur.Température");
            StringAssert.Contains(source, "(LCase$(moduleName & \".\" & procedureName)) Then GoTo Executed");
            StringAssert.Contains(source, "Select Case key");
        }

        [TestMethod]
        public void LargeCataloguePartitionsAllCallsIntoBoundedLeavesAndRoutesWithoutDuplicateDispatch()
        {
            var tests = Enumerable.Range(0, 12000).Select(index => Descriptor("TestsMath", "Test" + index.ToString("D5"),
                index % 2 == 0 ? "Sub" : "Function")).ToArray();
            var catalog = Catalog(tests);
            catalog.Modules[0].ModuleInitialize = Descriptor("TestsMath", "InitializeModule");
            catalog.Modules[0].ModuleCleanup = Descriptor("TestsMath", "CleanupModule");
            catalog.Modules[0].TestInitialize = Descriptor("TestsMath", "InitializeTest");
            catalog.Modules[0].TestCleanup = Descriptor("TestsMath", "CleanupTest");
            string signature = VbaTestRuntimeSource.DispatchSignature(catalog);
            string source = VbaTestRuntimeSource.Generate(catalog);
            Assert.AreEqual(signature, VbaTestRuntimeSource.DispatchSignature(catalog));
            catalog.Modules[0].Tests.Reverse();
            Assert.AreEqual(source, VbaTestRuntimeSource.Generate(catalog), "Partition layout must be deterministic.");
            var helpers = DispatchHelpers(source);
            var leaves = helpers.Where(pair => pair.Key.StartsWith("VBAiDispatchLeaf", StringComparison.Ordinal)).ToArray();
            var routes = helpers.Where(pair => pair.Key.StartsWith("VBAiDispatchRoute", StringComparison.Ordinal)).ToArray();
            Assert.IsTrue(leaves.Length > VbaTestRuntimeSource.MaximumRouteChildren);
            Assert.IsTrue(routes.Length > 1, "Beyond one routing level the root must not grow with catalogue size.");
            var actualCalls = new List<string>();
            foreach (var leaf in leaves)
            {
                var cases = Regex.Matches(leaf.Value, @"Case LCase\$\(""([^""]+)""\)").Cast<Match>().ToArray();
                Assert.IsTrue(cases.Length > 0 && cases.Length <= VbaTestRuntimeSource.MaximumCasesPerLeaf);
                Assert.IsTrue(leaf.Value.Length <= VbaTestRuntimeSource.MaximumLeafBodyCharacters + 256,
                    "Leaf body must remain bounded even for verbose Boolean branches.");
                actualCalls.AddRange(cases.Select(match => match.Groups[1].Value));
                StringAssert.Contains(leaf.Value, "Case Else\r\n            Exit Function");
                Assert.IsTrue(leaf.Value.LastIndexOf(leaf.Key + " = True", StringComparison.Ordinal)
                    > leaf.Value.IndexOf("End Select", StringComparison.Ordinal), "Handled must not mean Boolean test passed.");
                Assert.IsFalse(leaf.Value.Contains("On Error"), "Native errors must propagate to the central handler.");
            }
            var expected = tests.Concat(new[] { catalog.Modules[0].ModuleInitialize, catalog.Modules[0].ModuleCleanup,
                catalog.Modules[0].TestInitialize, catalog.Modules[0].TestCleanup }).Select(test => test.Module + "." + test.Procedure).ToArray();
            CollectionAssert.AreEquivalent(expected, actualCalls.ToArray());
            Assert.AreEqual(expected.Length, actualCalls.Distinct(StringComparer.OrdinalIgnoreCase).Count());
            var incoming = new List<string>();
            foreach (var route in routes)
            {
                var calls = Regex.Matches(route.Value, @"If (VBAiDispatch(?:Leaf|Route)\d+)\(key\) Then").Cast<Match>().ToArray();
                Assert.IsTrue(calls.Length > 0 && calls.Length <= VbaTestRuntimeSource.MaximumRouteChildren);
                Assert.IsFalse(route.Value.Contains(" Or "));
                foreach (var call in calls)
                {
                    string child = call.Groups[1].Value;
                    Assert.IsTrue(helpers.ContainsKey(child));
                    incoming.Add(child);
                    StringAssert.Contains(route.Value, "If " + child + "(key) Then\r\n        " + route.Key + " = True\r\n        Exit Function\r\n    End If");
                }
            }
            string root = Regex.Match(source, @"If (VBAiDispatch(?:Leaf|Route)\d+)\(LCase\$").Groups[1].Value;
            Assert.IsTrue(helpers.ContainsKey(root));
            string central = Regex.Match(source, @"^Public Function VBAiExecuteTest\(.*?^End Function\r?$",
                RegexOptions.Multiline | RegexOptions.Singleline).Value;
            Assert.IsTrue(central.Length > 0 && central.Length < 8192, "The public error/assertion envelope must remain bounded.");
            Assert.AreEqual(1, Regex.Matches(central, @"If VBAiDispatch(?:Leaf|Route)\d+\(").Count);
            Assert.IsFalse(central.Contains("Case LCase$"), "Catalogue cases must never accumulate in the central procedure.");
            Assert.IsFalse(incoming.Contains(root));
            CollectionAssert.AreEquivalent(helpers.Keys.Where(name => name != root).ToArray(), incoming.ToArray());
            Assert.AreEqual(incoming.Count, incoming.Distinct(StringComparer.Ordinal).Count(), "Each subtree must be visited through one parent only.");
            StringAssert.Contains(source, "On Error GoTo FailedCall");
            StringAssert.Contains(source, "This procedure is absent from the installed test dispatch table.");
        }

        [TestMethod]
        public void LongIdentifiersUseSourceBudgetBeforeCaseCountAndUnknownCatalogueCannotInvokeAnyTest()
        {
            string moduleName = "M" + new string('m', 254);
            var module = new VbaTestModule { Name = moduleName };
            for (int index = 0; index < 128; index++)
                module.Tests.Add(Descriptor(moduleName, "T" + index.ToString("D3") + new string('t', 251), "Function"));
            var catalog = new VbaTestCatalog(); catalog.Modules.Add(module);
            string source = VbaTestRuntimeSource.Generate(catalog);
            var leaves = DispatchHelpers(source).Where(pair => pair.Key.StartsWith("VBAiDispatchLeaf", StringComparison.Ordinal)).ToArray();
            Assert.IsTrue(leaves.Length > 1, "Long identifiers must be partitioned even below the case-count bound.");
            foreach (var leaf in leaves)
                Assert.IsTrue(leaf.Value.Length <= VbaTestRuntimeSource.MaximumLeafBodyCharacters + 256);
            Assert.IsTrue(source.Split('\n').All(line => line.Length <= 1023));
            var empty = VbaTestRuntimeSource.Generate(new VbaTestCatalog());
            var emptyHelpers = DispatchHelpers(empty);
            Assert.AreEqual(1, emptyHelpers.Count);
            StringAssert.Contains(emptyHelpers.Single().Value, "Case Else\r\n            Exit Function");
            Assert.AreEqual(0, Regex.Matches(empty, @"Case LCase\$").Count);
            StringAssert.Contains(empty, "VBAiExecuteTest = Array(\"Error\", \"This procedure is absent from the installed test dispatch table.\", \"0\")");
        }

        private static Dictionary<string, string> DispatchHelpers(string source)
        {
            return Regex.Matches(source, @"^Private Function (VBAiDispatch(?:Leaf|Route)\d+)\(ByVal key As String\) As Boolean\r\n(.*?)^End Function\r?$",
                RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.CultureInvariant).Cast<Match>()
                .ToDictionary(match => match.Groups[1].Value, match => match.Groups[2].Value, StringComparer.Ordinal);
        }

        [TestMethod]
        public void RuntimePreservesStickyAssertionsAndCapturesErrorsBeforeClear()
        {
            var source = VbaTestRuntimeSource.Generate(Catalog(Descriptor("TestsMath", "Good")));
            StringAssert.Contains(source, "mFailed = False");
            StringAssert.Contains(source, "If Not mFailed Then mMessage = Left$(message, 8192)");
            StringAssert.Contains(source, "If Not mFailed And Not mInconclusive Then");
            StringAssert.Contains(source, "If mFailed Then\r\n        outcome = \"Failed\"");
            Assert.IsTrue(source.IndexOf("errorNumber = Err.Number", StringComparison.Ordinal)
                < source.IndexOf("Err.Clear", StringComparison.Ordinal));
            StringAssert.Contains(source, VbaTestRuntimeSource.DispatcherProcedure + " = Array(outcome, Left$(mMessage, 8192), CStr(errorNumber))");
            StringAssert.Contains(source, "StrComp(expected, actual, vbBinaryCompare)");
            StringAssert.Contains(source, "AreEqual requires identical scalar types");
            StringAssert.Contains(source, "AreEqual does not support objects");
            StringAssert.Contains(source, "AreEqual does not support arrays");
            StringAssert.Contains(source, "Test is not implemented.");
        }

        [TestMethod]
        public void DecodeAcceptsNativeSafeArrayWithNonzeroLowerBound()
        {
            var native = Array.CreateInstance(typeof(object), new[] { 3 }, new[] { 1 });
            native.SetValue("Failed", 1);
            native.SetValue("Expected 5", 2);
            native.SetValue("-2147221504", 3);
            var test = Descriptor("TestsMath", "Add");
            var result = VbaTestRuntimeSource.Decode(test, native);
            Assert.AreSame(test, result.Test);
            Assert.AreEqual(VbaTestOutcome.Failed, result.Outcome);
            Assert.AreEqual("Expected 5", result.Message);
            Assert.AreEqual(-2147221504, result.ErrorNumber);
        }

        [TestMethod]
        public void DecodeAcceptsOnlyVerifiedCompletionOutcomes()
        {
            foreach (var outcome in new[] { "Passed", "Failed", "Error", "Inconclusive" })
                Assert.AreEqual(outcome, VbaTestRuntimeSource.Decode(Descriptor("TestsMath", "Add"),
                    new object[] { outcome, "Message", "0" }).Outcome.ToString());
            foreach (var outcome in new[] { "NotRun", "OutcomeUnknown", "Skipped", "passed", "1", "Completed", "" })
                Assert.ThrowsException<InvalidOperationException>(() => VbaTestRuntimeSource.Decode(
                    Descriptor("TestsMath", "Add"), new object[] { outcome, "Message", "0" }));
        }

        [TestMethod]
        public void DecodeRejectsMalformedUnboundedAndInconsistentEnvelopes()
        {
            foreach (var native in new object[]
            {
                null, true, new object[] { "Passed", "" }, new object[1, 3],
                new object[] { "Passed", "", 0 }, new object[] { "Passed", null, "0" },
                new object[] { "Passed", new string('x', 8193), "0" },
                new object[] { "Error", "", "2147483648" },
                new object[] { "Error", "", " 3" },
                new object[] { "Passed", "", "5" },
                new object[] { "Inconclusive", "", "-1" }
            })
                Assert.ThrowsException<InvalidOperationException>(() => VbaTestRuntimeSource.Decode(Descriptor("TestsMath", "Add"), native));
        }

        private static VbaTestDescriptor Descriptor(string module, string procedure, string kind = "Sub") =>
            new VbaTestDescriptor { Module = module, Procedure = procedure, Kind = kind, Id = module + "." + procedure };

        private static VbaTestCatalog Catalog(params VbaTestDescriptor[] tests)
        {
            var catalog = new VbaTestCatalog();
            catalog.Modules.Add(new VbaTestModule { Name = "TestsMath", Tests = tests.ToList() });
            return catalog;
        }
        [TestMethod]
        public void NullableCatalogueCollectionsAndDescriptorsCannotProduceUnsafeDispatch()
        {
            Assert.ThrowsException<ArgumentNullException>(() => VbaTestRuntimeSource.Generate(null));
            var catalog = Catalog(Descriptor("Foreign", "Foreign"), Descriptor("TestsMath", null), Descriptor("TestsMath", "WrongKind", "Property"));
            catalog.Modules[0].Tests.Add(null);
            catalog.Modules.Add(new VbaTestModule { Name = null });
            string source = VbaTestRuntimeSource.Generate(catalog);
            Assert.IsFalse(source.Contains("Case LCase"));
            catalog.Modules[0].Tests = null;
            source = VbaTestRuntimeSource.Generate(catalog);
            Assert.IsFalse(source.Contains("Case LCase"));
            catalog.Modules = null;
            Assert.IsFalse(VbaTestRuntimeSource.Generate(catalog).Contains("Case LCase"));
        }

        [TestMethod]
        public void DuplicateExistingSupportModulesRefuseEvenWhenBothAreOwned()
        {
            var catalog = Catalog(Descriptor("TestsMath", "Good"));
            string source = VbaTestRuntimeSource.Generate(catalog);
            catalog.Project = new VbaTestProjectSnapshot
            {
                Modules = new[] {
                new VbaTestModuleSnapshot { Name = VbaTestRuntimeSource.ModuleName, Source = source },
                new VbaTestModuleSnapshot { Name = VbaTestRuntimeSource.ModuleName, Source = source } }
            };
            Assert.ThrowsException<InvalidOperationException>(() => VbaTestRuntimeSource.Generate(catalog));
        }

        [TestMethod]
        public void DecoderRejectsNullTestAndExcessiveStatusOrErrorText()
        {
            Assert.ThrowsException<ArgumentNullException>(() => VbaTestRuntimeSource.Decode(null, null));
            foreach (var envelope in new[] { new object[] { null, "", "0" }, new object[] { new string('x', 33), "", "0" }, new object[] { "Error", "", new string('1', 13) } })
                Assert.ThrowsException<InvalidOperationException>(() => VbaTestRuntimeSource.Decode(Descriptor("TestsMath", "Good"), envelope));
        }
        [TestMethod]
        public void ShortSubCasesPartitionAtTheExactCaseCountBoundBeforeTheSourceBudget()
        {
            var catalog = Catalog(Enumerable.Range(0, 129).Select(index => Descriptor("TestsMath", "T" + index)).ToArray());
            var leaves = DispatchHelpers(VbaTestRuntimeSource.Generate(catalog)).Where(pair => pair.Key.StartsWith("VBAiDispatchLeaf", StringComparison.Ordinal)).OrderBy(pair => pair.Key).ToArray();
            Assert.AreEqual(2, leaves.Length);
            Assert.AreEqual(128, Regex.Matches(leaves[0].Value, @"Case LCase\$").Count);
            Assert.AreEqual(1, Regex.Matches(leaves[1].Value, @"Case LCase\$").Count);
            Assert.AreEqual(129, leaves.Sum(leaf => Regex.Matches(leaf.Value, @"Case LCase\$").Count));
        }
    }
}
