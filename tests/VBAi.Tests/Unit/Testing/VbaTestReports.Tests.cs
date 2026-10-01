using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbaTestReportsTests
    {
        [TestMethod]
        public void HumanAndCompactReportsAgreeOnProjectRevisionOutcomesAndErrors()
        {
            var run = Run(VbaTestOutcome.Passed, VbaTestOutcome.Failed, VbaTestOutcome.Error,
                VbaTestOutcome.Inconclusive, VbaTestOutcome.Skipped, VbaTestOutcome.Blocked, VbaTestOutcome.Cancelled);
            run.Results[1].Message = "Expected \"five\"\r\nActual was four.";
            run.Results[1].ErrorNumber = -2147221504;
            var human = VbaTestReports.Human(run);
            var compact = Read(VbaTestReports.Compact(run));
            Assert.AreEqual(run.Id, compact["run"]);
            Assert.AreEqual(run.Project, compact["project"]);
            Assert.AreEqual(run.Revision, compact["revision"]);
            StringAssert.Contains(human, run.Id);
            StringAssert.Contains(human, run.Project);
            StringAssert.Contains(human, run.Revision);
            var counts = (Dictionary<string, object>)compact["counts"];
            foreach (var result in run.Results)
            {
                Assert.AreEqual(1, Convert.ToInt32(counts[result.Outcome.ToString()], CultureInfo.InvariantCulture));
                StringAssert.Contains(human, UiText.Get(result.Outcome.ToString()) + ": 1");
            }
            var tests = ((object[])compact["tests"]).Cast<Dictionary<string, object>>().ToArray();
            Assert.AreEqual(run.Results.Count, tests.Length);
            Assert.AreEqual(run.Results[1].Message, tests[1]["message"]);
            Assert.AreEqual(-2147221504, Convert.ToInt32(tests[1]["error"], CultureInfo.InvariantCulture));
            StringAssert.Contains(human, run.Results[1].Message);
            StringAssert.Contains(human, "-2147221504");
            for (int index = 0; index < tests.Length; index++)
            {
                Assert.AreEqual(run.Results[index].Test.Id, tests[index]["id"]);
                Assert.AreEqual(run.Results[index].Phase, tests[index]["phase"]);
                Assert.AreEqual(run.Results[index].Outcome.ToString(), tests[index]["outcome"]);
            }
        }

        [TestMethod]
        public void PassRateUsesVerifiedPassFailErrorDenominatorAndDoesNotInventCoverage()
        {
            var run = Run(VbaTestOutcome.Passed, VbaTestOutcome.Failed, VbaTestOutcome.Inconclusive,
                VbaTestOutcome.Skipped, VbaTestOutcome.Blocked, VbaTestOutcome.Cancelled, VbaTestOutcome.OutcomeUnknown);
            var human = VbaTestReports.Human(run);
            var compact = Read(VbaTestReports.Compact(run));
            StringAssert.Contains(human, "50% (1/2)");
            Assert.AreEqual(50d, Convert.ToDouble(compact["passRate"], CultureInfo.InvariantCulture));
            StringAssert.Contains(human, UiText.Get("VBA code coverage: unavailable"));
            var coverage = (Dictionary<string, object>)compact["coverage"];
            Assert.AreEqual(false, coverage["available"]);
            Assert.IsFalse(coverage.ContainsKey("percentage"));
            Assert.IsFalse(coverage.ContainsKey("percent"));
            Assert.IsFalse(coverage.ContainsKey("hits"));
        }

        [TestMethod]
        public void EmptyVerifiedDenominatorIsUnavailableRatherThanZeroOrHundredPercent()
        {
            foreach (var run in new[] { Run(), Run(VbaTestOutcome.Inconclusive, VbaTestOutcome.OutcomeUnknown, VbaTestOutcome.Cancelled) })
            {
                StringAssert.Contains(VbaTestReports.Human(run), "N/A (0/0)");
                Assert.IsNull(Read(VbaTestReports.Compact(run))["passRate"]);
            }
        }

        [TestMethod]
        public void UnknownAttemptAndModuleCleanupErrorRemainExplicitInBothFormats()
        {
            var run = Run(VbaTestOutcome.Passed, VbaTestOutcome.OutcomeUnknown);
            run.OutcomeUnknown = true;
            run.Error = "Module cleanup transport completion is unknown.";
            var human = VbaTestReports.Human(run);
            var compact = Read(VbaTestReports.Compact(run));
            Assert.AreEqual(true, compact["uncertain"]);
            Assert.AreEqual(run.Error, compact["error"]);
            StringAssert.Contains(human, run.Error);
            StringAssert.Contains(human, UiText.Get("Execution outcome is uncertain. Inspect the host before any further run."));
            var tests = (object[])compact["tests"];
            Assert.AreEqual("Passed", ((Dictionary<string, object>)tests[0])["outcome"]);
            Assert.AreEqual("OutcomeUnknown", ((Dictionary<string, object>)tests[1])["outcome"]);
        }

        [TestMethod]
        public void CompactNumericValuesRemainMachineReadableUnderFrenchCulture()
        {
            var original = CultureInfo.CurrentCulture;
            try
            {
                CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("fr-FR");
                var run = Run(VbaTestOutcome.Passed, VbaTestOutcome.Failed, VbaTestOutcome.Error);
                run.Results[0].Duration = TimeSpan.FromTicks(123456);
                var compact = Read(VbaTestReports.Compact(run));
                Assert.AreEqual(33.3333d, Convert.ToDouble(compact["passRate"], CultureInfo.InvariantCulture), 0.0000001d);
                var tests = (object[])compact["tests"];
                Assert.AreEqual(12.346d, Convert.ToDouble(((Dictionary<string, object>)tests[0])["ms"], CultureInfo.InvariantCulture), 0.0000001d);
                StringAssert.Contains(VbaTestReports.Human(run), "33,33% (1/3)");
            }
            finally { CultureInfo.CurrentCulture = original; }
        }

        [TestMethod]
        public void MissingRunHasAnExplicitEmptyReport()
        {
            Assert.AreEqual(UiText.Get("No test run results."), VbaTestReports.Human(null));
            var compact = Read(VbaTestReports.Compact(null));
            Assert.AreEqual(1, Convert.ToInt32(compact["v"], CultureInfo.InvariantCulture));
            Assert.IsNull(compact["run"]);
        }

        private static Dictionary<string, object> Read(string json) =>
            (Dictionary<string, object>)new JavaScriptSerializer().DeserializeObject(json);

        [TestMethod]
        public void LargeFailureMessagesAreRetrievedExactlyAcrossDisjointBoundedPagesWithGlobalSummaries()
        {
            var run = Run(Enumerable.Repeat(VbaTestOutcome.Failed, 1000).ToArray());
            run.OutcomeUnknown = true; run.Error = "Retained module cleanup diagnostic.";
            foreach (var result in run.Results) result.Message = new string('\u0001', 8192);
            var json = new JavaScriptSerializer { MaxJsonLength = 128 * 1024 * 1024 };
            string full = VbaTestReports.Compact(run);
            Assert.IsTrue(full.Length > 10 * 1024 * 1024, "Fixture must exceed the old report bound.");
            Assert.AreEqual(1000, ((object[])((Dictionary<string, object>)json.DeserializeObject(full))["tests"]).Length);
            var seen = new HashSet<string>();
            int? offset = 0;
            while (offset.HasValue)
            {
                string serialized = VbaTestReports.CompactPage(run, offset.Value);
                Assert.IsTrue(serialized.Length < 10 * 1024 * 1024);
                var page = (Dictionary<string, object>)json.DeserializeObject(serialized);
                Assert.AreEqual(1, page["v"]); Assert.AreEqual(1000, page["total"]);
                Assert.AreEqual(true, page["uncertain"]); Assert.AreEqual(run.Error, page["error"]);
                Assert.AreEqual(0d, Convert.ToDouble(page["passRate"]));
                Assert.AreEqual(1000, ((Dictionary<string, object>)page["counts"])["Failed"]);
                var rows = ((object[])page["tests"]).Cast<Dictionary<string, object>>().ToArray();
                Assert.AreEqual(100, rows.Length);
                foreach (var row in rows)
                {
                    Assert.IsTrue(seen.Add((string)row["id"]), "Pages must not overlap.");
                    Assert.AreEqual(new string('\u0001', 8192), row["message"], "Messages must never be truncated.");
                }
                offset = page["nextOffset"] == null ? (int?)null : Convert.ToInt32(page["nextOffset"]);
            }
            CollectionAssert.AreEquivalent(run.Results.Select(result => result.Test.Id).ToArray(), seen.ToArray());
            var tail = (Dictionary<string, object>)json.DeserializeObject(VbaTestReports.CompactPage(run, 1000, 1));
            Assert.AreEqual(0, ((object[])tail["tests"]).Length);
            Assert.IsNull(tail["nextOffset"]);
        }

        [TestMethod]
        public void CoveragePreviewAndMeasuredPagesMapAllSixteenThousandProbesWithUnchangedGlobalHeaders()
        {
            var plan = new VbaCoveragePlan { Original = "original", Revision = "revision", EligibleProcedureCount = 16000 };
            for (int index = 0; index < 16000; index++)
                plan.Probes.Add(new VbaCoverageProbe { Id = "p" + index, Index1Based = index + 1,
                    Module = "Production", Procedure = "Method" + index, OriginalLine = index + 1 });
            plan.Exclusions.Add(new VbaCoverageExclusion { Module = "Excluded", Reason = "Unsupported syntax" });
            var coverage = new VbaCoverageReport { Available = true, Complete = true, Original = plan.Original,
                Revision = plan.Revision, DenominatorKnown = true, Eligible = 16000, Hit = 8000, Percent = 50,
                Hits = plan.Probes.Select(probe => new VbaCoverageHit { Probe = probe, Entered = probe.Index1Based <= 8000 }).ToList(),
                Exclusions = plan.Exclusions };
            var json = new JavaScriptSerializer();
            var seen = new HashSet<string>();
            for (int offset = 0; offset < 16000; offset += 100)
            {
                var preview = (Dictionary<string, object>)json.DeserializeObject(json.Serialize(
                    VbeTestExplorerService.CoveragePreviewPage("project", "revision", plan, null, offset)));
                Assert.AreEqual(16000, preview["Total"]); Assert.AreEqual(16000, preview["ProbeTotal"]);
                Assert.AreEqual(16000, preview["EligibleProcedureCount"]); Assert.AreEqual(true, preview["Supported"]);
                Assert.AreEqual("revision", preview["ExpectedProjectVersion"]);
                var measured = (Dictionary<string, object>)json.DeserializeObject(json.Serialize(VbaTestReports.CoveragePage(coverage, offset)));
                Assert.AreEqual(16000, measured["Eligible"]); Assert.AreEqual(8000, measured["Hit"]);
                Assert.AreEqual(50, measured["Percent"]); Assert.AreEqual(true, measured["Complete"]);
                Assert.AreEqual(16000, measured["HitTotal"]); Assert.AreEqual(1, measured["ExclusionTotal"]);
                var probes = ((object[])preview["Probes"]).Cast<Dictionary<string, object>>().ToArray();
                var hits = ((object[])measured["Hits"]).Cast<Dictionary<string, object>>().ToArray();
                Assert.AreEqual(100, probes.Length); Assert.AreEqual(100, hits.Length);
                for (int index = 0; index < probes.Length; index++)
                {
                    Assert.IsTrue(seen.Add((string)probes[index]["Id"]));
                    Assert.AreEqual(probes[index]["Id"], ((Dictionary<string, object>)hits[index]["Probe"])["Id"]);
                    Assert.AreEqual(probes[index]["OriginalLine"], ((Dictionary<string, object>)hits[index]["Probe"])["OriginalLine"]);
                }
                Assert.AreEqual(offset == 15900 ? null : (object)(offset + 100), preview["NextOffset"]);
                Assert.AreEqual(preview["NextOffset"], measured["NextOffset"]);
            }
            Assert.AreEqual(16000, seen.Count);
            var run = Run(VbaTestOutcome.Passed); run.Coverage = coverage;
            var page = Read(VbaTestReports.CompactPage(run, 15900));
            Assert.AreEqual(0, ((object[])page["tests"]).Length);
            Assert.AreEqual(50, ((Dictionary<string, object>)page["coverage"])["percent"]);
            Assert.AreEqual(100, ((object[])((Dictionary<string, object>)page["coverage"])["probes"]).Length);
            plan.Probes.Add(new VbaCoverageProbe { Id = "over-capacity", Module = "Production", Procedure = "Extra" });
            plan.Diagnostics.Add("The project exceeds the bounded coverage probe capacity.");
            var oversized = Read(json.Serialize(VbeTestExplorerService.CoveragePreviewPage("project", "revision", plan, "Unsupported capacity", 16000)));
            Assert.AreEqual(16001, oversized["ProbeTotal"]); Assert.AreEqual(false, oversized["Supported"]);
            Assert.AreEqual(1, ((object[])oversized["Probes"]).Length);
            Assert.IsNull(oversized["NextOffset"]);
            var beyond = Read(json.Serialize(VbeTestExplorerService.CoveragePreviewPage("project", "revision", plan, null, int.MaxValue)));
            Assert.AreEqual(0, ((object[])beyond["Probes"]).Length); Assert.IsNull(beyond["NextOffset"]);
        }

        [TestMethod]
        public void ReadablePagesRetainGlobalSummaryWithoutTruncatingMessagesOrChangingLocalExport()
        {
            var run = Run(VbaTestOutcome.Passed, VbaTestOutcome.Failed, VbaTestOutcome.Error);
            run.Results[1].Message = "Exact full assertion message " + new string('x', 8192);
            run.OutcomeUnknown = true;
            string full = VbaTestReports.Human(run);
            string page = VbaTestReports.HumanPage(run, 1, 1);
            StringAssert.Contains(page, "tests total=3; nextOffset=2");
            StringAssert.Contains(page, (100d / 3).ToString("0.##", CultureInfo.CurrentCulture) + "% (1/3)");
            StringAssert.Contains(page, run.Results[1].Message);
            StringAssert.Contains(page, "TestsMath.Method1");
            Assert.IsFalse(page.Contains("TestsMath.Method0")); Assert.IsFalse(page.Contains("TestsMath.Method2"));
            StringAssert.Contains(page, UiText.Get("Execution outcome is uncertain. Inspect the host before any further run."));
            StringAssert.Contains(full, "TestsMath.Method0"); StringAssert.Contains(full, "TestsMath.Method2");
            Assert.AreEqual(full, VbaTestReports.Human(run));
            var beyond = Read(VbaTestReports.CompactPage(run, int.MaxValue, 100));
            Assert.IsNull(beyond["nextOffset"]); Assert.AreEqual(0, ((object[])beyond["tests"]).Length);
        }

        [TestMethod]
        public void InvalidPagingCannotRenderOrExposeReportCollections()
        {
            foreach (var pair in new[] { new[] { -1, 100 }, new[] { 0, -1 }, new[] { 0, 101 } })
                Assert.ThrowsException<ArgumentException>(() => VbaTestReports.CompactPage(Run(VbaTestOutcome.Passed), pair[0], pair[1]));
        }

        private static VbaTestRun Run(params VbaTestOutcome[] outcomes)
        {
            var run = new VbaTestRun { Id = "run1", Project = "WorkbookProject", Revision = "project-revision" };
            for (int index = 0; index < outcomes.Length; index++)
                run.Results.Add(new VbaTestResult
                {
                    Test = new VbaTestDescriptor { Id = "test" + index, Module = "TestsMath", Procedure = "Method" + index },
                    Outcome = outcomes[index], Phase = "Test", Duration = TimeSpan.FromMilliseconds(index + 1)
                });
            return run;
        }
    }
}
