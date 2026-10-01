using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbaTestModelsTests
    {
        [TestMethod]
        public void CatalogEnumeratesOnlyTestsAndKeepsFixturesOutsideTheExecutionSelection()
        {
            var test = new VbaTestDescriptor { Procedure = "Check" };
            var fixture = new VbaTestDescriptor { Procedure = "Setup" };
            var catalog = new VbaTestCatalog();
            catalog.Modules.Add(new VbaTestModule { Name = "Tests", ModuleInitialize = fixture });
            catalog.Modules[0].Tests.Add(test);
            CollectionAssert.AreEqual(new[] { test }, catalog.Tests.ToArray());
            Assert.AreEqual(0, new VbaTestProjectSnapshot().Modules.Length);
            Assert.AreEqual(0, new VbaTestDescriptor().Categories.Length);
            Assert.AreEqual(VbaTestOutcome.NotRun, new VbaTestResult().Outcome);
        }

        [TestMethod]
        public void HistoricalRunRetainsRevisionDurationAndUncertainErrorEvidence()
        {
            var run = new VbaTestRun { Id = "run", Revision = "previous-source", OutcomeUnknown = true, Error = "Host closed" };
            run.Results.Add(new VbaTestResult
            {
                Test = new VbaTestDescriptor { Id = "test" }, Outcome = VbaTestOutcome.OutcomeUnknown,
                Message = "Completion was not verified", Phase = "Test", ErrorNumber = 5, Duration = TimeSpan.FromSeconds(2)
            });
            Assert.AreEqual("previous-source", run.Revision);
            Assert.AreEqual(VbaTestOutcome.OutcomeUnknown, run.Results.Single().Outcome);
            Assert.AreEqual(TimeSpan.FromSeconds(2), run.Results.Single().Duration);
            Assert.AreEqual(5, run.Results.Single().ErrorNumber);
            Assert.IsTrue(run.OutcomeUnknown);
        }
    }
}
