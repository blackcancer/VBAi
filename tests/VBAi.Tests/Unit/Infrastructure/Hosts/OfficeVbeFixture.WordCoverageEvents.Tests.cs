using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class OfficeVbeFixtureWordCoverageEventsTests
    {
        [DataTestMethod, DataRow(null), DataRow("0"), DataRow("1")]
        public void EventScopeIsDisabledByDefaultAndCannotAdoptAnUnownedHost(string optIn)
        {
            const string variable = "VBAi_RUN_WORD_COVERAGE_EVENT_SCOPE";
            string prior = Environment.GetEnvironmentVariable(variable);
            try
            {
                Environment.SetEnvironmentVariable(variable, optIn);
                var fixture = (OfficeVbeFixture)Activator.CreateInstance(typeof(OfficeVbeFixture), true);
                if (optIn == "1") Assert.ThrowsException<AssertFailedException>(() => fixture.StartWordCoverageEventScope());
                else Assert.IsNull(fixture.StartWordCoverageEventScope());
                fixture.RestoreWordCoverageEventScope(null);
                Assert.IsFalse(fixture.NativeExecutionUnsettled);
            }
            finally { Environment.SetEnvironmentVariable(variable, prior); }
        }

        [TestMethod]
        public void FixtureNeverRestoresAnUnsettledCopyOrReadsAnotherNativeObject()
        {
            int writes = 0;
            var scope = new OfficeWordCoverageEventScope(new OfficeWordCoverageEventScope.Boundary
            {
                RequireOwner = () => { }, ReviewTemplates = () => { }, ReadSecurity = () => 3,
                AcquireCommands = () => new object(), ApplySuppression = (value, state) => writes++, Record = phase => { }, Release = value => { }
            });
            scope.Enter();
            var fixture = (OfficeVbeFixture)Activator.CreateInstance(typeof(OfficeVbeFixture), true);
            fixture.NativeExecutionUnsettled = true;
            fixture.RestoreWordCoverageEventScope(scope);
            Assert.AreEqual(1, writes); Assert.IsTrue(scope.Uncertain && fixture.NativeExecutionUnsettled);
        }
    }
}
