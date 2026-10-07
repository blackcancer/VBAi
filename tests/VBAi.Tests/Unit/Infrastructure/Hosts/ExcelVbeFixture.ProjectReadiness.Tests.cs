using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Readiness keeps the original generation, STA and explicit Main/private ownership distinct.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class ExcelVbeFixtureProjectReadinessTests
    {
        [DataTestMethod]
        [DataRow(false, 17, 17, 0, false, true, true)]
        [DataRow(true, 0, 0, 0, false, true, true)]
        [DataRow(true, 17, 18, 0, false, true, true)]
        [DataRow(true, 17, 17, 0, true, true, true)]
        [DataRow(true, 17, 17, 0, false, false, true)]
        [DataRow(true, 17, 17, 17, false, true, true)]
        [DataRow(true, 17, 17, 0, false, true, false)]
        [DataRow(true, 17, 17, 18, false, true, false)]
        public void InvalidOwnerRefusesBeforeAnyComRead(bool owned, int expected, int original,
            int child, bool exited, bool sta, bool main)
            => Assert.ThrowsException<InvalidOperationException>(() =>
                ExcelVbeFixture.RequireProjectReadinessOwner(owned, expected, original, child, exited, sta, main));

        [DataTestMethod]
        [DataRow(0, true)]
        [DataRow(17, false)]
        public void OriginalSelectedGenerationIsAccepted(int child, bool main)
            => ExcelVbeFixture.RequireProjectReadinessOwner(true, 17, 17, child, false, true, main);
    }
}
