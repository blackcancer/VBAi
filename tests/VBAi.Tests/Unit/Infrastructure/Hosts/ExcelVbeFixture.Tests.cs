using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit.Hosts
{
    [TestClass, TestCategory("Unit")]
    public sealed class ExcelVbeFixtureTests
    {
        [TestMethod]
        [DataRow(null)]
        [DataRow("")]
        [DataRow("relative-evidence")]
        [DataRow(@"C:relative-evidence")]
        [DataRow(@"\\server\share\evidence")]
        [DataRow(@"C:\evidence:stream")]
        public void ExplicitLaunchRejectsNonLocalOrMissingDurableEvidenceBeforeStartingExcel(string output)
        {
            Assert.ThrowsException<ArgumentException>(() => ExcelVbeFixture.ExplicitBootstrapTracePath(output));
        }

        [TestMethod]
        public void ExplicitLaunchPreparesDistinctTracePathsWithinSelectedEvidenceDirectory()
        {
            string parent = Path.Combine(Path.GetTempPath(), "VBAi-Fixture-Plan-" + Guid.NewGuid().ToString("N"));
            string first = ExcelVbeFixture.ExplicitBootstrapTracePath(parent);
            string second = ExcelVbeFixture.ExplicitBootstrapTracePath(parent);
            Assert.AreEqual(Path.GetFullPath(parent), Path.GetDirectoryName(first));
            Assert.AreEqual(".jsonl", Path.GetExtension(first));
            Assert.AreNotEqual(first, second);
            Assert.IsFalse(Directory.Exists(parent), "Planning must not create evidence or launch a host.");
        }
    }
}
