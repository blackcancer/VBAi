using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Reflection;
using System.Web.Script.Serialization;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks consent admission without an Office process, UI enumeration or document activation.</summary>
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class OfficeVbeFixtureCoverageConsentTests
    {
        [DataTestMethod]
        [DataRow("coverage-copy-consent", "intent")]
        [DataRow("coverage-copy-consent", "returned")]
        [DataRow("word-coverage-events", "enable-intent")]
        [DataRow("word-coverage-events", "enable-returned")]
        [DataRow("word-coverage-events", "enable-checked")]
        [DataRow("word-coverage-events", "restore-intent")]
        [DataRow("word-coverage-events", "restore-returned")]
        [DataRow("word-coverage-events", "restore-released")]
        public void IndependentCopyScopesPublishTheSamePhaseWithoutReplacingEarlierEvidence(string prefix, string phase)
        {
            string root = Path.Combine(Path.GetTempPath(), "VBAi-CoverageReceiptTests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try
            {
                string first = Path.Combine(root, prefix + "-" + phase + ".json");
                string second = Path.Combine(root, prefix + "-2-" + phase + ".json");
                string third = Path.Combine(root, prefix + "-3-" + phase + ".json");
                OfficeVbeFixture.RecordCoverageReceipt(root, prefix, 1, phase, new { Cycle = 1, Original = "retained" });
                byte[] baseline = File.ReadAllBytes(first);
                OfficeVbeFixture.RecordCoverageReceipt(root, prefix, 2, phase, new { Cycle = 2 });
                OfficeVbeFixture.RecordCoverageReceipt(root, prefix, 3, phase, new { Cycle = 3 });
                var serializer = new JavaScriptSerializer();
                Assert.AreEqual(1, serializer.Deserialize<System.Collections.Generic.Dictionary<string, object>>(File.ReadAllText(first))["Cycle"]);
                Assert.AreEqual(2, serializer.Deserialize<System.Collections.Generic.Dictionary<string, object>>(File.ReadAllText(second))["Cycle"]);
                Assert.AreEqual(3, serializer.Deserialize<System.Collections.Generic.Dictionary<string, object>>(File.ReadAllText(third))["Cycle"]);
                Assert.AreEqual(3, Directory.GetFiles(root).Length);
                Assert.ThrowsException<IOException>(() => OfficeVbeFixture.RecordCoverageReceipt(root, prefix, 1, phase, new { Cycle = 99 }));
                CollectionAssert.AreEqual(baseline, File.ReadAllBytes(first), "An old receipt must remain immutable.");
            }
            finally { Directory.Delete(root, true); }
        }

        [DataTestMethod]
        [DataRow("coverage-copy-consent", null)]
        [DataRow("coverage-copy-consent", "")]
        [DataRow("coverage-copy-consent", "restore-intent")]
        [DataRow("coverage-copy-consent", "../intent")]
        [DataRow("coverage-copy-consent", "intent/other")]
        [DataRow("coverage-copy-consent", "Intent")]
        [DataRow("word-coverage-events", "intent")]
        [DataRow("word-coverage-events", "restore")]
        [DataRow("word-coverage-events", "../enable-intent")]
        [DataRow(null, "intent")]
        [DataRow("", "intent")]
        [DataRow("other", "intent")]
        [DataRow("../coverage-copy-consent", "intent")]
        public void UnknownOrPathBearingReceiptNamesAreRejectedBeforeFileCreation(string prefix, string phase)
            => Assert.ThrowsException<ArgumentException>(() => OfficeVbeFixture.CoverageReceiptPath(Path.GetTempPath(), prefix, 2, phase));

        [DataTestMethod, DataRow(0), DataRow(-1), DataRow(int.MinValue)]
        public void NonpositiveScopeSequenceCannotReuseAnEvidenceName(int sequence)
            => Assert.ThrowsException<ArgumentOutOfRangeException>(() =>
                OfficeVbeFixture.CoverageReceiptPath(Path.GetTempPath(), "coverage-copy-consent", sequence, "intent"));

        [DataTestMethod, DataRow(null), DataRow(""), DataRow(" "), DataRow("relative-folder")]
        public void ReceiptRootMustBeExplicitAndAbsolute(string root)
            => Assert.ThrowsException<ArgumentException>(() =>
                OfficeVbeFixture.CoverageReceiptPath(root, "word-coverage-events", 1, "enable-intent"));

        [TestMethod]
        public void DefaultQualificationDoesNotInspectOrActivateAnySecurityWindow()
        {
            WithOptIn(null, fixture => Assert.IsNull(fixture.CreateCoverageCopyConsent(null, null)));
        }

        [DataTestMethod, DataRow("foreignHost"), DataRow("unownedWord"), DataRow("unownedPowerPoint")]
        public void AnEnabledDiagnosticCannotAdoptAForeignOrUnownedApplication(string scope)
        {
            WithOptIn("1", fixture =>
            {
                string host = scope == "foreignHost" ? "Excel" : scope == "unownedWord" ? "Word" : "PowerPoint";
                typeof(OfficeVbeFixture).GetField("<Kind>k__BackingField", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(fixture, host);
                Assert.ThrowsException<AssertFailedException>(() => fixture.CreateCoverageCopyConsent(null, null));
            });
        }

        private static void WithOptIn(string value, Action<OfficeVbeFixture> test)
        {
            const string variable = "VBAi_RUN_OFFICE_COPY_CONSENT";
            string before = Environment.GetEnvironmentVariable(variable);
            try
            {
                Environment.SetEnvironmentVariable(variable, value);
                var fixture = (OfficeVbeFixture)Activator.CreateInstance(typeof(OfficeVbeFixture), true);
                test(fixture);
            }
            finally { Environment.SetEnvironmentVariable(variable, before); }
        }
    }
}
