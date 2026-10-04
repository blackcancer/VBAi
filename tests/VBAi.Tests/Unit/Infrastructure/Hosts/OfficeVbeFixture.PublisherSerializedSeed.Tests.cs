using System;
using System.Collections;
using System.Collections.Generic;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class OfficeVbeFixturePublisherSerializedSeedTests
    {
        private const string ModuleCode = "Option Explicit\r\n' Synthetic native General qualification: pending-c2468bacba634672be8b7ce6d656fd37\r\n";
        private const string ModuleSha = "48f46f56d9e9a767aabd8d9f79aadc4873b63cf4ff8672d735f5164eb91477e4";

        private static string ShutdownJson() => new JavaScriptSerializer().Serialize(new {
            DocumentPath = OfficeVbeFixture.PublisherSeedPath,
            Lifecycle = new { ProcessId = 183824, ProcessStartedUtc = "2026-10-03T19:42:01.6309813Z",
                ProcessExitObserved = true, ExitCodeObserved = true, ExitCode = 0, ForcedTermination = false,
                State = "EXIT_OBSERVED_HANDLE_RELEASED" } });

        private static string ProgressJson(string sha, long bytes = 90624) => new JavaScriptSerializer().Serialize(new {
            Steps = new object[] { new { AdapterStage = "PublisherGeneralClosedFileEvidence", Inputs = new {
                Path = OfficeVbeFixture.PublisherSeedPath, Bytes = bytes, Sha256 = sha,
                PreviousProcessId = 183824, HostExitedBeforeHash = true } } } });

        [TestMethod]
        public void ActualTypedSerializerArrayListReceiptIsValidatedBeforeHostCreation()
        {
            string progress = ProgressJson(OfficeVbeFixture.PublisherSeedSha);
            var parsed = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(progress);
            Assert.IsInstanceOfType(parsed["Steps"], typeof(ArrayList));
            OfficeVbeFixture.RequirePublisherSeedReceipts(ShutdownJson(), progress);
        }

        [TestMethod]
        public void TypedReceiptParserStillRefusesChangedProvenanceAndNonArraySteps()
        {
            Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequirePublisherSeedReceipts(ShutdownJson(), ProgressJson("wrong-sha")));
            Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequirePublisherSeedReceipts(ShutdownJson(), ProgressJson(OfficeVbeFixture.PublisherSeedSha, 1)));
            Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequirePublisherSeedReceipts(ShutdownJson().Replace("183824", "183825"), ProgressJson(OfficeVbeFixture.PublisherSeedSha)));
            Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequirePublisherSeedReceipts(ShutdownJson(), "{\"Steps\":\"not-an-array\"}"));
        }

        [TestMethod]
        public void MissingOrDuplicateClosedFileReceiptsRemainRefused()
        {
            Assert.ThrowsException<AssertFailedException>(() => OfficeVbeFixture.RequirePublisherSeedReceipts(ShutdownJson(), "{\"Steps\":[]}"));
            var serializer = new JavaScriptSerializer();
            var original = serializer.Deserialize<Dictionary<string, object>>(ProgressJson(OfficeVbeFixture.PublisherSeedSha));
            var step = ((ArrayList)original["Steps"])[0];
            string duplicate = serializer.Serialize(new { Steps = new[] { step, step } });
            Assert.ThrowsException<AssertFailedException>(() => OfficeVbeFixture.RequirePublisherSeedReceipts(ShutdownJson(), duplicate));
        }

        [TestMethod]
        public void OnlyFrozenOwnedInputIsAcceptedBeforeCopyOrOpen()
        {
            OfficeVbeFixture.RequirePublisherSeedBytes(OfficeVbeFixture.PublisherSeedPath, 90624, OfficeVbeFixture.PublisherSeedSha);
            Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequirePublisherSeedBytes("C:\\foreign.pub", 90624, OfficeVbeFixture.PublisherSeedSha));
            Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequirePublisherSeedBytes(OfficeVbeFixture.PublisherSeedPath, 90625, OfficeVbeFixture.PublisherSeedSha));
            Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequirePublisherSeedBytes(OfficeVbeFixture.PublisherSeedPath, 90624, "wrong-sha"));
        }

        [TestMethod]
        public void InertExactTypedComponentIsAccepted()
        {
            OfficeVbeFixture.RequirePublisherSeedComponent("PublisherGeneralModule", 1, ModuleCode, ModuleSha);
            OfficeVbeFixture.RequirePublisherSeedComponent("ThisDocument", 100, "", null);
        }

        [DataTestMethod]
        [DataRow("PublisherGeneralModule", 2, "source", "hash")]
        [DataRow("ExtraModule", 1, "source", "hash")]
        [DataRow("ThisDocument", 100, "Private Sub Document_Open()\r\nEnd Sub", "hash")]
        [DataRow("ThisDocument", 1, "", "hash")]
        public void WrongTypeForeignComponentOrProcedureIsRefused(string name, int type, string code, string sha)
        { Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequirePublisherSeedComponent(name, type, code, sha)); }

        [TestMethod]
        public void SourceAndRevisionDriftAreRefused()
        {
            Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequirePublisherSeedComponent("PublisherGeneralModule", 1, ModuleCode + "Sub Run()\r\nEnd Sub", ModuleSha));
            Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequirePublisherSeedComponent("PublisherGeneralModule", 1, ModuleCode, "wrong-revision"));
        }

        [DataTestMethod]
        [DataRow(183825, "2026-10-03T19:42:01.6309813Z", true, 0, false)]
        [DataRow(183824, "wrong-birth", true, 0, false)]
        [DataRow(183824, "2026-10-03T19:42:01.6309813Z", false, 0, false)]
        [DataRow(183824, "2026-10-03T19:42:01.6309813Z", true, 1, false)]
        [DataRow(183824, "2026-10-03T19:42:01.6309813Z", true, 0, true)]
        public void ChangedProducerOrUncertainClosureIsRefusedBeforeInputUse(int pid, string birth, bool exited, int exitCode, bool forced)
        { Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequirePublisherSeedProducer(pid, birth, exited, exitCode, forced, "EXIT_OBSERVED_HANDLE_RELEASED", OfficeVbeFixture.PublisherSeedPath)); }
    }
}
