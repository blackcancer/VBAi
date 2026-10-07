using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Web.Script.Serialization;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class OfficeVbeFixturePublisherSerializedSeedTests
    {
        private const string ModuleCode = "Option Explicit\r\n' Synthetic native General qualification: pending-c2468bacba634672be8b7ce6d656fd37\r\n";
        private const string ModuleSha = "48f46f56d9e9a767aabd8d9f79aadc4873b63cf4ff8672d735f5164eb91477e4";

        [TestMethod]
        public void MainPublisherSelectionRequiresExactSeedOptInAndEmptyPrivateDescriptors()
        {
            int checks = 0;
            Action requireMain = () => checks++;
            string seed = OfficeVbeFixture.PublisherSeedPath;
            Assert.IsTrue(OfficeVbeFixture.SelectMainPublisherDesktop("Publisher", seed, "1", null, null, requireMain));
            Assert.AreEqual(1, checks);
            Assert.IsFalse(OfficeVbeFixture.SelectMainPublisherDesktop("Publisher", seed, null, null, null, requireMain));
            Assert.IsFalse(OfficeVbeFixture.SelectMainPublisherDesktop("Publisher", null, "1", null, null, requireMain));
            Assert.IsFalse(OfficeVbeFixture.SelectMainPublisherDesktop("Access", seed, "1", null, null, requireMain));
            Assert.AreEqual(1, checks);
            Assert.ThrowsException<InvalidOperationException>(() =>
                OfficeVbeFixture.SelectMainPublisherDesktop("Publisher", seed, "1", "VBAiTests_owned", null, requireMain));
            Assert.ThrowsException<InvalidOperationException>(() =>
                OfficeVbeFixture.SelectMainPublisherDesktop("Publisher", seed, "1", null, "VBAiTests_owned", requireMain));
            Assert.ThrowsException<ArgumentNullException>(() =>
                OfficeVbeFixture.SelectMainPublisherDesktop("Publisher", seed, "1", null, null, null));
            Assert.AreEqual(1, checks);
        }

        [TestMethod]
        public void MainPublisherRequiresOwnedRootPidThreadAndCompleteDefaultInventory()
        {
            var root = new IntPtr(42);
            var child = new IntPtr(43);
            var inventory = new IsolatedTestDesktop.MainInventory
            {
                Desktop = "Default", Complete = true, Visited = 1,
                Windows = new[] { new IsolatedTestDesktop.MainWindow
                    { Handle = root.ToInt64(), ProcessId = 123, ThreadId = 7, ClassName = "PublisherRoot", Visible = true } }
            };
            OfficeVbeFixture.RequireMainPublisherPlacement(inventory, 123, true, child, root, 123, 123, 7, 7);
            Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequireMainPublisherPlacement(
                inventory, 123, true, child, root, 124, 123, 7, 7));
            Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequireMainPublisherPlacement(
                inventory, 123, true, child, root, 123, 123, 0, 7));
            Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequireMainPublisherPlacement(
                inventory, 123, true, child, new IntPtr(99), 123, 123, 7, 7));
            inventory.Complete = false;
            Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequireMainPublisherPlacement(
                inventory, 123, false, IntPtr.Zero, IntPtr.Zero, 0, 0, 0, 0));
            inventory.Complete = true; inventory.Desktop = "VBAiTests_private";
            Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequireMainPublisherPlacement(
                inventory, 123, false, IntPtr.Zero, IntPtr.Zero, 0, 0, 0, 0));
            inventory.Desktop = "Default"; inventory.Windows = new IsolatedTestDesktop.MainWindow[0];
            Assert.ThrowsException<InvalidOperationException>(() => OfficeVbeFixture.RequireMainPublisherPlacement(
                inventory, 123, true, IntPtr.Zero, IntPtr.Zero, 0, 0, 0, 0));
        }

        [TestMethod]
        public void MainPublisherImageValidationDoesNotAdmitDefaultToPrivateExecutableGate()
        {
            const string image = @"C:\Program Files\Microsoft Office\root\Office16\MSPUB.EXE";
            Assert.AreEqual(image, OfficeVbeFixture.RequireMainPublisherExecutable(image));
            Assert.ThrowsException<ArgumentException>(() => OfficeVbeFixture.RequireMainPublisherExecutable(@"C:\foreign\WINWORD.EXE"));
            Assert.ThrowsException<ArgumentException>(() => OfficeVbeFixture.RequireMainPublisherExecutable("MSPUB.EXE"));
            Assert.ThrowsException<ArgumentException>(() =>
                OfficeVbeFixture.RequirePrivateOfficeExecutable("Publisher", "Default", image));
        }

        private static string ShutdownJson() => new JavaScriptSerializer().Serialize(new
        {
            DocumentPath = OfficeVbeFixture.PublisherSeedPath,
            Lifecycle = new
            {
                ProcessId = 183824,
                ProcessStartedUtc = "2026-10-03T19:42:01.6309813Z",
                ProcessExitObserved = true,
                ExitCodeObserved = true,
                ExitCode = 0,
                ForcedTermination = false,
                State = "EXIT_OBSERVED_HANDLE_RELEASED"
            }
        });

        private static string ProgressJson(string sha, long bytes = 90624) => new JavaScriptSerializer().Serialize(new
        {
            Steps = new object[] { new { AdapterStage = "PublisherGeneralClosedFileEvidence", Inputs = new {
                Path = OfficeVbeFixture.PublisherSeedPath, Bytes = bytes, Sha256 = sha,
                PreviousProcessId = 183824, HostExitedBeforeHash = true } } }
        });

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
