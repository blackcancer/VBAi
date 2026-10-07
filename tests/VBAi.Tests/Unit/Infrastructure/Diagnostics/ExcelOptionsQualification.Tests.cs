using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.Script.Serialization;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ExcelOptionsQualificationTests
    {
        private const string Baseline = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";
        private const string Changed = "bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb";
        private sealed class Probe
        {
            internal string Fault;
            internal string EvidenceFault;
            internal int Reads, Writes, Cleanup, Preserved;
            internal bool Value = true;
            internal readonly List<string> Phases = new List<string>();
            internal readonly List<IDictionary<string, object>> Requests = new List<IDictionary<string, object>>();
            internal readonly List<object> Records = new List<object>();
            internal ExcelOptionsQualification Create() => new ExcelOptionsQualification(Dispatch, () => Preserved++, () =>
            {
                Cleanup++; if (Fault == "Cleanup" || Fault == "PrimaryAndCleanup") throw new IOException("cleanup failure");
            }, (phase, value) =>
            {
                Phases.Add(phase); Records.Add(value);
                if (EvidenceFault == phase) { EvidenceFault = null; throw new IOException("evidence failure: " + phase); }
            });
            private IDictionary<string, object> Dispatch(object request)
            {
                var json = new JavaScriptSerializer();
                var item = json.Deserialize<Dictionary<string, object>>(json.Serialize(request));
                Assert.IsTrue(Phases.Last().EndsWith("Intent"), "Intent must be recorded before any emission.");
                Requests.Add(item);
                if ((string)item["Command"] == "read_vbe_options")
                {
                    Reads++;
                    if (Fault == "BeforeWritePipeClosed" && Reads == 2)
                        throw new IOException("response lost after emission");
                    if (Fault == "AfterWriteTimeout" && Reads == 3) throw new TimeoutException("response timeout after emission");
                    if ((Fault == "ReadbackMismatch" || Fault == "PrimaryAndRestoration" || Fault == "PrimaryAndCleanup") && Reads == 3)
                        return Reply(State(true));
                    if (Fault == "PrimaryAndRestoration" && Reads == 4) throw new TimeoutException("restoration read uncertain");
                    var state = State(Value);
                    if (Fault == "RevisionNotRestored" && Reads == 5) state["OptionsVersion"] = Changed;
                    return Reply(state);
                }
                Writes++;
                Assert.AreEqual(Value ? Baseline : Changed, item["ExpectedOptionsVersion"], "Every write must use the exact most recent complete revision.");
                bool next = (bool)item["Value"];
                Value = next;
                if (Fault == "WritePipeClosed") throw new IOException("write response lost after emission");
                var data = new Dictionary<string, object>
                {
                    ["CommitRequested"] = true,
                    ["ControlValueVerified"] = true,
                    ["DialogClosed"] = Fault != "DialogPending",
                    ["After"] = Value ? "On" : "Off"
                };
                if (Fault == "ReplyUncertain") data["Uncertain"] = true;
                if (Fault == "ReplyPending") data["Pending"] = true;
                return Reply(data);
            }
        }
        private static IDictionary<string, object> Reply(object data) => new Dictionary<string, object> { ["Ok"] = true, ["Data"] = data };
        private static IDictionary<string, object> State(bool value) => new Dictionary<string, object>
        {
            ["OptionsVersion"] = value ? Baseline : Changed,
            ["DialogClosed"] = true,
            ["Tabs"] = new object[] { Tab("Format de l'éditeur", "Barre des indicateurs en marge", value), Tab("Ancrage", "Exécution", value) }
        };
        private static IDictionary<string, object> Tab(string name, string property, bool value) => new Dictionary<string, object>
        {
            ["Tab"] = name,
            ["Controls"] = new object[] { new Dictionary<string, object> {
                ["Name"] = property, ["Type"] = "ControlType.CheckBox", ["Value"] = value ? "On" : "Off", ["Error"] = null } }
        };

        [TestMethod]
        public void BothNativeTabsRestoreTheCompleteBaselineAndUseFreshRevisionGuards()
        {
            var probe = new Probe(); probe.Create().Run();
            Assert.AreEqual(4, probe.Writes); Assert.AreEqual(9, probe.Reads);
            Assert.AreEqual(1, probe.Cleanup); Assert.AreEqual(0, probe.Preserved);
            Assert.IsTrue(probe.Value);
            Assert.AreEqual(2, probe.Phases.Count(phase => phase == "BaselineRestored"));
        }

        [DataTestMethod]
        [DataRow("BeforeWritePipeClosed", 0, 2)]
        [DataRow("WritePipeClosed", 1, 2)]
        [DataRow("AfterWriteTimeout", 1, 3)]
        [DataRow("ReplyUncertain", 1, 2)]
        [DataRow("ReplyPending", 1, 2)]
        [DataRow("DialogPending", 1, 2)]
        public void UncertainDeliveryStopsAllFurtherDispatchAndNativeCleanup(string fault, int writes, int reads)
        {
            var probe = new Probe { Fault = fault }; var lifecycle = probe.Create();
            Exception failure = null;
            try { lifecycle.Run(); } catch (Exception error) { failure = error; }
            Assert.IsNotNull(failure, "Uncertain native delivery must remain a failure.");
            Assert.IsTrue(lifecycle.HostRetained);
            Assert.AreEqual(writes, probe.Writes); Assert.AreEqual(reads, probe.Reads);
            Assert.AreEqual(0, probe.Cleanup); Assert.AreEqual(1, probe.Preserved);
            Assert.IsFalse(probe.Phases.Contains("RestorationIntent"));
        }

        [TestMethod]
        public void KnownReadbackMismatchRestoresThenPreservesTheOriginalAssertionFailure()
        {
            var probe = new Probe { Fault = "ReadbackMismatch" };
            var error = Assert.ThrowsException<InvalidOperationException>(() => probe.Create().Run());
            StringAssert.Contains(error.Message, "checkbox readback");
            Assert.AreEqual(2, probe.Writes); Assert.IsTrue(probe.Value);
            Assert.AreEqual(1, probe.Cleanup); Assert.AreEqual(0, probe.Preserved);
        }

        [TestMethod]
        public void PrimaryAndUncertainRestorationFailuresRemainDistinctAndRetainHost()
        {
            var probe = new Probe { Fault = "PrimaryAndRestoration" };
            var error = Assert.ThrowsException<AggregateException>(() => probe.Create().Run());
            Assert.AreEqual(2, error.InnerExceptions.Count);
            StringAssert.Contains(error.InnerExceptions[0].Message, "checkbox readback");
            StringAssert.Contains(error.InnerExceptions[1].Message, "restoration read uncertain");
            Assert.AreEqual(1, probe.Writes); Assert.AreEqual(0, probe.Cleanup); Assert.AreEqual(1, probe.Preserved);
        }

        [TestMethod]
        public void BaselineRevisionMismatchIsNeverReportedAsRestored()
        {
            var probe = new Probe { Fault = "RevisionNotRestored" };
            var error = Assert.ThrowsException<InvalidOperationException>(() => probe.Create().Run());
            StringAssert.Contains(error.Message, "complete native preferences");
            Assert.AreEqual(0, probe.Cleanup); Assert.AreEqual(1, probe.Preserved);
            Assert.IsFalse(probe.Phases.Contains("BaselineRestored"));
        }

        [TestMethod]
        public void PrimaryAndCleanupErrorsAreBothRetainedAfterVerifiedRestoration()
        {
            var probe = new Probe { Fault = "PrimaryAndCleanup" };
            var error = Assert.ThrowsException<AggregateException>(() => probe.Create().Run());
            Assert.AreEqual(2, error.InnerExceptions.Count);
            StringAssert.Contains(error.InnerExceptions[0].Message, "checkbox readback");
            StringAssert.Contains(error.InnerExceptions[1].Message, "cleanup failure");
            Assert.AreEqual(2, probe.Writes); Assert.IsTrue(probe.Value);
        }

        [DataTestMethod]
        [DataRow("BeforeMutationIntent", 0)]
        [DataRow("MutationReply", 2)]
        public void EvidenceFailureNeverHidesKnownNativeCompletionOrCreatesAnUnnecessaryWrite(string phase, int writes)
        {
            var probe = new Probe { EvidenceFault = phase };
            var error = Assert.ThrowsException<IOException>(() => probe.Create().Run());
            StringAssert.Contains(error.Message, "evidence failure");
            Assert.AreEqual(writes, probe.Writes); Assert.IsTrue(probe.Value);
            Assert.AreEqual(1, probe.Cleanup); Assert.AreEqual(0, probe.Preserved);
        }

        [TestMethod]
        public void PrimaryUncertainReadAndRecordingFailureAreNotMaskedByFinally()
        {
            var probe = new Probe { Fault = "BeforeWritePipeClosed", EvidenceFault = "RestorationNotEmitted" };
            var error = Assert.ThrowsException<AggregateException>(() => probe.Create().Run());
            Assert.AreEqual(2, error.InnerExceptions.Count);
            StringAssert.Contains(error.InnerExceptions[0].Message, "response lost");
            StringAssert.Contains(error.InnerExceptions[1].Message, "evidence failure");
            Assert.AreEqual(0, probe.Writes); Assert.AreEqual(0, probe.Cleanup); Assert.AreEqual(1, probe.Preserved);
        }

        [TestMethod]
        public void SemanticSummaryKeepsTheFullRevisionAndBoundsHugeCatalogues()
        {
            var data = State(true);
            var control = (IDictionary<string, object>)((object[])((IDictionary<string, object>)((object[])data["Tabs"])[0])["Controls"])[0];
            control["Choices"] = new[] { new string('x', 300000) };
            control["Value"] = new string('v', 300000);
            string summary = new JavaScriptSerializer().Serialize(ExcelOptionsQualification.Summary(data));
            Assert.IsTrue(summary.Length < 2048); StringAssert.Contains(summary, Baseline);
            StringAssert.Contains(summary, "CataloguesOmitted"); StringAssert.Contains(summary, "\"ValueTruncated\":true");
            Assert.IsFalse(summary.Contains(new string('x', 100)));
        }
    }
}
