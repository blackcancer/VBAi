using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace VBAi.Tests.Unit
{
    /// <summary>Checks that diagnostic publication cannot change native operation semantics.</summary>
    [ToolBoundaryTestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class NativeExecutionTimingsTests
    {
        [TestMethod]
        [DataRow(0), DataRow(1), DataRow(2), DataRow(3), DataRow(4), DataRow(5)]
        [DataRow(6), DataRow(7), DataRow(8), DataRow(9), DataRow(10), DataRow(11)]
        [DataRow(12), DataRow(13), DataRow(14)]
        public void StagePairsRetainReturnIdentityAndOriginalThread(int stage)
        {
            var events = new List<NativeExecutionTimings.StageEvent>();
            var timings = new NativeExecutionTimings(events.Add);
            var expected = new object();
            int calls = 0, thread = Thread.CurrentThread.ManagedThreadId;
            var value = timings.Measure((NativeTestTimingStage)stage, () => { calls++; return expected; });
            Assert.AreSame(expected, value); Assert.AreEqual(1, calls);
            Assert.AreEqual(2, events.Count);
            Assert.AreEqual("Started", events[0].State); Assert.AreEqual("Returned", events[1].State);
            Assert.AreEqual(events[0].Span, events[1].Span);
            Assert.AreEqual(events[0].ObserverId, events[1].ObserverId);
            Assert.IsTrue(events.All(item => item.ManagedThreadId == thread && item.Apartment == Thread.CurrentThread.GetApartmentState().ToString()));
            Assert.IsTrue(events[1].Milliseconds >= 0);
            Assert.IsNull(timings.DiagnosticFailure);
        }

        [TestMethod]
        [DataRow(1), DataRow(8), DataRow(10)]
        [DataRow(12), DataRow(13), DataRow(14)]
        public void NativeFailureRetainsExactExceptionWithoutWritingItsPrivateMessage(int stage)
        {
            var events = new List<NativeExecutionTimings.StageEvent>();
            var timings = new NativeExecutionTimings(events.Add);
            var failure = new InvalidOperationException("synthetic private source and argument");
            int calls = 0;
            var observed = Assert.ThrowsException<InvalidOperationException>(() => timings.Measure((NativeTestTimingStage)stage,
                () => { calls++; throw failure; }));
            Assert.AreSame(failure, observed); Assert.AreEqual(1, calls);
            Assert.AreEqual("Failed", events[1].State);
            Assert.AreEqual(typeof(InvalidOperationException).FullName, events[1].ExceptionType);
            Assert.IsFalse(new JavaScriptSerializer().Serialize(events).Contains(failure.Message));
        }

        [TestMethod]
        [DataRow(true), DataRow(false)]
        public void PublicationFailureCannotReplaceTheNativeReturn(bool atEntry)
        {
            var observerFailure = new IOException("synthetic diagnostic failure");
            int publications = 0, calls = 0;
            var timings = new NativeExecutionTimings(item =>
            {
                publications++;
                if ((item.State == "Started") == atEntry) throw observerFailure;
            });
            Assert.AreEqual(42, timings.Measure(NativeTestTimingStage.ExecuteNativeDispatch, () => { calls++; return 42; }));
            Assert.AreEqual(1, calls); Assert.AreSame(observerFailure, timings.DiagnosticFailure);
            int originalPublications = publications;
            Assert.AreEqual(43, timings.Measure(NativeTestTimingStage.ExecuteNativeDispatch, () => { calls++; return 43; }));
            Assert.AreEqual(2, calls); Assert.AreEqual(originalPublications, publications);
        }

        [TestMethod]
        [DataRow(true), DataRow(false)]
        public void NativeAndObserverFailuresRemainDistinctWithoutAReplay(bool atEntry)
        {
            var nativeFailure = new InvalidOperationException("native");
            var observerFailure = new IOException("observer");
            var timings = new NativeExecutionTimings(item => { if ((item.State == "Started") == atEntry) throw observerFailure; });
            int calls = 0;
            var observed = Assert.ThrowsException<InvalidOperationException>(() => timings.Measure(NativeTestTimingStage.ExecuteNativeDispatch,
                () => { calls++; throw nativeFailure; }));
            Assert.AreSame(nativeFailure, observed); Assert.AreSame(observerFailure, timings.DiagnosticFailure);
            Assert.AreEqual(1, calls);
        }

        [TestMethod]
        public void NestedOperationsHaveIndependentSpansAndOrderedCompletion()
        {
            var events = new List<NativeExecutionTimings.StageEvent>();
            var timings = new NativeExecutionTimings(events.Add);
            Assert.AreEqual(23, timings.Measure(NativeTestTimingStage.PrepareNativeDispatch,
                () => timings.Measure(NativeTestTimingStage.ProjectSnapshot, () => 23)));
            CollectionAssert.AreEqual(new[] { "Started", "Started", "Returned", "Returned" }, events.Select(item => item.State).ToArray());
            Assert.AreEqual(events[0].Span, events[3].Span); Assert.AreEqual(events[1].Span, events[2].Span);
            Assert.AreNotEqual(events[0].Span, events[1].Span);
            Assert.AreEqual(1, events.Select(item => item.ObserverId).Distinct().Count());
        }

        [TestMethod]
        public void EventLimitBoundsOutputWithoutDroppingAnOperation()
        {
            var events = new List<NativeExecutionTimings.StageEvent>();
            var timings = new NativeExecutionTimings(events.Add);
            int calls = 0, attempts = NativeExecutionTimings.MaximumRecords / 2 + 4;
            for (int i = 0; i < attempts; i++) timings.Measure(NativeTestTimingStage.ReadNativeMode, () => { calls++; });
            Assert.AreEqual(attempts, calls); Assert.AreEqual(NativeExecutionTimings.MaximumRecords, events.Count);
            Assert.IsTrue(timings.Truncated);
        }

        [TestMethod]
        [DataRow(null), DataRow(""), DataRow(" "), DataRow("\t")]
        public void UnconfiguredDirectoryKeepsDiagnosticsDisabled(string root)
        {
            Assert.AreSame(NativeExecutionTimings.Disabled, NativeExecutionTimings.ForRoot(root));
            Assert.AreEqual(7, NativeExecutionTimings.ForRoot(root).Measure(NativeTestTimingStage.ProjectSnapshot, () => 7));
        }

        [TestMethod]
        [DataRow("relative"), DataRow("missing"), DataRow("file")]
        public void InvalidDirectoryCannotPreventAnOperation(string condition)
        {
            string path = condition == "relative" ? "relative-diagnostic-folder" : Path.Combine(Path.GetTempPath(), "VBAi-Timing-" + Guid.NewGuid().ToString("N"));
            try
            {
                if (condition == "file") File.WriteAllText(path, "owned file");
                var timings = NativeExecutionTimings.ForRoot(path);
                int calls = 0;
                Assert.AreEqual(19, timings.Measure(NativeTestTimingStage.ExecuteNativeDispatch, () => { calls++; return 19; }));
                Assert.AreEqual(1, calls); Assert.IsInstanceOfType(timings.DiagnosticFailure, typeof(ArgumentException));
                if (condition != "file") Assert.IsFalse(Directory.Exists(path));
            }
            finally { if (condition == "file" && File.Exists(path)) File.Delete(path); }
        }

        [TestMethod]
        public void JsonLinesRecordOnlyMetadataAndSeparateRecordersSharingAProcess()
        {
            string root = Path.Combine(Path.GetTempPath(), "VBAi-Timing-é-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try
            {
                var first = NativeExecutionTimings.ForRoot(root);
                var second = NativeExecutionTimings.ForRoot(root);
                first.Measure(NativeTestTimingStage.ProjectSnapshot, () => { });
                second.Measure(NativeTestTimingStage.ExecuteNativeDispatch, () => { });
                string file = Directory.GetFiles(root).Single();
                byte[] bytes = File.ReadAllBytes(file);
                Assert.IsFalse(bytes.Length >= 3 && bytes[0] == 239 && bytes[1] == 187 && bytes[2] == 191);
                var json = new JavaScriptSerializer();
                var rows = File.ReadAllLines(file).Select(line => json.Deserialize<Dictionary<string, object>>(line)).ToArray();
                Assert.AreEqual(4, rows.Length);
                Assert.IsTrue(rows.All(item => Convert.ToString(item["ProductMvid"]) == typeof(NativeExecutionTimings).Module.ModuleVersionId.ToString("D")));
                var events = rows.Select(item => (Dictionary<string, object>)item["Event"]).ToArray();
                Assert.AreEqual(2, events.Select(item => item["ObserverId"]).Distinct().Count());
                Assert.IsFalse(events.Any(item => item.ContainsKey("Source") || item.ContainsKey("Project") || item.ContainsKey("Arguments") || item.ContainsKey("Message")));
                Assert.IsNull(first.DiagnosticFailure); Assert.IsNull(second.DiagnosticFailure);
            }
            finally { foreach (string file in Directory.GetFiles(root)) File.Delete(file); Directory.Delete(root); }
        }

        [TestMethod]
        public void ConcurrentDiagnosticCallersRetainPairsAndBoundedUniqueSequences()
        {
            var events = new List<NativeExecutionTimings.StageEvent>();
            var timings = new NativeExecutionTimings(events.Add);
            int calls = 0;
            Parallel.For(0, 32, _ => timings.Measure(NativeTestTimingStage.ProjectSnapshot, () => { Interlocked.Increment(ref calls); }));
            Assert.AreEqual(32, calls); Assert.AreEqual(64, events.Count);
            CollectionAssert.AreEqual(Enumerable.Range(1, 64).ToArray(), events.Select(item => item.Sequence).ToArray());
            Assert.IsTrue(events.GroupBy(item => item.Span).All(pair => pair.Count() == 2 && pair.First().State == "Started" && pair.Last().State == "Returned"));
        }

        [TestMethod]
        public void DisabledAndEnabledObserversRejectMissingActionsBeforePublication()
        {
            var events = new List<NativeExecutionTimings.StageEvent>();
            foreach (var timings in new[] { NativeExecutionTimings.Disabled, new NativeExecutionTimings(events.Add) })
            {
                Assert.ThrowsException<ArgumentNullException>(() => timings.Measure(NativeTestTimingStage.ProjectSnapshot, (Action)null));
                Assert.ThrowsException<ArgumentNullException>(() => timings.Measure<object>(NativeTestTimingStage.ProjectSnapshot, null));
            }
            Assert.AreEqual(0, events.Count);
        }

        [TestMethod]
        [DataRow(false), DataRow(true)]
        public void EnvironmentOptInRequiresAnExplicitExistingDirectory(bool enabled)
        {
            string before = Environment.GetEnvironmentVariable(NativeExecutionTimings.RootEnvironmentVariable);
            string root = Path.Combine(Path.GetTempPath(), "VBAi-Timing-Environment-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try
            {
                Environment.SetEnvironmentVariable(NativeExecutionTimings.RootEnvironmentVariable, enabled ? root : null);
                var timings = NativeExecutionTimings.ForCurrentProcess();
                timings.Measure(NativeTestTimingStage.ProjectSnapshot, () => { });
                Assert.AreEqual(enabled ? 1 : 0, Directory.GetFiles(root).Length);
                Assert.IsNull(timings.DiagnosticFailure);
            }
            finally
            {
                Environment.SetEnvironmentVariable(NativeExecutionTimings.RootEnvironmentVariable, before);
                foreach (string file in Directory.GetFiles(root)) File.Delete(file);
                Directory.Delete(root);
            }
        }

        [TestMethod]
        public void LockedDiagnosticFileCannotPreventExecutionOrTriggerAnotherWrite()
        {
            string root = Path.Combine(Path.GetTempPath(), "VBAi-Timing-Locked-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try
            {
                int pid;
                using (var process = System.Diagnostics.Process.GetCurrentProcess()) pid = process.Id;
                string path = Path.Combine(root, "native-stages-" + pid + ".jsonl");
                var timings = NativeExecutionTimings.ForRoot(root);
                using (var locked = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None))
                {
                    Assert.AreEqual(31, timings.Measure(NativeTestTimingStage.ExecuteNativeDispatch, () => 31));
                    Assert.IsInstanceOfType(timings.DiagnosticFailure, typeof(IOException));
                }
                Assert.AreEqual(32, timings.Measure(NativeTestTimingStage.ExecuteNativeDispatch, () => 32));
                Assert.AreEqual(0L, new FileInfo(path).Length);
            }
            finally { foreach (string file in Directory.GetFiles(root)) File.Delete(file); Directory.Delete(root); }
        }
    }
}
