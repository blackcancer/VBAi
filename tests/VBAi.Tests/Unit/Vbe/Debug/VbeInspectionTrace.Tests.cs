using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbeInspectionTraceTests
    {
        [TestMethod]
        public void CorrelationFlowsToObserverWithoutLeakingFailureDetailsAndRestoresNestedScope()
        {
            var rows = new List<string>();
            var outer = new VbeInspectionTrace(rows.Add);
            var inner = new VbeInspectionTrace(rows.Add);
            var original = VbeInspectionTrace.Current;
            using (outer.Enter())
            {
                outer.Record(VbeInspectionTrace.Phase.Enqueue);
                using (inner.Enter()) Assert.AreSame(inner, VbeInspectionTrace.Current);
                Assert.AreSame(outer, VbeInspectionTrace.Current);
                Task.Run(() => {
                    Assert.AreSame(outer, VbeInspectionTrace.Current);
                    outer.Record(VbeInspectionTrace.Phase.ObserverEntered);
                    outer.Record(VbeInspectionTrace.Phase.ObserverTerminal, new InvalidOperationException("SECRET_SOURCE_AND_VALUE"));
                }).GetAwaiter().GetResult();
                Assert.AreSame(outer, VbeInspectionTrace.Current);
            }
            Assert.AreSame(original, VbeInspectionTrace.Current);
            var parsed = rows.Select(row => new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(row)).ToArray();
            Assert.AreEqual(3, parsed.Length);
            Assert.AreEqual(1, parsed.Select(row => row["Correlation"]).Distinct().Count());
            Assert.AreEqual("InvalidOperationException", parsed[2]["ErrorType"]);
            CollectionAssert.AreEqual(new[] { "Enqueue", "ObserverEntered", "ObserverTerminal" }, parsed.Select(row => (string)row["Phase"]).ToArray());
            Assert.IsFalse(string.Join("", rows).Contains("SECRET_SOURCE_AND_VALUE"));
            CollectionAssert.AreEquivalent(new[] { "Correlation", "Sequence", "Utc", "HostProcessId", "ThreadId", "Apartment", "ElapsedMilliseconds", "Phase", "ErrorType" }, parsed[0].Keys.ToArray());
        }

        [TestMethod]
        public void EventLimitAndBrokenWriterNeverAlterNativeFailureOrScope()
        {
            int writes = 0;
            var trace = new VbeInspectionTrace(_ => { writes++; throw new IOException("private path"); });
            var original = VbeInspectionTrace.Current;
            var expected = new InvalidOperationException("native failure");
            var actual = Assert.ThrowsException<InvalidOperationException>(() => {
                using (trace.Enter())
                {
                    for (int index = 0; index < 1000; index++) trace.Record(VbeInspectionTrace.Phase.ContextValidation);
                    trace.Record(VbeInspectionTrace.Phase.Terminal, expected);
                    throw expected;
                }
            });
            Assert.AreSame(expected, actual);
            Assert.AreEqual(VbeInspectionTrace.MaximumEvents, writes);
            Assert.AreSame(original, VbeInspectionTrace.Current);
        }

        [TestMethod]
        public void MissingRelativeNetworkAndAlternateStreamDestinationsDisableTracing()
        {
            foreach (string path in new[] { null, "", "relative.jsonl", @"C:relative.jsonl", @"\trace.jsonl", @"\\server\share\trace.jsonl", @"C:\trace.jsonl:secret" })
                Assert.IsNull(VbeInspectionTrace.ForPath(path), path);
        }

        [TestMethod]
        public void DurableFileIsAppendedWithinByteLimitAndUnavailableParentIsPassive()
        {
            string directory = Path.Combine(Path.GetTempPath(), "VBAi-trace-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            string path = Path.Combine(directory, "phases.jsonl");
            try
            {
                var trace = VbeInspectionTrace.ForPath(path);
                trace.Record(VbeInspectionTrace.Phase.Enqueue);
                trace.Record(VbeInspectionTrace.Phase.Terminal);
                Assert.AreEqual(2, File.ReadAllLines(path).Length);
                File.WriteAllBytes(path, new byte[VbeInspectionTrace.MaximumFileBytes]);
                trace.Record(VbeInspectionTrace.Phase.CallbackEntered);
                Assert.AreEqual(VbeInspectionTrace.MaximumFileBytes, new FileInfo(path).Length);
                var unavailable = VbeInspectionTrace.ForPath(Path.Combine(directory, "missing", "trace.jsonl"));
                unavailable.Record(VbeInspectionTrace.Phase.Enqueue);
                Assert.IsFalse(Directory.Exists(Path.Combine(directory, "missing")));
            }
            finally { Directory.Delete(directory, true); }
        }
    }
}
