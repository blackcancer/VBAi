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
    public sealed class NativeFixtureProgressTraceTests
    {
        [TestMethod]
        public void DisabledAndFailedWriterPreserveOneActionAndItsExactFailure()
        {
            foreach (bool failing in new[] { false, true })
            {
                int calls = 0; var failure = new InvalidOperationException("private source content");
                using (NativeFixtureProgressTrace.Begin(failing ? new Action<string>(line => { throw new IOException(); }) : null))
                {
                    try { NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.InitialSaveAs, () => { calls++; throw failure; }); }
                    catch (InvalidOperationException actual) { Assert.AreSame(failure, actual); }
                }
                Assert.AreEqual(1, calls);
            }
        }
        [TestMethod]
        public void TraceRecordsEnteredReturnedAndFaultedAroundOriginalActionsWithoutValues()
        {
            var lines = new List<string>(); var failure = new IOException("SECRET_DOCUMENT_NAME");
            using (NativeFixtureProgressTrace.Begin(lines.Add))
            {
                Assert.AreEqual(82700, NativeFixtureProgressTrace.Read(NativeFixtureProgressTrace.Phase.PrintWindow, () => 82700));
                Assert.ThrowsException<IOException>(() => NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.ReleaseCode, () => { throw failure; }));
            }
            var rows = lines.Select(line => new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(line)).ToArray();
            CollectionAssert.AreEqual(new[] { "Scope:Entered", "PrintWindow:Entered", "PrintWindow:Returned", "ReleaseCode:Entered", "ReleaseCode:Faulted", "Scope:Returned" },
                rows.Select(row => row["Phase"] + ":" + row["Boundary"]).ToArray());
            Assert.IsTrue(lines.All(line => !line.Contains("SECRET_DOCUMENT_NAME") && !line.Contains("82700")));
        }
        [TestMethod]
        public void ScopeRestoresPreviousWriterAndEventBudgetDoesNotSkipActions()
        {
            var outer = new List<string>(); var inner = new List<string>(); int calls = 0;
            using (NativeFixtureProgressTrace.Begin(outer.Add))
            {
                using (NativeFixtureProgressTrace.Begin(inner.Add)) NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.FormAdd, () => calls++);
                for (int i = 0; i < NativeFixtureProgressTrace.MaximumEvents; i++) NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.PrintWindow, () => calls++);
            }
            Assert.AreEqual(NativeFixtureProgressTrace.MaximumEvents + 1, calls);
            Assert.AreEqual(NativeFixtureProgressTrace.MaximumEvents, outer.Count);
            Assert.AreEqual(4, inner.Count);
        }
        [TestMethod]
        public void FreshUtf8TraceNeverTruncatesExistingEvidenceAndHasNoBom()
        {
            string folder = Path.Combine(Path.GetTempPath(), "VBAi-trace-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(folder);
            string file = Path.Combine(folder, "setup.jsonl");
            try
            {
                using (var trace = NativeFixtureProgressTrace.BeginAtPath(file))
                { Assert.IsTrue(trace.WriterAvailable); NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.InitialSaveAs, () => { }); }
                byte[] before = File.ReadAllBytes(file); Assert.AreEqual((byte)'{', before[0]);
                using (var trace = NativeFixtureProgressTrace.BeginAtPath(file)) Assert.IsFalse(trace.WriterAvailable);
                CollectionAssert.AreEqual(before, File.ReadAllBytes(file));
                Assert.AreEqual(4, File.ReadAllLines(file).Length);
            }
            finally { File.Delete(file); Directory.Delete(folder); }
        }
    }
}
