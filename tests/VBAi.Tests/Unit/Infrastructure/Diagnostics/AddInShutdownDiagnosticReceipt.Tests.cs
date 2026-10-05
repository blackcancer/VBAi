using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Mirrors the independent chain and physical receipt oracles, including incomplete-publication refusals.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class AddInShutdownDiagnosticReceiptTests
    {
        private static List<object> Chain()
        {
            var o = new ShutdownDiagnosticTestData.Observation(); var d = o.Begin(); ShutdownDiagnosticTestData.FinishDispose(d);
            for (int i = 0; i < o.Events.Count; i++) ((Dictionary<string, object>)o.Events[i])["ElapsedTicks"] = (long)i;
            return o.Events;
        }
        private static Dictionary<string, object> Row(object value) { return (Dictionary<string, object>)value; }
        private static void Number(List<object> events)
        { for (int i = 0; i < events.Count; i++) { Row(events[i])["Sequence"] = i + 1; Row(events[i])["ElapsedTicks"] = (long)i; } }

        [DataTestMethod]
        [DataRow("version")][DataRow("nonce")][DataRow("pid")][DataRow("birth")][DataRow("image")][DataRow("product")][DataRow("mvid")][DataRow("hash")]
        [DataRow("thread")][DataRow("apartment")][DataRow("managed")][DataRow("typed sequence")][DataRow("gap")][DataRow("extra")][DataRow("missing")]
        [DataRow("unknown entry")][DataRow("foreign parent")][DataRow("foreign invocation")][DataRow("unknown phase")][DataRow("unknown operation")]
        [DataRow("negative elapsed")][DataRow("decreasing elapsed")][DataRow("nonutc")][DataRow("returned error")][DataRow("missing terminal")]
        [DataRow("duplicate return")][DataRow("returned after fault")][DataRow("clear order")][DataRow("missing clear")][DataRow("noninteger thread")]
        public void CompleteChainRejectsEveryForeignIncompleteOrContradictoryBoundary(string kind)
        {
            var events = Chain(); var last = Row(events.Last()); var first = Row(events.First());
            if (kind == "version") last["Version"] = 2;
            if (kind == "nonce") last["Nonce"] = Guid.NewGuid().ToString("N");
            if (new[] { "pid", "birth", "image", "product", "mvid", "hash" }.Contains(kind))
            {
                var id = Row(last["Identity"]);
                if (kind == "pid") id["ProcessId"] = 999;
                if (kind == "birth") id["ProcessStartedUtc"] = "2026-10-05T17:27:04.6593947Z";
                if (kind == "image") id["HostImagePath"] = @"C:\Foreign\EXCEL.EXE";
                if (kind == "product") id["ProductPath"] = @"C:\Foreign\VBAi.dll";
                if (kind == "mvid") id["ProductMvid"] = Guid.NewGuid().ToString("D");
                if (kind == "hash") id["ProductSha256"] = new string('B', 64);
            }
            if (kind == "thread") Row(last["Thread"])["NativeThreadId"] = 999;
            if (kind == "apartment") Row(last["Thread"])["Apartment"] = "MTA";
            if (kind == "managed") Row(last["Thread"])["ManagedThreadId"] = 0;
            if (kind == "noninteger thread") Row(last["Thread"])["NativeThreadId"] = "73";
            if (kind == "typed sequence") last["Sequence"] = Convert.ToString(last["Sequence"]);
            if (kind == "gap") last["Sequence"] = events.Count + 1;
            if (kind == "extra") last["NativeRcwReleaseProven"] = true;
            if (kind == "missing") last.Remove("Error");
            if (kind == "unknown entry") first["EntryPoint"] = "OnConnection";
            if (kind == "foreign parent") last["ParentInvocationId"] = Guid.NewGuid().ToString("N");
            if (kind == "foreign invocation") last["InvocationId"] = Guid.NewGuid().ToString("N");
            if (kind == "unknown phase") last["Phase"] = "Success";
            if (kind == "unknown operation") Row(events[1])["Stage"] = "Unknown.ReleaseRcw";
            if (kind == "negative elapsed") first["ElapsedTicks"] = -1;
            if (kind == "decreasing elapsed") last["ElapsedTicks"] = 0;
            if (kind == "nonutc") last["Utc"] = "2026-10-05T20:10:00.0000000+02:00";
            if (kind == "returned error") last["Error"] = "caught";
            if (kind == "missing terminal") events.RemoveAt(events.Count - 1);
            if (kind == "duplicate return" || kind == "returned after fault")
            {
                var duplicate = ShutdownDiagnosticTestData.Clone(events[2]);
                if (kind == "returned after fault") { Row(events[2])["Phase"] = "Fault"; Row(events[2])["Error"] = "original"; Row(events[2])["FaultCaught"] = true; Row(events[2])["ErrorIsComException"] = false; }
                events.Insert(3, duplicate); Number(events);
            }
            if (kind == "clear order") Row(events.First(value => Row(value)["Phase"] as string == "ManagedReferenceCleared"))["Stage"] = "vbe";
            if (kind == "missing clear") { events.Remove(events.First(value => Row(value)["Phase"] as string == "ManagedReferenceCleared")); Number(events); }
            Assert.ThrowsException<InvalidOperationException>(() => AddInShutdownDiagnosticReceipt.Validate(events.ToArray(), ShutdownDiagnosticTestData.Nonce, ShutdownDiagnosticTestData.Identity()));
        }

        [TestMethod]
        public void CompleteFaultedInvocationAndBackwardUtcRetainTheirPreciseScope()
        {
            var o = new ShutdownDiagnosticTestData.Observation(); var d = o.Begin(); var original = new IOException("original cleanup error");
            d.Enter("modernEditor.Dispose"); d.Fault(original);
            var result = AddInShutdownDiagnosticReceipt.Validate(o.Events.ToArray(), ShutdownDiagnosticTestData.Nonce, ShutdownDiagnosticTestData.Identity());
            Assert.AreEqual("Fault", result.Single().Outcome); Assert.AreEqual("Fault", result.Single().Stages["modernEditor.Dispose"]);
            var returned = Chain(); Row(returned.First())["Utc"] = "2026-10-05T20:00:00.0000000Z";
            Row(returned.Last())["Utc"] = "2026-10-05T19:59:00.0000000Z";
            Assert.AreEqual("Returned", AddInShutdownDiagnosticReceipt.Validate(returned.ToArray(), ShutdownDiagnosticTestData.Nonce, ShutdownDiagnosticTestData.Identity()).Single().Outcome);
        }

        [DataTestMethod, DataRow("propagated modern fault"), DataRow("invented modern catch"), DataRow("dispatcher moved"), DataRow("noncom test catch")]
        public void SourceOperationOrderAndExactCatchPoliciesCannotBeRewrittenAsSuccessfulCleanup(string kind)
        {
            var events = Chain();
            string stage = kind == "dispatcher moved" ? "dispatcher.Dispose" : kind == "noncom test catch" ? "nativeTestWindow.Close" : "modernEditor.Dispose";
            string before = kind == "dispatcher moved" ? "editorNavigation" : kind == "noncom test catch" ? "nativeTestWindow" : "modernEditor";
            int position = events.FindIndex(value => Row(value)["Phase"] as string == "ManagedReferenceCleared" && Row(value)["Stage"] as string == before);
            var entry = Row(ShutdownDiagnosticTestData.Clone(events[1])); entry["Stage"] = stage;
            var end = Row(ShutdownDiagnosticTestData.Clone(events[2])); end["Stage"] = stage;
            if (kind != "dispatcher moved") { end["Phase"] = "Fault"; end["Error"] = "original non-COM cleanup failure"; end["FaultCaught"] = kind != "propagated modern fault"; end["ErrorIsComException"] = false; }
            events.Insert(position, entry); events.Insert(position + 1, end); Number(events);
            Assert.ThrowsException<InvalidOperationException>(() => AddInShutdownDiagnosticReceipt.Validate(events.ToArray(), ShutdownDiagnosticTestData.Nonce, ShutdownDiagnosticTestData.Identity()));
        }

        [TestMethod]
        public void CallbackRequiresItsOwnCompletedChildAndCannotAdoptAForeignParent()
        {
            var o = new ShutdownDiagnosticTestData.Observation(); var parent = o.Begin("OnBeginShutdown"); parent.Complete();
            Assert.ThrowsException<InvalidOperationException>(() => AddInShutdownDiagnosticReceipt.Validate(o.Events.ToArray(), ShutdownDiagnosticTestData.Nonce, ShutdownDiagnosticTestData.Identity()));
            o = new ShutdownDiagnosticTestData.Observation(); parent = o.Begin("OnBeginShutdown"); var child = parent.Child("Dispose"); ShutdownDiagnosticTestData.FinishDispose(child); parent.Complete();
            Assert.AreEqual(2, AddInShutdownDiagnosticReceipt.Validate(o.Events.ToArray(), ShutdownDiagnosticTestData.Nonce, ShutdownDiagnosticTestData.Identity()).Length);
            Row(o.Events[1])["ParentInvocationId"] = Guid.NewGuid().ToString("N");
            Assert.ThrowsException<InvalidOperationException>(() => AddInShutdownDiagnosticReceipt.Validate(o.Events.ToArray(), ShutdownDiagnosticTestData.Nonce, ShutdownDiagnosticTestData.Identity()));
        }

        internal static void PublishChain(string output, string nonce)
        {
            var identity = ShutdownDiagnosticTestData.Identity();
            var d = AddInShutdownDiagnostic.Begin(ShutdownDiagnosticTestData.Json.Serialize(new { Version = 1, Nonce = nonce, Identity = identity }), nonce, "Dispose",
                () => identity, ShutdownDiagnosticTestData.Thread, (s, v, t) => AddInShutdownDiagnostic.Publish(output, s, v, t));
            ShutdownDiagnosticTestData.FinishDispose(d); Assert.IsFalse(d.DiagnosticFailed);
        }

        [DataTestMethod, DataRow("valid"), DataRow("missing ready"), DataRow("ready on entry"), DataRow("extra temporary"), DataRow("missing event"), DataRow("oversized"), DataRow("duplicate property"), DataRow("foreign filename"), DataRow("subdirectory")]
        public void PhysicalReaderRequiresClosedCanonicalBoundedFilesAndEveryTerminalReady(string kind)
        {
            using (var directory = new ShutdownDiagnosticTestData.DirectoryScope())
            {
                string nonce = Path.GetFileName(directory.Root); PublishChain(directory.Root, nonce);
                string[] events = Directory.GetFiles(directory.Root, "*.json").OrderBy(path => path, StringComparer.Ordinal).ToArray();
                string terminal = events.Last();
                if (kind == "missing ready") File.Delete(terminal + ".ready");
                if (kind == "ready on entry") File.WriteAllText(events[0] + ".ready", "");
                if (kind == "extra temporary") File.WriteAllText(Path.Combine(directory.Root, "unexpected.tmp"), "partial");
                if (kind == "missing event") File.Delete(events[1]);
                if (kind == "oversized") File.WriteAllText(events[1], new string('x', AddInShutdownDiagnostic.MaximumBytes + 1));
                if (kind == "duplicate property") File.WriteAllText(events[1], File.ReadAllText(events[1]).Replace("\"Version\":1", "\"Version\":1,\"Version\":1"));
                if (kind == "foreign filename") File.Move(events[1], Path.Combine(directory.Root, "foreign.json"));
                if (kind == "subdirectory") Directory.CreateDirectory(Path.Combine(directory.Root, "unrelated"));
                if (kind == "valid")
                {
                    var report = Row(ShutdownDiagnosticTestData.Clone(AddInShutdownDiagnosticReceipt.Read(directory.Root, nonce, ShutdownDiagnosticTestData.Identity())));
                    Assert.AreEqual(false, report["HostExitProven"]); Assert.AreEqual(false, report["NativeRcwReleaseProven"]);
                    Assert.AreEqual("VALID_CLEANUP_OBSERVATIONS_ONLY", report["State"]);
                }
                else Assert.ThrowsException<InvalidOperationException>(() => AddInShutdownDiagnosticReceipt.Read(directory.Root, nonce, ShutdownDiagnosticTestData.Identity()));
            }
        }
    }
}
