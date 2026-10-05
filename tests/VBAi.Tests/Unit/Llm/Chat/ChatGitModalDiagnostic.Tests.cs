using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Exercises the actual opt-in transaction, deferred dispatch evidence, strict identity and atomic publication without an Office host.</summary>
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class ChatGitModalDiagnosticTests
    {
        private const string Nonce = "3dbce6bfb1dc49cb9b6c14c450b5e324";
        private static readonly JavaScriptSerializer Json = new JavaScriptSerializer();
        private static ChatGitModalDiagnostic.Identity Identity()
        { return new ChatGitModalDiagnostic.Identity { DocumentPath = @"E:\Qualification\Disposable.docm", ProcessId = 71,
            ProcessStartedUtc = "2026-10-05T16:19:22.8081032Z", ThreadId = 73, ChatHandle = 75, RootHandle = 77,
            ProductMvid = "c4b7e1e3-dbf5-4dce-8498-1c0b7b45c973", ProductSha256 = new string('A', 64) }; }
        private static string Request(ChatGitModalDiagnostic.Identity identity = null)
        { return Json.Serialize(new { Version = 1, Nonce, Identity = identity ?? Identity() }); }
        private sealed class Observation
        {
            internal ChatGitModalDiagnostic.Identity Current = Identity();
            internal readonly List<object> Chain = new List<object>();
            internal readonly List<string> Attempts = new List<string>();
            internal string ThrowPhase;
            internal Action Callback;
            internal int Posts;
            internal ChatGitModalDiagnostic Begin()
            { return ChatGitModalDiagnostic.Begin(Request(Current), Nonce, () => Current, Write); }
            internal void Write(string phase, object value)
            { Attempts.Add(phase); if (phase == ThrowPhase) throw new IOException("write:" + phase); Chain.Add(Json.DeserializeObject(Json.Serialize(value))); }
            internal void Post(Action callback) { Posts++; Callback = callback; }
        }

        [STATestMethod]
        public void DisabledDiagnosticDoesNotCreateAWindowOrReadAnyManifest()
        {
            string prior = Environment.GetEnvironmentVariable(ChatGitModalDiagnostic.EnvironmentName);
            try { Environment.SetEnvironmentVariable(ChatGitModalDiagnostic.EnvironmentName, null);
                using (var chat = new Form()) { Assert.IsNull(ChatGitModalDiagnostic.BeginFromEnvironment(chat, null)); Assert.IsFalse(chat.IsHandleCreated); } }
            finally { Environment.SetEnvironmentVariable(ChatGitModalDiagnostic.EnvironmentName, prior); }
        }

        [TestMethod]
        public void CompleteTransactionRequiresOneLaterCallbackAndValidatesAllOriginalObservedPhases()
        {
            var o = new Observation(); var d = o.Begin(); int show = 0, dispose = 0;
            d.RunModal(() => show++, () => dispose++); d.SchedulePostHandler(o.Post);
            Assert.IsFalse(d.Completed); Assert.AreEqual(1, o.Posts); Assert.AreEqual(1, show); Assert.AreEqual(1, dispose);
            CollectionAssert.AreEqual(ChatGitModalDiagnostic.Phases.Take(4).ToArray(), o.Attempts.ToArray());
            o.Callback(); Assert.IsTrue(d.Completed); Assert.IsFalse(d.Failed);
            ChatGitDiagnosticReceipt.Validate(o.Chain.ToArray(), Nonce, Identity());
            CollectionAssert.AreEqual(ChatGitModalDiagnostic.Phases, o.Attempts.ToArray());
        }

        [DataTestMethod]
        [DataRow("Show")][DataRow("Dispose")][DataRow("Both")]
        public void NativeModalAndDisposalFailuresRemainPrimaryAndNeverProduceADeferredSuccess(string which)
        {
            var o = new Observation(); var d = o.Begin(); int shows = 0, disposals = 0;
            var show = new InvalidOperationException("original show"); var dispose = new InvalidOperationException("original dispose");
            Exception error = null;
            try { d.RunModal(() => { shows++; if (which != "Dispose") throw show; }, () => { disposals++; if (which != "Show") throw dispose; }); }
            catch (Exception actual) { error = actual; }
            Assert.IsNotNull(error); Assert.AreEqual(1, shows); Assert.AreEqual(1, disposals);
            if (which == "Both") { var aggregate = error as AggregateException; Assert.IsNotNull(aggregate); CollectionAssert.AreEqual(new Exception[] { show, dispose }, aggregate.InnerExceptions.ToArray()); }
            else Assert.AreSame(which == "Show" ? show : dispose, error);
            d.Fail(error); d.SchedulePostHandler(o.Post); Assert.AreEqual(0, o.Posts); Assert.IsFalse(d.Completed); Assert.IsTrue(d.Failed);
            Assert.IsFalse(o.Attempts.Contains("PostHandlerCallbackObserved"));
        }

        [DataTestMethod]
        [DataRow("Started")][DataRow("ShowModalReturned")][DataRow("DisposeReturned")][DataRow("PostHandlerIntent")][DataRow("PostHandlerCallbackObserved")]
        public void EveryReceiptPublicationFailurePreventsSuccessWithoutRetryingShowDisposalOrPublication(string phase)
        {
            var o = new Observation { ThrowPhase = phase }; int shows = 0, disposals = 0;
            if (phase == "Started") { Assert.ThrowsException<IOException>(() => o.Begin()); CollectionAssert.AreEqual(new[] { "Started" }, o.Attempts); return; }
            var d = o.Begin();
            try { d.RunModal(() => shows++, () => disposals++); }
            catch (Exception error) { d.Fail(error); }
            d.SchedulePostHandler(o.Post); if (o.Callback != null) o.Callback();
            Assert.IsFalse(d.Completed); Assert.IsTrue(d.Failed); Assert.AreEqual(1, shows); Assert.AreEqual(1, disposals);
            Assert.AreEqual(1, o.Attempts.Count(value => value == phase));
        }

        [TestMethod]
        public void FailureWriterCannotMaskTheOriginalModalErrorOrCreateAnotherPublicationAttempt()
        {
            var o = new Observation { ThrowPhase = "Failed" }; var d = o.Begin(); var original = new InvalidOperationException("primary");
            Assert.ThrowsException<InvalidOperationException>(() => d.RunModal(() => { throw original; }, () => { }));
            d.Fail(original); d.Fail(original); var errors = ((AggregateException)d.Error).InnerExceptions;
            Assert.AreSame(original, errors[0]); Assert.IsInstanceOfType<IOException>(errors[1]);
            Assert.AreEqual(1, o.Attempts.Count(value => value == "Failed")); Assert.IsFalse(d.Completed);
        }

        [DataTestMethod][DataRow(false)][DataRow(true)]
        public void FailedQueuePublicationCannotBeRehabilitatedByAQueuedCallback(bool acceptedBeforeThrow)
        {
            var o = new Observation(); var d = o.Begin(); d.RunModal(() => { }, () => { });
            d.SchedulePostHandler(callback => { if (acceptedBeforeThrow) o.Callback = callback; throw new InvalidOperationException("queue failed"); });
            if (o.Callback != null) o.Callback(); Assert.IsTrue(d.Failed); Assert.IsFalse(d.Completed);
            Assert.IsFalse(o.Attempts.Contains("PostHandlerCallbackObserved"));
        }

        [TestMethod]
        public void SynchronousCallbackIsNotALaterDispatcherObservationAndMissingCallbackNeverCompletes()
        {
            var o = new Observation(); var d = o.Begin(); d.RunModal(() => { }, () => { }); d.SchedulePostHandler(callback => callback());
            Assert.IsTrue(d.Failed); Assert.IsFalse(d.Completed);
            var missing = new Observation(); var pending = missing.Begin(); pending.RunModal(() => { }, () => { }); pending.SchedulePostHandler(missing.Post);
            Assert.IsFalse(pending.Completed); Assert.IsFalse(pending.Failed); Assert.AreEqual(4, missing.Chain.Count);
        }

        [TestMethod]
        public void RepeatedBodyQueueAndCallbackRefuseAndIdentitySnapshotCannotBeChangedThroughTheCaptureObject()
        {
            var o = new Observation(); var d = o.Begin(); d.RunModal(() => { }, () => { });
            Assert.ThrowsException<InvalidOperationException>(() => d.RunModal(() => Assert.Fail("No second show"), () => Assert.Fail("No second dispose")));
            d.SchedulePostHandler(o.Post); o.Current.ThreadId++; o.Callback(); Assert.IsTrue(d.Failed); Assert.IsFalse(d.Completed);
            d.SchedulePostHandler(o.Post); Assert.AreEqual(1, o.Posts);
            var complete = new Observation(); var done = complete.Begin(); done.RunModal(() => { }, () => { }); done.SchedulePostHandler(complete.Post); complete.Callback();
            complete.Callback(); Assert.IsTrue(done.Failed); Assert.IsFalse(done.Completed); Assert.AreEqual(1, complete.Attempts.Count(phase => phase == "PostHandlerCallbackObserved"));
        }

        [DataTestMethod]
        [DataRow("DocumentPath")][DataRow("ProcessId")][DataRow("ProcessStartedUtc")][DataRow("ThreadId")]
        [DataRow("ChatHandle")][DataRow("RootHandle")][DataRow("ProductMvid")][DataRow("ProductSha256")]
        public void EveryChangedIdentityRefusesBeforeStartedAndAlsoInTheDeferredCallback(string field)
        {
            var expected = Identity(); var changed = Json.Deserialize<ChatGitModalDiagnostic.Identity>(Json.Serialize(expected));
            var property = typeof(ChatGitModalDiagnostic.Identity).GetProperty(field); object value = property.GetValue(changed);
            if (value is int) property.SetValue(changed, (int)value + 1);
            else if (value is uint) property.SetValue(changed, (uint)value + 1);
            else if (value is long) property.SetValue(changed, (long)value + 1);
            else property.SetValue(changed, field == "DocumentPath" ? @"E:\Qualification\Other.docm" : field == "ProcessStartedUtc" ? "2026-10-05T16:19:23.8081032Z" : field == "ProductMvid" ? Guid.NewGuid().ToString("D") : new string('B', 64));
            int writes = 0;
            Assert.ThrowsException<InvalidOperationException>(() => ChatGitModalDiagnostic.Begin(Request(expected), Nonce, () => changed, (phase, row) => writes++)); Assert.AreEqual(0, writes);
            var o = new Observation(); var d = o.Begin(); d.RunModal(() => { }, () => { }); d.SchedulePostHandler(o.Post); o.Current = changed; o.Callback();
            Assert.IsTrue(d.Failed); Assert.IsFalse(d.Completed); Assert.IsFalse(o.Attempts.Contains("PostHandlerCallbackObserved"));
        }

        [DataTestMethod]
        [DataRow("Version")][DataRow("Nonce")][DataRow("Extra")][DataRow("Missing")][DataRow("Integer")][DataRow("RelativePath")]
        public void InvalidManifestRefusesBeforeAnyReceiptPublication(string mutation)
        {
            var request = (IDictionary<string, object>)Json.DeserializeObject(Request());
            if (mutation == "Version") request["Version"] = 2;
            else if (mutation == "Nonce") request["Nonce"] = Guid.NewGuid().ToString("N");
            else if (mutation == "Extra") request["Unexpected"] = true;
            else if (mutation == "Missing") ((IDictionary<string, object>)request["Identity"]).Remove("RootHandle");
            else if (mutation == "Integer") ((IDictionary<string, object>)request["Identity"])["ThreadId"] = "73";
            else ((IDictionary<string, object>)request["Identity"])["DocumentPath"] = "Disposable.docm";
            int writes = 0; Assert.ThrowsException<InvalidOperationException>(() => ChatGitModalDiagnostic.Begin(Json.Serialize(request), Nonce, Identity, (phase, row) => writes++)); Assert.AreEqual(0, writes);
        }

        [DataTestMethod]
        [DataRow("Relative")][DataRow("Unc")][DataRow("Parent")][DataRow("Name")][DataRow("Nonce")][DataRow("Traversal")]
        public void ManifestPathCannotEscapeTheFixedOwnedTemporaryChild(string mutation)
        {
            string temp = Path.GetTempPath(), root = Path.Combine(temp, ChatGitModalDiagnostic.DirectoryName, Nonce), path = Path.Combine(root, "request.json");
            Assert.AreEqual(root, ChatGitModalDiagnostic.RequireManifestPath(path, temp));
            if (mutation == "Relative") path = "request.json";
            else if (mutation == "Unc") path = @"\\server\share\request.json";
            else if (mutation == "Parent") path = Path.Combine(temp, Nonce, "request.json");
            else if (mutation == "Name") path = Path.Combine(root, "other.json");
            else if (mutation == "Nonce") path = Path.Combine(temp, ChatGitModalDiagnostic.DirectoryName, "not-owned", "request.json");
            else path = Path.Combine(root, "..", Nonce, "request.json");
            Assert.ThrowsException<ArgumentException>(() => ChatGitModalDiagnostic.RequireManifestPath(path, temp));
        }

        [TestMethod]
        public void RealAtomicWriterConsumesOneInvocationAndReadyNeverOverwritesPublishedBytes()
        {
            string root = Path.Combine(Path.GetTempPath(), "vbai-chat-diagnostic-unit-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(root);
            try
            {
                ChatGitModalDiagnostic.Publish(root, "Started", new { Value = "évidence" });
                string path = Path.Combine(root, "Started.json"), original = File.ReadAllText(path);
                Assert.IsTrue(File.Exists(path + ".ready")); Assert.IsFalse(File.Exists(path + ".tmp"));
                Assert.AreEqual("évidence", ((IDictionary<string, object>)Json.DeserializeObject(original))["Value"]);
                Assert.ThrowsException<IOException>(() => ChatGitModalDiagnostic.Publish(root, "Started", new { Value = "changed" }));
                Assert.AreEqual(original, File.ReadAllText(path));
                Assert.ThrowsException<ArgumentException>(() => ChatGitModalDiagnostic.Publish(root, "Unknown", new { Value = 1 }));
            }
            finally { foreach (string path in Directory.GetFiles(root)) File.Delete(path); Directory.Delete(root); }
        }
    }
}