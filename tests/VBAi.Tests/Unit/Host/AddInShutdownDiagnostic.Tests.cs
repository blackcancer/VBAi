using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Shared managed fixtures for cleanup-observation and receipt mirrors; never launches a native host.</summary>
    internal static class ShutdownDiagnosticTestData
    {
        internal const string Nonce = "dd75799f5ff442b092a1b1fe6b3e53cf";
        internal static readonly JavaScriptSerializer Json = new JavaScriptSerializer();
        internal static AddInShutdownDiagnostic.Identity Identity()
        {
            return new AddInShutdownDiagnostic.Identity
            {
                ProcessId = 71,
                ProcessStartedUtc = "2026-10-05T17:27:03.6593947Z",
                HostImagePath = @"C:\Qualification\EXCEL.EXE",
                ProductPath = @"C:\Qualification\VBAi.dll",
                ProductMvid = "f855d463-0d5e-48f3-8c45-1a35db5a1025",
                ProductSha256 = new string('A', 64),
                ThreadId = 73
            };
        }
        internal static AddInShutdownDiagnostic.ThreadIdentity Thread()
        { return new AddInShutdownDiagnostic.ThreadIdentity { ManagedThreadId = 7, NativeThreadId = 73, Apartment = "STA" }; }
        internal static string Request(AddInShutdownDiagnostic.Identity identity = null)
        { return Json.Serialize(new { Version = 1, Nonce, Identity = identity ?? Identity() }); }
        internal static object Clone(object value) { return Json.DeserializeObject(Json.Serialize(value)); }
        internal sealed class Observation
        {
            internal readonly List<object> Events = new List<object>();
            internal AddInShutdownDiagnostic.Identity Actual = Identity();
            internal AddInShutdownDiagnostic.ThreadIdentity Caller = Thread();
            internal int PublishCalls, ThrowAt;
            internal readonly IOException PublicationFailure = new IOException("synthetic publication failure");
            internal AddInShutdownDiagnostic Begin(string entry = "Dispose")
            { return AddInShutdownDiagnostic.Begin(Request(Actual), Nonce, entry, () => Actual, () => Caller, Publish); }
            internal void Publish(int sequence, object value, bool terminal)
            { PublishCalls++; if (PublishCalls == ThrowAt) throw PublicationFailure; Events.Add(Clone(value)); }
        }
        internal static readonly string[] Cleared = { "editorNavigation", "editorWorkspace", "modernEditor", "testExplorerWindow", "nativeTestWindow",
            "nativeTestControl", "testExplorerService", "crashReporter", "menu", "chat", "nativeChatWindow", "nativeChatControl", "addIn", "server", "dispatcher", "vbe" };
        internal static void FinishDispose(AddInShutdownDiagnostic d, bool caughtCloseFault = false, Action<AddInShutdownDiagnostic> duringClose = null)
        {
            d.Enter("VbeNativeTheme.Disconnect"); d.Returned("VbeNativeTheme.Disconnect");
            foreach (string field in Cleared)
            {
                if (field == "crashReporter") { d.Enter("StopUpdateCheck"); d.Returned("StopUpdateCheck"); }
                if (field == "nativeChatWindow") { d.Enter("nativeChatWindow.Close"); duringClose?.Invoke(d); if (caughtCloseFault) d.CaughtFault("nativeChatWindow.Close", null); else d.Returned("nativeChatWindow.Close"); }
                d.ManagedReferenceCleared(field);
            }
            d.Complete();
        }
        internal sealed class DirectoryScope : IDisposable
        {
            internal readonly string Root = Path.Combine(Path.GetTempPath(), AddInShutdownDiagnostic.DirectoryName, Guid.NewGuid().ToString("N"));
            internal DirectoryScope() { Directory.CreateDirectory(Root); }
            public void Dispose()
            { AddInShutdownDiagnostic.RequireRoot(Root, Path.GetTempPath()); Directory.Delete(Root, true); }
        }
    }

    /// <summary>Mirrors strict opt-in admission, bounded no-throw observation, and atomic publication.</summary>
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class AddInShutdownDiagnosticTests
    {
        [DataTestMethod, DataRow(null), DataRow(""), DataRow(" ")]
        public void DisabledObservationReturnsWithoutAnyRuntimeSession(string setting)
        {
            string prior = Environment.GetEnvironmentVariable(AddInShutdownDiagnostic.EnvironmentName);
            try { Environment.SetEnvironmentVariable(AddInShutdownDiagnostic.EnvironmentName, setting); Assert.IsNull(AddInShutdownDiagnostic.TryBeginFromEnvironment("Dispose")); }
            finally { Environment.SetEnvironmentVariable(AddInShutdownDiagnostic.EnvironmentName, prior); }
        }

        [DataTestMethod]
        [DataRow("ProcessId")]
        [DataRow("ProcessStartedUtc")]
        [DataRow("HostImagePath")]
        [DataRow("ProductPath")]
        [DataRow("ProductMvid")]
        [DataRow("ProductSha256")]
        [DataRow("ThreadId")]
        public void AdmissionComparesEveryRequestedMemberAgainstIndependentActualIdentity(string member)
        {
            var identity = ShutdownDiagnosticTestData.Identity();
            if (member == "ProcessId") identity.ProcessId++;
            if (member == "ProcessStartedUtc") identity.ProcessStartedUtc = "2026-10-05T17:27:04.6593947Z";
            if (member == "HostImagePath") identity.HostImagePath = @"C:\Foreign\EXCEL.EXE";
            if (member == "ProductPath") identity.ProductPath = @"C:\Foreign\VBAi.dll";
            if (member == "ProductMvid") identity.ProductMvid = Guid.NewGuid().ToString("D");
            if (member == "ProductSha256") identity.ProductSha256 = new string('B', 64);
            if (member == "ThreadId") identity.ThreadId++;
            int calls = 0;
            Assert.ThrowsException<InvalidOperationException>(() => AddInShutdownDiagnostic.Begin(ShutdownDiagnosticTestData.Request(identity), ShutdownDiagnosticTestData.Nonce,
                "Dispose", ShutdownDiagnosticTestData.Identity, ShutdownDiagnosticTestData.Thread, (s, v, t) => calls++));
            Assert.AreEqual(0, calls);
        }

        [DataTestMethod]
        [DataRow("null")]
        [DataRow("truncated")]
        [DataRow("version")]
        [DataRow("missing")]
        [DataRow("extra")]
        [DataRow("duplicate")]
        [DataRow("nonce")]
        [DataRow("string number")]
        [DataRow("oversized")]
        [DataRow("noncanonical")]
        public void AdmissionRejectsMalformedDuplicateAndNoncanonicalRequestsBeforePublication(string kind)
        {
            string request = ShutdownDiagnosticTestData.Request();
            var fields = (Dictionary<string, object>)ShutdownDiagnosticTestData.Clone(ShutdownDiagnosticTestData.Json.DeserializeObject(request));
            if (kind == "version") fields["Version"] = 2;
            if (kind == "missing") fields.Remove("Identity");
            if (kind == "extra") fields["Extra"] = true;
            if (kind == "nonce") fields["Nonce"] = Guid.NewGuid().ToString("N");
            if (kind == "string number") ((Dictionary<string, object>)fields["Identity"])["ProcessId"] = "71";
            request = ShutdownDiagnosticTestData.Json.Serialize(fields);
            if (kind == "null") request = null;
            if (kind == "truncated") request = "{";
            if (kind == "duplicate") request = request.Replace("\"Version\":1", "\"Version\":1,\"Version\":1");
            if (kind == "oversized") request = new string(' ', AddInShutdownDiagnostic.MaximumBytes + 1);
            if (kind == "noncanonical") request = " " + request;
            int calls = 0; Exception failure = null;
            try { AddInShutdownDiagnostic.Begin(request, ShutdownDiagnosticTestData.Nonce, "Dispose", ShutdownDiagnosticTestData.Identity, ShutdownDiagnosticTestData.Thread, (s, v, t) => calls++); }
            catch (Exception error) { failure = error; }
            Assert.IsNotNull(failure); Assert.AreEqual(0, calls);
        }

        [TestMethod]
        public void ChildInvocationAndCaughtFaultPreserveNormalCallbackReturnWithoutClaimingNativeRelease()
        {
            var o = new ShutdownDiagnosticTestData.Observation(); var parent = o.Begin("OnDisconnection");
            var child = parent.Child("Dispose"); ShutdownDiagnosticTestData.FinishDispose(child, true); parent.Complete();
            Assert.IsFalse(parent.DiagnosticFailed); Assert.AreNotEqual(parent.InvocationId, child.InvocationId);
            var result = AddInShutdownDiagnosticReceipt.Validate(o.Events.ToArray(), ShutdownDiagnosticTestData.Nonce, ShutdownDiagnosticTestData.Identity());
            Assert.AreEqual(2, result.Length); Assert.AreEqual("Returned", result.Single(r => r.Id == parent.InvocationId).Outcome);
            Assert.AreEqual("Fault", result.Single(r => r.Id == child.InvocationId).Stages["nativeChatWindow.Close"]);
        }

        [TestMethod]
        public void ReentrantCallbackHasTheActiveDisposeAsItsParentAndPreservesBothOriginalInvocations()
        {
            var o = new ShutdownDiagnosticTestData.Observation(); var outer = o.Begin("OnBeginShutdown"); var dispose = outer.Child("Dispose");
            ShutdownDiagnosticTestData.FinishDispose(dispose, false, active =>
            { var callback = active.Child("OnDisconnection"); ShutdownDiagnosticTestData.FinishDispose(callback.Child("Dispose")); callback.Complete(); });
            outer.Complete();
            Assert.IsFalse(outer.DiagnosticFailed);
            var result = AddInShutdownDiagnosticReceipt.Validate(o.Events.ToArray(), ShutdownDiagnosticTestData.Nonce, ShutdownDiagnosticTestData.Identity());
            Assert.AreEqual(4, result.Length); Assert.AreEqual(dispose.InvocationId, result.Single(r => r.EntryPoint == "OnDisconnection").Parent);
            Assert.IsTrue(result.All(r => r.Outcome == "Returned"));
        }

        [DataTestMethod, DataRow(1), DataRow(2), DataRow(3), DataRow(4), DataRow(5), DataRow(6)]
        public void FirstPublicationFaultStopsAllFurtherPublicationAndNeverThrowsIntoCleanup(int at)
        {
            var o = new ShutdownDiagnosticTestData.Observation { ThrowAt = at }; var d = o.Begin();
            d.Enter("VbeNativeTheme.Disconnect"); d.Returned("VbeNativeTheme.Disconnect");
            d.ManagedReferenceCleared("editorNavigation"); d.Enter("StopUpdateCheck"); d.Returned("StopUpdateCheck");
            d.Complete(); d.Fault(new Exception("original error")); d.ManagedReferenceCleared("vbe");
            Assert.IsTrue(d.DiagnosticFailed); Assert.AreSame(o.PublicationFailure, d.PublicationError);
            Assert.AreEqual(at, o.PublishCalls); Assert.AreEqual(at - 1, o.Events.Count);
            Assert.IsFalse(o.Events.Cast<Dictionary<string, object>>().Any(row => row["Stage"] == null && (string)row["Phase"] != "Entry"));
        }

        [DataTestMethod, DataRow("return without entry"), DataRow("double entry"), DataRow("unresolved completion"), DataRow("cleared during operation"), DataRow("post terminal")]
        public void InvalidObservationStatePermanentlyRefusesAFalseTerminal(string kind)
        {
            var o = new ShutdownDiagnosticTestData.Observation(); var d = o.Begin();
            if (kind == "return without entry") d.Returned("missing");
            else if (kind == "post terminal") { d.Complete(); d.Enter("later"); }
            else { d.Enter("StopUpdateCheck"); if (kind == "double entry") d.Enter("second"); if (kind == "unresolved completion") d.Complete(); if (kind == "cleared during operation") d.ManagedReferenceCleared("vbe"); }
            int calls = o.PublishCalls; d.Complete(); d.Fault(new Exception());
            Assert.IsTrue(d.DiagnosticFailed); Assert.AreEqual(calls, o.PublishCalls);
        }

        [DataTestMethod, DataRow("native"), DataRow("managed"), DataRow("apartment"), DataRow("missing")]
        public void ChangedActualCallerInvalidatesTraceWithoutChangingOriginalOperation(string kind)
        {
            var o = new ShutdownDiagnosticTestData.Observation(); var d = o.Begin();
            if (kind == "native") o.Caller.NativeThreadId++;
            if (kind == "managed") o.Caller.ManagedThreadId = 0;
            if (kind == "apartment") o.Caller.Apartment = "MTA";
            if (kind == "missing") o.Caller = null;
            int originalOperations = 0; d.Enter("StopUpdateCheck"); originalOperations++; d.Returned("StopUpdateCheck"); d.Complete();
            Assert.AreEqual(1, originalOperations); Assert.IsTrue(d.DiagnosticFailed); Assert.AreEqual(1, o.PublishCalls);
        }

        [TestMethod]
        public void CapturedIdentityCannotBeChangedByMutatingTheOriginalCaptureObject()
        {
            var o = new ShutdownDiagnosticTestData.Observation(); var d = o.Begin(); o.Actual.ProcessId = 999;
            d.Enter("StopUpdateCheck"); d.Returned("StopUpdateCheck"); d.Complete();
            foreach (Dictionary<string, object> row in o.Events) Assert.AreEqual(71, ((Dictionary<string, object>)row["Identity"])["ProcessId"]);
        }

        [DataTestMethod, DataRow("invocations"), DataRow("events")]
        public void BoundedSessionStopsBeforePublishingBeyondItsPreparedMaximum(string kind)
        {
            var o = new ShutdownDiagnosticTestData.Observation(); var root = o.Begin("OnBeginShutdown");
            if (kind == "invocations")
            {
                for (int i = 1; i < AddInShutdownDiagnostic.MaximumInvocations; i++) { var child = root.Child("Dispose"); Assert.IsNotNull(child); child.Complete(); }
                Assert.IsNull(root.Child("Dispose")); Assert.IsTrue(root.DiagnosticFailed);
            }
            else
            {
                for (int i = 0; i < AddInShutdownDiagnostic.MaximumEvents; i++) { root.Enter("stage" + i); root.Returned("stage" + i); }
                Assert.IsTrue(root.DiagnosticFailed); Assert.AreEqual(AddInShutdownDiagnostic.MaximumEvents, o.Events.Count);
            }
        }

        [DataTestMethod, DataRow("temporary"), DataRow("final"), DataRow("ready")]
        public void AtomicPublicationCollisionPreservesExistingBytesAndDoesNotRetry(string target)
        {
            using (var directory = new ShutdownDiagnosticTestData.DirectoryScope())
            {
                string path = Path.Combine(directory.Root, "000001.json"), collision = path + (target == "temporary" ? ".tmp" : target == "ready" ? ".ready" : "");
                File.WriteAllText(collision, "unchanged", new UTF8Encoding(false));
                Assert.ThrowsException<IOException>(() => AddInShutdownDiagnostic.Publish(directory.Root, 1, new { Marker = true }, true));
                Assert.AreEqual("unchanged", File.ReadAllText(collision));
            }
        }

        [TestMethod]
        public void ClosedAtomicPublicationHasNoBomAndMarksOnlyTerminalFilesReady()
        {
            using (var directory = new ShutdownDiagnosticTestData.DirectoryScope())
            {
                AddInShutdownDiagnostic.Publish(directory.Root, 1, new { Marker = true }, false);
                AddInShutdownDiagnostic.Publish(directory.Root, 2, new { Marker = true }, true);
                Assert.AreEqual((byte)'{', File.ReadAllBytes(Path.Combine(directory.Root, "000001.json"))[0]);
                Assert.IsFalse(File.Exists(Path.Combine(directory.Root, "000001.json.ready")));
                Assert.IsTrue(File.Exists(Path.Combine(directory.Root, "000002.json.ready")));
                Assert.AreEqual(0, Directory.GetFiles(directory.Root, "*.tmp").Length);
                Assert.ThrowsException<InvalidOperationException>(() => AddInShutdownDiagnostic.Publish(directory.Root, 3, new { Huge = new string('x', AddInShutdownDiagnostic.MaximumBytes) }, true));
                Assert.IsFalse(File.Exists(Path.Combine(directory.Root, "000003.json.tmp")));
            }
        }

        [TestMethod]
        public void EveryExistingLinkedPathComponentIsRefusedIncludingTheRequestFile()
        {
            string path = @"C:\Qualification\Diagnostic\request-71.json";
            var observed = new List<string>(); AddInShutdownDiagnostic.RequireNoReparse(path, current => { observed.Add(current); return FileAttributes.Normal; });
            foreach (string link in observed)
                Assert.ThrowsException<InvalidOperationException>(() => AddInShutdownDiagnostic.RequireNoReparse(path,
                    current => current == link ? FileAttributes.ReparsePoint : FileAttributes.Normal));
        }

        [DataTestMethod, DataRow("outside"), DataRow("traversal"), DataRow("nonce"), DataRow("relative"), DataRow("unc")]
        public void RootAdmissionRefusesAnyUnownedOrNoncanonicalDirectory(string kind)
        {
            using (var directory = new ShutdownDiagnosticTestData.DirectoryScope())
            {
                string root = directory.Root;
                if (kind == "outside") root = Path.Combine(Path.GetTempPath(), ShutdownDiagnosticTestData.Nonce);
                if (kind == "traversal") root += @"\..\" + Path.GetFileName(root);
                if (kind == "nonce") root = Path.Combine(Path.GetDirectoryName(root), "not-owned");
                if (kind == "relative") root = "relative";
                if (kind == "unc") root = @"\\localhost\C$\" + ShutdownDiagnosticTestData.Nonce;
                Assert.ThrowsException<InvalidOperationException>(() => AddInShutdownDiagnostic.RequireRoot(root, Path.GetTempPath()));
            }
        }
    }
}
