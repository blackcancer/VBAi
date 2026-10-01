using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Uses only fake delivery/cleanup and local synthetic files; no Office, COM or UI is activated.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class OfficeCommandContainmentTests
    {
        [TestMethod]
        public void TimeoutKeepsDurableIntentOriginalFailureAndPendingStateWithoutReplay()
        {
            InScratch(root => {
                var state = new OfficeCommandContainment();
                var records = new List<object>();
                string path = Path.Combine(root, "commands.json");
                Action persist = () => File.WriteAllText(path, new JavaScriptSerializer().Serialize(records));
                int sends = 0, retained = 0, nativeCleanup = 0;
                var timeout = new TimeoutException("Synthetic uncertain form property write.");
                var observed = Assert.ThrowsException<TimeoutException>(() => state.Send("set_form_property", new { Command = "set_form_property", Property = "Caption", Value = "Synthetic" }, records.Add, persist, () => {
                    sends++;
                    var before = (IDictionary<string, object>)((object[])new JavaScriptSerializer().DeserializeObject(File.ReadAllText(path)))[0];
                    Assert.AreEqual("set_form_property", before["Command"]);
                    Assert.AreEqual("Caption", ((IDictionary<string, object>)before["Request"])["Property"]);
                    StringAssert.StartsWith((string)before["State"], "Prepared;");
                    Assert.IsTrue(state.Pending);
                    throw timeout;
                }, () => retained++));
                Assert.AreSame(timeout, observed);
                Assert.AreSame(timeout, state.Failure);
                Assert.AreEqual(1, sends); Assert.AreEqual(1, retained);
                Assert.IsTrue(state.Pending); Assert.IsTrue(state.Uncertain);
                Assert.ThrowsException<InvalidOperationException>(() => { state.RequireTerminal(); nativeCleanup++; });
                Assert.ThrowsException<InvalidOperationException>(() => state.Send("set_form_property", new object(), records.Add, persist, () => { sends++; return Reply(true); }, () => retained++));
                Assert.AreEqual(0, nativeCleanup); Assert.AreEqual(1, sends); Assert.AreEqual(1, records.Count);
                var final = (IDictionary<string, object>)records[0];
                Assert.IsNull(final["Response"]);
                StringAssert.Contains((string)final["Error"], timeout.Message);
                Assert.IsTrue(Convert.ToInt64(final["ElapsedMilliseconds"]) >= 0);
                Assert.IsNotNull(final["FinishedUtc"]);
            });
        }

        [DataTestMethod, DataRow("null"), DataRow("missing"), DataRow("wrong-type")]
        public void MissingOrMalformedProtocolResponseRetainsNativeOwnership(string kind)
        {
            var state = new OfficeCommandContainment();
            int retained = 0;
            Assert.ThrowsException<InvalidDataException>(() => state.Send("create_form", new object(), _ => { }, () => { },
                () => kind == "null" ? null : kind == "missing" ? new Dictionary<string, object>() : new Dictionary<string, object> { ["Ok"] = "true" }, () => retained++));
            Assert.IsTrue(state.Pending); Assert.IsTrue(state.Uncertain); Assert.AreEqual(1, retained);
        }

        [DataTestMethod, DataRow(true), DataRow(false)]
        public void TerminalHostResponseRemainsUnchangedAndAllowsIndependentCleanup(bool ok)
        {
            var state = new OfficeCommandContainment();
            var expected = Reply(ok);
            int saves = 0, retained = 0, cleanup = 0;
            var actual = state.Send("save_host_document", new object(), _ => { }, () => saves++, () => expected, () => retained++);
            Assert.AreSame(expected, actual);
            Assert.AreEqual(ok, actual["Ok"]);
            Assert.IsFalse(state.Pending); Assert.IsFalse(state.Uncertain);
            state.RequireTerminal(); cleanup++;
            Assert.AreEqual(1, cleanup); Assert.AreEqual(0, retained); Assert.AreEqual(2, saves);
        }

        [TestMethod]
        public void EvidenceFailureBeforeDispatchPreventsEmissionWithoutMarkingNativeWorkPending()
        {
            var state = new OfficeCommandContainment();
            int sends = 0, retained = 0;
            Assert.ThrowsException<IOException>(() => state.Send("create_form", new object(), _ => { }, () => { throw new IOException("Prepared write failed"); },
                () => { sends++; return Reply(true); }, () => retained++));
            Assert.AreEqual(0, sends); Assert.AreEqual(0, retained);
            Assert.IsFalse(state.Pending); Assert.IsFalse(state.Uncertain);
            state.RequireTerminal();
        }

        [TestMethod]
        public void DeliveryAndFinalEvidenceErrorsAreBothRetainedWithoutAnotherSend()
        {
            var state = new OfficeCommandContainment();
            int writes = 0, sends = 0;
            var primary = new TimeoutException("Synthetic response timeout");
            var evidence = new IOException("Synthetic result write failure");
            var observed = Assert.ThrowsException<AggregateException>(() => state.Send("set_project_property", new object(), _ => { },
                () => { if (++writes == 2) throw evidence; }, () => { sends++; throw primary; }, () => { }));
            Assert.AreSame(primary, observed.InnerExceptions[0]); Assert.AreSame(evidence, observed.InnerExceptions[1]);
            Assert.AreSame(observed, state.Failure);
            Assert.AreEqual(1, sends); Assert.IsTrue(state.Pending); Assert.IsTrue(state.Uncertain);
        }

        [TestMethod]
        public void FixtureTimeoutPreservesOriginalErrorAndRefusesFakeNativeCloseQuitResetAndReopen()
        {
            InScratch(root => {
                var fixture = (OfficeVbeFixture)Activator.CreateInstance(typeof(OfficeVbeFixture), true);
                SetProperty(fixture, "Root", root); SetProperty(fixture, "Kind", "Word"); SetProperty(fixture, "ProcessId", 42);
                SetProperty(fixture, "DocumentPath", Path.Combine(root, "Synthetic.docm"));
                var application = new FakeApplication(); var document = new FakeDocument();
                SetField(fixture, "owned", true); SetField(fixture, "application", application); SetField(fixture, "document", document);
                var primary = new TimeoutException("Synthetic emitted native mutation timeout");
                int sends = 0;
                fixture.Dispatch = (pid, request) => {
                    sends++;
                    Assert.AreEqual(42, pid);
                    var before = Read(Path.Combine(root, "adapter-only-progress.json"));
                    Assert.AreEqual("set_form_property", ((IDictionary<string, object>)((object[])before["Steps"])[0])["Command"]);
                    throw primary;
                };
                Assert.AreSame(primary, Assert.ThrowsException<TimeoutException>(() => fixture.Response("set_form_property", "Property", "Caption", "Value", "Synthetic")));
                var cleanup = Assert.ThrowsException<AggregateException>(() => fixture.Dispose());
                Assert.AreSame(primary, cleanup.InnerExceptions[0]);
                Assert.IsInstanceOfType<AssertFailedException>(cleanup.InnerExceptions[1]);
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Response("debug_global", "Action", "reset"));
                Assert.ThrowsException<InvalidOperationException>(() => fixture.SaveNative());
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Reopen());
                Assert.AreEqual(1, sends); Assert.AreEqual(0, application.QuitCount); Assert.AreEqual(0, document.CloseCount);
                Assert.AreSame(application, GetField(fixture, "application")); Assert.AreSame(document, GetField(fixture, "document"));
                var final = Read(Path.Combine(root, "qualification.json"));
                Assert.AreEqual(true, final["CommandPending"]); Assert.AreEqual(true, final["DeliveryUncertain"]);
                Assert.AreEqual("set_form_property", final["PendingCommand"]);
            });
        }

        [TestMethod]
        public void FixtureFinalWriteFailureCannotMaskTheOriginalTimeoutOrCleanupRefusal()
        {
            InScratch(root => {
                var fixture = (OfficeVbeFixture)Activator.CreateInstance(typeof(OfficeVbeFixture), true);
                SetProperty(fixture, "Root", root); SetProperty(fixture, "Kind", "Word");
                var primary = new TimeoutException("Synthetic command deadline");
                fixture.Dispatch = (pid, request) => { throw primary; };
                Assert.AreSame(primary, Assert.ThrowsException<TimeoutException>(() => fixture.Response("set_form_property")));
                using (var locked = new FileStream(Path.Combine(root, "qualification.json"), FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    var failures = Assert.ThrowsException<AggregateException>(() => fixture.Dispose());
                    var refusal = (AggregateException)failures.InnerExceptions[0];
                    Assert.AreSame(primary, refusal.InnerExceptions[0]);
                    Assert.IsInstanceOfType<AssertFailedException>(refusal.InnerExceptions[1]);
                    Assert.IsInstanceOfType<IOException>(failures.InnerExceptions[1]);
                }
                var progress = Read(Path.Combine(root, "adapter-only-progress.json"));
                Assert.AreEqual(true, progress["CommandPending"]); Assert.AreEqual(true, progress["DeliveryUncertain"]);
            });
        }

        [TestMethod]
        public void EvidenceBoundRefusesTheNextDispatchBeforeItCouldBecomeUnrecorded()
        {
            var state = new OfficeCommandContainment();
            int sends = 0;
            Func<IDictionary<string, object>> dispatch = () => { sends++; return Reply(true); };
            for (int i = 0; i < 512; i++) state.Send("status", new object(), _ => { }, () => { }, dispatch, () => Assert.Fail("Terminal reads need no retention."));
            Assert.ThrowsException<InvalidOperationException>(() => state.Send("create_form", new object(), _ => { }, () => { }, dispatch, () => { }));
            Assert.AreEqual(512, sends);
            Assert.IsFalse(state.Pending); Assert.IsFalse(state.Uncertain);
        }

        private static IDictionary<string, object> Reply(bool ok) => new Dictionary<string, object> { ["Ok"] = ok, ["Error"] = ok ? null : "Synthetic native refusal", ["Data"] = null };
        private static IDictionary<string, object> Read(string path) => (IDictionary<string, object>)new JavaScriptSerializer().DeserializeObject(File.ReadAllText(path));
        private static void SetProperty(object target, string name, object value) => target.GetType().GetProperty(name, BindingFlags.Instance | BindingFlags.NonPublic).SetValue(target, value);
        private static void SetField(object target, string name, object value) => target.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).SetValue(target, value);
        private static object GetField(object target, string name) => target.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(target);
        private static void InScratch(Action<string> action)
        {
            string root = Path.Combine(Path.GetTempPath(), "VBAi-OfficeContainment-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try { action(root); }
            finally { Directory.Delete(root, true); }
        }

        public sealed class FakeApplication { public int QuitCount; public void Quit(int option) { QuitCount++; } }
        public sealed class FakeDocument { public int CloseCount; public void Close(int option) { CloseCount++; } }
    }
}
