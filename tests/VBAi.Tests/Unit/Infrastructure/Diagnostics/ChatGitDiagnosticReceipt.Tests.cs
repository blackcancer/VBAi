using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Exercises the complete first-close authorization oracle and strict deadline without any native host or real wait.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class ChatGitDiagnosticReceiptTests
    {
        private const string Nonce = "47e3e856cfde40e08cdde96e7027e568";
        private static readonly JavaScriptSerializer Json = new JavaScriptSerializer();
        private static ChatGitModalDiagnostic.Identity Identity()
        { return new ChatGitModalDiagnostic.Identity { DocumentPath = @"E:\Qualification\Disposable.docm", ProcessId = 71,
            ProcessStartedUtc = "2026-10-05T16:19:22.8081032Z", ThreadId = 73, ChatHandle = 75, RootHandle = 77,
            ProductMvid = "c4b7e1e3-dbf5-4dce-8498-1c0b7b45c973", ProductSha256 = new string('A', 64) }; }
        private static object[] Chain()
        {
            var expected = Identity(); var rows = new List<object>(); Action callback = null;
            var d = ChatGitModalDiagnostic.Begin(Json.Serialize(new { Version = 1, Nonce, Identity = expected }), Nonce, Identity,
                (phase, row) => rows.Add(Json.DeserializeObject(Json.Serialize(row))));
            d.RunModal(() => { }, () => { }); d.SchedulePostHandler(action => callback = action); callback();
            Assert.IsTrue(d.Completed); return rows.ToArray();
        }

        [TestMethod]
        public void ExactChainAuthorizesOnlyAnObservationNotAClaimOfComReadiness()
        { ChatGitDiagnosticReceipt.Validate(Chain(), Nonce, Identity()); }

        [DataTestMethod]
        [DataRow("Nonce")][DataRow("Version")][DataRow("Phase")][DataRow("Errors")][DataRow("Missing")][DataRow("Extra")]
        [DataRow("EarlySuccess")][DataRow("NoSuccess")][DataRow("Modal")][DataRow("Dispose")][DataRow("Timestamp")]
        [DataRow("Identity")][DataRow("IdentityExtra")][DataRow("IdentityNumber")][DataRow("IdentityMissing")]
        public void EveryForeignPartialOrFailedChainRefusesBeforeAnyClose(string mutation)
        {
            object[] chain = Chain(); var last = (IDictionary<string, object>)chain[4];
            if (mutation == "Nonce") last["Nonce"] = Guid.NewGuid().ToString("N");
            else if (mutation == "Version") last["Version"] = 2;
            else if (mutation == "Phase") last["Phase"] = "CallbackReturned";
            else if (mutation == "Errors") last["Errors"] = new object[] { "original error" };
            else if (mutation == "Missing") chain = chain.Take(4).ToArray();
            else if (mutation == "Extra") last["Unexpected"] = true;
            else if (mutation == "EarlySuccess") ((IDictionary<string, object>)chain[1])["Success"] = true;
            else if (mutation == "NoSuccess") last["Success"] = false;
            else if (mutation == "Modal") ((IDictionary<string, object>)chain[1])["ModalReturned"] = false;
            else if (mutation == "Dispose") ((IDictionary<string, object>)chain[2])["DisposeReturned"] = false;
            else if (mutation == "Timestamp") last["Utc"] = "2026-01-01T00:00:00.0000000Z";
            else if (mutation == "Identity") ((IDictionary<string, object>)last["Identity"])["ChatHandle"] = 999;
            else if (mutation == "IdentityExtra") ((IDictionary<string, object>)last["Identity"])["Unexpected"] = true;
            else if (mutation == "IdentityNumber") ((IDictionary<string, object>)last["Identity"])["ThreadId"] = "73";
            else ((IDictionary<string, object>)last["Identity"]).Remove("RootHandle");
            Assert.ThrowsException<InvalidOperationException>(() => ChatGitDiagnosticReceipt.Validate(chain, Nonce, Identity()));
        }

        [DataTestMethod]
        [DataRow("LateReady")][DataRow("LateRead")][DataRow("NeverReady")][DataRow("AtDeadline")]
        public void NoTerminalOrValidationAtOrAfterTheFixedDeadlineCanAuthorizeClose(string boundary)
        {
            long elapsed = boundary == "AtDeadline" ? 15000 : 0; int reads = 0, pauses = 0;
            object[] chain = Chain();
            Assert.ThrowsException<TimeoutException>(() => ChatGitDiagnosticReceipt.WaitCore(() => elapsed,
                () => { if (boundary == "LateReady") elapsed = 15000; return boundary != "NeverReady"; }, () => false,
                () => { reads++; if (boundary == "LateRead") elapsed = 15000; return chain; },
                () => { pauses++; elapsed = 15000; }, Nonce, Identity()));
            Assert.AreEqual(boundary == "LateRead" ? 1 : 0, reads); Assert.AreEqual(boundary == "NeverReady" ? 1 : 0, pauses);
        }

        [TestMethod]
        public void ReadyAndValidatedBeforeDeadlinePassWithoutAnyDelayOrReadReplay()
        {
            object[] chain = Chain(); int reads = 0;
            Assert.AreSame(chain, ChatGitDiagnosticReceipt.WaitCore(() => 14999, () => true, () => false,
                () => { reads++; return chain; }, () => Assert.Fail("No pause"), Nonce, Identity())); Assert.AreEqual(1, reads);
        }

        [DataTestMethod][DataRow(false)][DataRow(true)]
        public void FailureBeforeOrAfterReadNeverBecomesAnAuthorization(bool afterRead)
        {
            int reads = 0, checks = 0;
            Assert.ThrowsException<InvalidOperationException>(() => ChatGitDiagnosticReceipt.WaitCore(() => 0, () => true,
                () => { checks++; return !afterRead || checks > 1; }, () => { reads++; return Chain(); },
                () => Assert.Fail("No pause"), Nonce, Identity())); Assert.AreEqual(afterRead ? 1 : 0, reads);
        }

        [TestMethod]
        public void AReadErrorIsPreservedAndNeverRetried()
        {
            int reads = 0; var original = new IOException("original reader");
            var actual = Assert.ThrowsException<IOException>(() => ChatGitDiagnosticReceipt.WaitCore(() => 0, () => true, () => false,
                () => { reads++; throw original; }, () => Assert.Fail("No pause"), Nonce, Identity()));
            Assert.AreSame(original, actual); Assert.AreEqual(1, reads);
        }

        [TestMethod]
        public void ActualAtomicFilesAreReadOnceAndAnyFailedReceiptOrMissingReadyRefuses()
        {
            string root = Path.Combine(Path.GetTempPath(), "vbai-chat-receipt-unit-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(root);
            try
            {
                object[] chain = Chain(); for (int i = 0; i < chain.Length; i++) ChatGitModalDiagnostic.Publish(root, ChatGitModalDiagnostic.Phases[i], chain[i]);
                ChatGitDiagnosticReceipt.Validate(ChatGitDiagnosticReceipt.Wait(root, Nonce, Identity()), Nonce, Identity());
                string ready = Path.Combine(root, "ShowModalReturned.json.ready"); File.Delete(ready);
                Assert.ThrowsException<InvalidOperationException>(() => ChatGitDiagnosticReceipt.Wait(root, Nonce, Identity()));
                using (File.Create(ready)) { }
                ChatGitModalDiagnostic.Publish(root, "Failed", new { Error = "original" });
                Assert.ThrowsException<InvalidOperationException>(() => ChatGitDiagnosticReceipt.Wait(root, Nonce, Identity()));
            }
            finally { foreach (string path in Directory.GetFiles(root)) File.Delete(path); Directory.Delete(root); }
        }
    }
}