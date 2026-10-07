using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Web.Script.Serialization;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Qualifies native-scalar evidence and error containment with synthetic delegates, without Excel or UI.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class ExcelScalarQualificationEvidenceTests
    {
        /// <summary>Intent survives a response timeout, including elapsed time and the original exception, without replay.</summary>
        [TestMethod]
        public void PendingRequestIsPersistedBeforeTheOnlySendAndItsTimeoutIsRetained()
        {
            WithEvidence((evidence, path) =>
            {
                int sends = 0;
                var timeout = new TimeoutException("Synthetic uncertain bridge delivery.");
                var observed = Assert.ThrowsException<TimeoutException>(() => evidence.Send(new { Command = "inspect_local_scalars" }, () =>
                {
                    sends++;
                    var prepared = Entry(path, 0);
                    Assert.AreEqual("inspect_local_scalars", ((IDictionary<string, object>)prepared["Request"])["Command"]);
                    StringAssert.StartsWith((string)prepared["State"], "Prepared;");
                    throw timeout;
                }, response => Assert.Fail("Timeout cannot return a response.")));
                Assert.AreSame(timeout, observed);
                Assert.AreEqual(1, sends);
                var failed = Entry(path, 0);
                StringAssert.Contains((string)failed["Error"], timeout.Message);
                StringAssert.StartsWith((string)failed["State"], "Failed;");
                Assert.IsNull(failed["Response"]);
                Assert.IsTrue(Convert.ToInt64(failed["ElapsedMilliseconds"]) >= 0);
                Assert.IsNotNull(failed["FinishedUtc"]);
            });
        }

        /// <summary>A disconnected cleanup read preserves the scenario timeout and leaves shutdown unverified.</summary>
        [TestMethod]
        public void CleanupNullResponseDoesNotMaskInspectionTimeoutOrPermitClosure()
        {
            WithEvidence((evidence, path) =>
            {
                var timeout = new TimeoutException("Inspection timed out after emission.");
                int sends = 0, closes = 0, attachments = 0;
                var failure = Assert.ThrowsException<AggregateException>(() => evidence.Run(() =>
                {
                    evidence.Send(new { Command = "inspect_local_scalars" }, () => { sends++; throw timeout; }, _ => { });
                }, () =>
                {
                    evidence.Send(new { Command = "debug_state" }, () => { sends++; return null; }, response => Assert.IsNotNull(response, "Owned bridge unavailable; retain host."));
                    closes++;
                }, () => attachments++));
                Assert.AreEqual(2, failure.InnerExceptions.Count);
                Assert.AreSame(timeout, failure.InnerExceptions[0]);
                Assert.IsInstanceOfType<AssertFailedException>(failure.InnerExceptions[1]);
                Assert.AreEqual(2, sends);
                Assert.AreEqual(0, closes);
                Assert.AreEqual(1, attachments);
                var final = Read(path);
                StringAssert.Contains((string)final["PrimaryError"], timeout.Message);
                StringAssert.Contains((string)final["CleanupError"], "Owned bridge unavailable");
                StringAssert.Contains((string)final["Shutdown"], "NotVerified");
                StringAssert.Contains((string)Entry(path, 1)["Error"], "Owned bridge unavailable");
            });
        }

        /// <summary>A real locked evidence file preserves command and persistence errors without a second send.</summary>
        [TestMethod]
        public void LockedResultWritePreservesDeliveryErrorAndPreparedEvidenceWithoutRetry()
        {
            WithEvidence((evidence, path) =>
            {
                var timeout = new TimeoutException("Synthetic mutation response deadline.");
                FileStream locked = null;
                int sends = 0;
                try
                {
                    var failure = Assert.ThrowsException<AggregateException>(() => evidence.Send(new { Command = "run_sub" }, () =>
                    {
                        sends++;
                        locked = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                        throw timeout;
                    }, _ => { }));
                    Assert.AreSame(timeout, failure.InnerExceptions[0]);
                    Assert.IsInstanceOfType<IOException>(failure.InnerExceptions[1]);
                    Assert.AreEqual(1, sends);
                }
                finally { locked?.Dispose(); }
                StringAssert.StartsWith((string)Entry(path, 0)["State"], "Prepared;");
            });
        }

        /// <summary>Cleanup and attachment run independently, and their distinct errors remain visible.</summary>
        [TestMethod]
        public void CleanupAndAttachmentErrorsPreservePrimaryFailureAndFinalReport()
        {
            WithEvidence((evidence, path) =>
            {
                var primary = new TimeoutException("Original inspection failure.");
                var cleanup = new IOException("Synthetic cleanup failure.");
                var attachment = new InvalidOperationException("Synthetic attachment failure.");
                var failure = Assert.ThrowsException<AggregateException>(() => evidence.Run(() => { throw primary; },
                    () => { throw cleanup; }, () => { throw attachment; }));
                CollectionAssert.AreEqual(new Exception[] { primary, cleanup, attachment }, failure.InnerExceptions);
                StringAssert.Contains((string)Read(path)["PrimaryError"], primary.Message);
                StringAssert.Contains((string)Read(path)["CleanupError"], cleanup.Message);
            });
        }

        /// <summary>Successful response validation and normal shutdown are persisted as separate evidence.</summary>
        [TestMethod]
        public void SuccessfulResponseAndShutdownAreDurableWithoutInventingHostAcceptance()
        {
            WithEvidence((evidence, path) =>
            {
                int sends = 0, validations = 0, attachments = 0;
                var reply = new Dictionary<string, object> { ["Ok"] = true };
                evidence.Run(() => Assert.AreSame(reply, evidence.Send(new { Command = "debug_state" },
                    () => { sends++; return reply; }, response => { validations++; Assert.AreEqual(true, response["Ok"]); })),
                    () => evidence.Shutdown = "Synthetic cleanup completed", () => attachments++);
                Assert.AreEqual(1, sends);
                Assert.AreEqual(1, validations);
                Assert.AreEqual(1, attachments);
                var final = Read(path);
                Assert.IsNull(final["PrimaryError"]);
                Assert.IsNull(final["CleanupError"]);
                Assert.AreEqual("Synthetic cleanup completed", final["Shutdown"]);
                Assert.AreEqual("ResponseVerified", Entry(path, 0)["State"]);
                Assert.AreEqual(true, ((IDictionary<string, object>)Entry(path, 0)["Response"])["Ok"]);
            });
        }

        private static IDictionary<string, object> Read(string path) =>
            (IDictionary<string, object>)new JavaScriptSerializer().DeserializeObject(File.ReadAllText(path));
        private static IDictionary<string, object> Entry(string path, int index) =>
            (IDictionary<string, object>)((object[])Read(path)["Evidence"])[index];

        private static void WithEvidence(Action<ExcelScalarQualificationEvidence, string> action)
        {
            string root = Path.Combine(Path.GetTempPath(), "VBAi-scalar-evidence-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            string path = Path.Combine(root, "synthetic-evidence.json");
            try { action(new ExcelScalarQualificationEvidence(path, 4242, root), path); }
            finally { if (File.Exists(path)) File.Delete(path); Directory.Delete(root, false); }
        }
    }
}
