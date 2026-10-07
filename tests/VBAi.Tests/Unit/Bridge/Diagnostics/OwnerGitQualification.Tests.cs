using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using System.Web.Script.Serialization;
using System.IO;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class OwnerGitQualificationTests
    {
        [STATestMethod]
        public void DistinctCommandScopesAreCapturedBeforeTheirFirstOwnerContinuation()
        {
            var original = VbeInspectionTrace.Current;
            var rows = new List<string>();
            int calls = 0, owner = Thread.CurrentThread.ManagedThreadId;
            for (int command = 0; command < 2; command++)
            {
                var trace = new VbeInspectionTrace(rows.Add);
                var task = OwnerGitQualification.RunOnOwnerAsync(async () =>
                {
                    calls++;
                    Assert.AreSame(trace, VbeInspectionTrace.Current);
                    await Task.Yield();
                    Assert.AreSame(trace, VbeInspectionTrace.Current);
                    Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                    return (object)42;
                }, trace);
                Pump(task); Assert.AreEqual(42, task.GetAwaiter().GetResult());
                Assert.AreSame(original, VbeInspectionTrace.Current);
            }
            Assert.AreEqual(2, calls);
            var parsed = rows.Select(Parse).ToArray();
            var commands = parsed.GroupBy(row => (string)row["Correlation"]).ToArray();
            Assert.AreEqual(2, commands.Length);
            foreach (var command in commands)
            {
                var phases = command.Select(row => (string)row["Phase"]).ToArray();
                Assert.AreEqual("CallbackEntered", phases[0]);
                Assert.AreEqual("OwnerSta", phases[1]);
                Assert.AreEqual(1, phases.Count(phase => phase == "ContinuationEnqueued"));
                Assert.AreEqual(1, phases.Count(phase => phase == "ContinuationPostReturned"));
                Assert.AreEqual(1, phases.Count(phase => phase == "ContinuationEntered"));
                Assert.AreEqual(1, phases.Count(phase => phase == "Terminal"));
                Assert.IsTrue(command.Count() <= VbeInspectionTrace.MaximumEvents);
            }
        }

        [STATestMethod]
        public void DisabledTracePreservesOneSuccessfulCommand() { VerifyOptionalTrace(false, false); }

        [STATestMethod]
        public void DisabledTracePreservesTheOriginalCommandFailure() { VerifyOptionalTrace(false, true); }

        [STATestMethod]
        public void BrokenTracePreservesOneSuccessfulCommand() { VerifyOptionalTrace(true, false); }

        [STATestMethod]
        public void BrokenTracePreservesTheOriginalCommandFailure() { VerifyOptionalTrace(true, true); }

        private static void VerifyOptionalTrace(bool brokenWriter, bool fail)
        {
            var original = VbeInspectionTrace.Current;
            var trace = brokenWriter ? new VbeInspectionTrace(_ => { throw new IOException("SECRET_PATH"); }) : VbeInspectionTrace.ForPath(null);
            var failure = new InvalidOperationException("original operation failure");
            int calls = 0;
            var task = OwnerGitQualification.RunOnOwnerAsync(async () =>
            {
                calls++;
                await Task.Yield();
                if (fail) throw failure;
                return (object)73;
            }, trace);
            Pump(task);
            if (fail) Assert.AreSame(failure, Assert.ThrowsException<InvalidOperationException>(() => task.GetAwaiter().GetResult()));
            else Assert.AreEqual(73, task.GetAwaiter().GetResult());
            Assert.AreEqual(1, calls); Assert.AreSame(original, VbeInspectionTrace.Current);
        }

        private static Dictionary<string, object> Parse(string row)
        {
            return new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(row);
        }

        private static void Pump(Task task)
        {
            var elapsed = Stopwatch.StartNew();
            while (!task.IsCompleted && elapsed.ElapsedMilliseconds < 3000)
            {
                Application.DoEvents(); Thread.Sleep(1);
            }
            Assert.IsTrue(task.IsCompleted, "The owned test dispatcher did not finish within its bounded message pump.");
        }

        [TestMethod]
        public void ExpectedCorruptionRefusalNeedsMeasuredUnchangedProjectAndRecoveryRefs()
        {
            Func<bool, bool, bool, string, string, bool, bool> prove = (started, pending, sameProject, before, after, sameAfter) =>
                OwnerGitQualification.IsProvenPrewriteRefusal("Form resources are missing", "Form resources are missing: Form1",
                    started, false, pending, before, after, "after-0", sameAfter ? "after-0" : "after-1", sameProject);
            Assert.IsTrue(prove(false, false, true, "backup-0", "backup-0", true));
            Assert.IsFalse(prove(true, false, true, "backup-0", "backup-0", true));
            Assert.IsFalse(prove(false, true, true, "backup-0", "backup-0", true));
            Assert.IsFalse(prove(false, false, false, "backup-0", "backup-0", true));
            Assert.IsFalse(prove(false, false, true, "backup-0", "backup-1", true));
            Assert.IsFalse(prove(false, false, true, "backup-0", "backup-0", false));
            Assert.IsFalse(OwnerGitQualification.IsProvenPrewriteRefusal("missing", "other error", false,
                false, false, null, null, null, null, true));
        }

        [TestMethod]
        public async Task DisabledAtConnectionCannotBeEnabledByChangingEnvironmentLater()
        {
            string previous = Environment.GetEnvironmentVariable(OwnerGitQualificationManifest.EnvironmentName);
            try
            {
                Environment.SetEnvironmentVariable(OwnerGitQualificationManifest.EnvironmentName, null);
                var disabled = new OwnerGitQualification(null, 42);
                Environment.SetEnvironmentVariable(OwnerGitQualificationManifest.EnvironmentName,
                    @"C:\Evidence\0123456789abcdef0123456789abcdef.owner-git.json");
                var request = new Request
                {
                    Command = OwnerGitQualificationManifest.CommandName,
                    Action = "0123456789abcdef0123456789abcdef",
                    ExpectedSha256 = new string('a', 64)
                };
                string json = "{\"Command\":\"diagnostic_userform_git\",\"Action\":\"" + request.Action +
                    "\",\"ExpectedSha256\":\"" + request.ExpectedSha256 + "\"}";
                var error = await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => disabled.ExecuteAsync(request, json));
                StringAssert.Contains(error.Message, "disabled");
            }
            finally { Environment.SetEnvironmentVariable(OwnerGitQualificationManifest.EnvironmentName, previous); }
        }
    }
}
