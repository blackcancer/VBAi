using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.ExceptionServices;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    public sealed partial class EmbeddedGitWindowTests
    {
        private static readonly List<ExcelVbeFixture> RetainedPersistenceHosts = new List<ExcelVbeFixture>();

        /// <summary>Qualifies disk persistence only after the separate exact owner-dispatched import contract succeeds.</summary>
        [TestMethod, TestCategory("NativeEmbeddedGitPersistence")]
        [DataRow("LabelButton")][DataRow("TextBox")][DataRow("ComboBox")][DataRow("ListBox")]
        [DataRow("CheckBox")][DataRow("OptionButton")][DataRow("ToggleButton")][DataRow("ScrollBar")]
        [DataRow("SpinButton")][DataRow("TabStrip")][DataRow("Image")][DataRow("FrameMultiPage")]
        public void InstalledOwnerImportedFormSurvivesOneSaveAndFreshReadOnlyProcess(string layout)
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_USERFORM_OWNER_PERSISTENCE_TESTS") != "1" ||
                Environment.GetEnvironmentVariable("VBAi_RUN_USERFORM_OWNER_RESTORE_TESTS") != "1")
                Assert.Inconclusive("Owner import and its separate one-save/fresh-process persistence opt-ins are required.");
            RunInstalledOwner(layout, persistence: true);
        }

        /// <summary>Runs on the same test STA, after the original retained process proved an unforced zero exit.</summary>
        private static void VerifyFreshImportedWorkbook(RunContext context, Guid expected, string hash)
        {
            EmbeddedGitPersistenceContract.RequireNormalExit(context.Fixture.ShutdownDiagnostics);
            string originalRoot = context.Fixture.Root;
            string originalStart = context.Fixture.EmbeddedProcessStartedUtc;
            string savedHash = context.WorkbookSha256;
            ExcelVbeFixture fresh = null;
            Exception failure = null;
            bool pending = false;
            Action<bool> observePending = value => {
                if (value && context.Stop)
                {
                    fresh?.PreserveMonacoNativeOutcome();
                    throw new InvalidOperationException("Coordinator stopped before the next fresh-process dispatch.");
                }
                pending = value;
            };
            try
            {
                if (context.Stop) throw new InvalidOperationException("Coordinator stopped before fresh-process launch; no second host was started.");
                context.Record(new { Phase = "FreshProcessLaunchIntent", Path = context.Scope.Path,
                    OriginalRoot = originalRoot, OriginalProcessId = context.Fixture.ProcessId, OriginalStartedUtc = originalStart,
                    SavedWorkbookSha256 = savedHash, ImportReplayAttempts = 0 });
                fresh = ExcelVbeFixture.StartOwnedWithTrace(Path.Combine(context.Output, "fresh-process-unused-trace.jsonl"));
                EmbeddedGitPersistenceContract.RequireFreshIdentity(originalRoot, originalStart, context.Fixture.ProcessId,
                    fresh.Root, fresh.EmbeddedProcessStartedUtc, fresh.ProcessId);
                observePending(true);
                fresh.OpenOwnedReadOnlyWorkbook(context.Scope.Path);
                pending = false;
                fresh.VerifyEmbeddedPersistenceCandidate(expected, hash, observePending, context.Record);
                EmbeddedGitPersistenceContract.RequireDiskHash(savedHash, Sha(context.Scope.Path));
                fresh.VerifyEmbeddedReopenedForm(context.Scope, observePending, context.Record);
                EmbeddedGitPersistenceContract.RequireDiskHash(savedHash, Sha(context.Scope.Path));
                context.Record(new { Phase = "FreshProcessReadbackVerified", fresh.ProcessId, fresh.Root,
                    StartedUtc = fresh.EmbeddedProcessStartedUtc, SavedWorkbookSha256 = savedHash,
                    MacroExecutions = 0, SavesInFreshProcess = 0, ImportReplayAttempts = 0 });
            }
            catch (Exception error)
            {
                failure = error;
                if (fresh != null && (pending || context.Stop || fresh.PreserveForDiagnosticRecovery))
                {
                    fresh.PreserveMonacoNativeOutcome();
                    lock (RetainedPersistenceHosts) RetainedPersistenceHosts.Add(fresh);
                    try
                    {
                        context.Record(new { Phase = "FreshProcessRetained", fresh.ProcessId, fresh.Root,
                            NativePending = pending, Error = error.ToString(), CleanupReplayAttempts = 0 });
                    }
                    catch (Exception evidence) { failure = new AggregateException("Fresh-process readback and retention evidence both failed.", error, evidence); }
                }
            }
            if (fresh != null && context.Stop) fresh.PreserveMonacoNativeOutcome();
            if (fresh != null && !fresh.PreserveForDiagnosticRecovery)
            {
                try
                {
                    fresh.Dispose();
                    EmbeddedGitPersistenceContract.RequireNormalExit(fresh.ShutdownDiagnostics);
                    EmbeddedGitPersistenceContract.RequireDiskHash(savedHash, Sha(context.Scope.Path));
                    context.Record(new { Phase = "FreshProcessNormalExitVerified", Shutdown = fresh.ShutdownDiagnostics,
                        SavedWorkbookSha256 = savedHash });
                }
                catch (Exception cleanup)
                {
                    lock (RetainedPersistenceHosts) RetainedPersistenceHosts.Add(fresh);
                    failure = failure == null ? cleanup : new AggregateException("Fresh-process readback and cleanup both failed; no action is replayed.", failure, cleanup);
                }
            }
            if (failure != null) ExceptionDispatchInfo.Capture(failure).Throw();
        }
    }
}
