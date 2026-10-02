using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    /// <summary>Qualifies the actual installed Git menu/window, not an external-STA VbaGitProject.</summary>
    [TestClass, DoNotParallelize, TestCategory("NativeEmbeddedGitUi")]
    public sealed partial class EmbeddedGitWindowTests
    {
        private static readonly List<RunContext> Retained = new List<RunContext>();

        [TestMethod]
        public void InstalledOwnerGitWindowCapturesSyntheticProjectAndCreatesLocalCheckpoint()
        {
            RunInstalledOwner(null);
        }

        /// <summary>Qualifies the actual VBE owner-dispatched import for every prepared persisted native layout.</summary>
        [TestMethod]
        [DataRow("LabelButton")][DataRow("TextBox")][DataRow("ComboBox")][DataRow("ListBox")]
        [DataRow("CheckBox")][DataRow("OptionButton")][DataRow("ToggleButton")][DataRow("ScrollBar")]
        [DataRow("SpinButton")][DataRow("TabStrip")][DataRow("Image")][DataRow("FrameMultiPage")]
        public void InstalledOwnerGitWindowRestoresPersistedFormCheckpoint(string layout)
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_USERFORM_OWNER_RESTORE_TESTS") != "1")
                Assert.Inconclusive("Explicit owned UserForm checkpoint import opt-in is required.");
            RunInstalledOwner(layout);
        }

        private void RunInstalledOwner(string layout, bool persistence = false)
        {
            var fontObservation = RootFontObservationManifest.Prepare(Environment.GetEnvironmentVariable, layout, persistence);
            if (Environment.GetEnvironmentVariable("VBAi_RUN_EMBEDDED_GIT_UI_TESTS") != "1" ||
                Environment.GetEnvironmentVariable("VBAi_RUN_USERFORM_GITHUB_TESTS") != "1" ||
                Environment.GetEnvironmentVariable("VBAi_RUN_EXCEL_TESTS") != "1")
                Assert.Inconclusive("Separate owned Excel, authorized synthetic GitHub and embedded UI opt-ins are required.");
            if (!string.IsNullOrEmpty(Environment.GetEnvironmentVariable(PathVisibilityDiagnostic.EnvironmentName)))
                throw new InvalidOperationException("No unrelated path/token diagnostic manifest is allowed.");
            string root = ExcelOwnedBootstrapPlan.RequireLocalAbsolutePath(Environment.GetEnvironmentVariable("VBAi_EXCEL_RESULTS"));
            string manifestPath = ExcelOwnedBootstrapPlan.RequireLocalAbsolutePath(Environment.GetEnvironmentVariable("VBAi_TEST_GITHUB_MANIFEST"));
            var manifest = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(File.ReadAllText(manifestPath));
            var plan = ValidateManifest(manifest);
            plan.Layout = layout;
            plan.Persistence = persistence;
            Guid expected = Guid.Parse(Environment.GetEnvironmentVariable("VBAi_TEST_EMBEDDED_GIT_MVID"));
            string hash = Environment.GetEnvironmentVariable("VBAi_TEST_EMBEDDED_GIT_SHA256");
            Assert.AreEqual(typeof(VbeSession).Module.ModuleVersionId, expected, "Tests must reference the exact frozen product candidate.");
            Assert.AreEqual(hash, Sha(typeof(VbeSession).Assembly.Location), true);
            var context = new RunContext(Path.Combine(root, "embedded-git-" + Guid.NewGuid().ToString("N")), plan) { FontObservation = fontObservation };
            context.Record(new { Phase = "Preflight", ExpectedMvid = expected.ToString("D"), ExpectedSha256 = hash,
                Remote = plan.Remote, Branch = plan.Branch, BranchCommit = plan.Commit,
                Layout = layout, Scope = persistence ? "Owner-dispatched checkpoint import, one Save, normal exit, independent read-only fresh-process reopen; no remote push or recovery replay." : layout == null ? "Owner-dispatched capture/checkpoint only." : "Owner-dispatched persisted UserForm checkpoint import with exact native snapshot and font readback; no remote push, recovery replay or save/reopen acceptance." });
            var owner = new Thread(() => Owner(context, expected, hash)) { IsBackground = true };
            owner.SetApartmentState(ApartmentState.STA); owner.Start();
            try
            {
                if (!context.Ready.Wait(TimeSpan.FromSeconds(90))) throw new TimeoutException("Owned STA startup/preparation did not reach a known terminal state.");
                if (context.OwnerError != null) ExceptionDispatchInfo.Capture(context.OwnerError).Throw();
                var ui = new Thread(() => Ui(context)) { IsBackground = true };
                ui.SetApartmentState(ApartmentState.MTA); ui.Start();
                if (!context.UiReady.Wait(TimeSpan.FromSeconds(10))) throw new TimeoutException("MTA UIA discovery was not armed before modal invocation.");
                context.StartMenu.Set();
                if (!context.UiDone.Wait(TimeSpan.FromSeconds(300))) throw new TimeoutException("UIA operation delivery/observation remains uncertain.");
                if (!context.OwnerDone.Wait(TimeSpan.FromSeconds(persistence ? 180 : 30))) throw new TimeoutException("Modal Execute, persistence readback or native shutdown has no terminal evidence.");
                var failures = new[] { context.UiError, context.OwnerError }.Where(x => x != null).ToArray();
                context.Record(new { Phase = failures.Length == 0 ? "PASS" : "FAILED", NativeScope = persistence ? "PersistedFormOneSaveFreshProcessReopen" : layout == null ? "CaptureCheckpointOnly" : "PersistedFormCheckpointImport", Shutdown = context.Fixture.ShutdownDiagnostics });
                if (failures.Length > 1) throw new AggregateException("Embedded Git UI and owner STA failed; original errors preserved.", failures);
                if (failures.Length == 1) ExceptionDispatchInfo.Capture(failures[0]).Throw();
            }
            catch (Exception primary)
            {
                // No coordinator COM/UI/Close/Quit. Retain exact workers/RCWs if native delivery is not terminal.
                if (!context.MenuEmitted && !context.Retain && context.Ready.IsSet)
                {
                    context.Stop = true; context.StartMenu.Set();
                    context.OwnerDone.Wait(TimeSpan.FromSeconds(30)); // Known pre-emission failure permits the owner STA's normal cleanup.
                }
                if (!context.OwnerDone.IsSet)
                {
                    context.Stop = context.Retain = true;
                    context.Fixture?.PreserveMonacoNativeOutcome();
                    lock (Retained) Retained.Add(context);
                }
                Exception preserved = EmbeddedGitUiProtocol.PreserveFailures(primary, context.UiError, context.OwnerError);
                try
                {
                    context.Record(new { Phase = "FailedOrRetained", Retained = context.Retain, Error = preserved.ToString(),
                        ProcessId = context.Fixture?.ProcessId, MenuEmitted = context.MenuEmitted, ModalClosed = context.ModalClosed,
                        OwnerTerminal = context.OwnerDone.IsSet, UiTerminal = context.UiDone.IsSet, ReplayAttempts = 0 });
                }
                catch (Exception evidence) { throw new AggregateException("Original scenario and durable failure recording both failed.", preserved, evidence); }
                if (!ReferenceEquals(primary, preserved)) ExceptionDispatchInfo.Capture(preserved).Throw();
                throw;
            }
        }

        private static void Owner(RunContext context, Guid expected, string hash)
        {
            bool nativePending = false;
            try
            {
                context.Fixture = ExcelVbeFixture.StartOwnedWithTrace(Path.Combine(context.Output, "unused-scalar-trace.jsonl"));
                context.Record(new { Phase = "OwnedStaReady", context.Fixture.ProcessId, OwnerSta = Thread.CurrentThread.ManagedThreadId });
                nativePending = true;
                var status = context.Fixture.Command("status"); Assert.IsNotNull(status);
                nativePending = false;
                Assert.AreEqual(true, status["Ok"]);
                var data = VbeBridgeClient.Object(status["Data"]);
                ExcelVbeFixture.RequireMonacoCandidate(expected, typeof(VbeSession).Module.ModuleVersionId, context.Fixture.ProcessId, data);
                Assert.AreEqual(hash, Sha(Convert.ToString(data["AssemblyPath"])), true, "Loaded installed bytes differ from the frozen candidate.");
                context.Scope = context.Fixture.PrepareEmbeddedGitScope(context.Nonce, value => nativePending = value,
                    context.Record, context.Plan.Layout, context.FontObservation?.SeedProfile);
                nativePending = false;
                if (context.FontObservation != null)
                {
                    var claim = RootFontObservationManifest.Build(context.FontObservation, context.Output, context.Scope.Path,
                        context.Scope.Baseline, expected, Guid.Parse(Convert.ToString(data["AssemblyModuleVersionId"])), Guid.NewGuid());
                    RootFontObservationManifest.Publish(context.FontObservation, claim, Environment.GetEnvironmentVariable);
                    context.Record(new { Phase = "RootFontDiagnosticManifestPublished", ManifestPath = context.FontObservation.Path,
                        BaselineFontSeedProfile = context.FontObservation.SeedProfile,
                        Manifest = claim, NativeFontDelivery = "NOT_RUN", GetterOrExportRequestsAdded = 0 });
                }
                context.WorkbookSha256 = Sha(context.Scope.Path);
                context.Record(new { Phase = "Prepared", ProcessId = context.Fixture.ProcessId, context.Scope.ThreadId,
                    VbeHandle = context.Scope.VbeHandle.ToInt64(), context.Scope.Path, context.Scope.Cache,
                    WorkbookSha256 = context.WorkbookSha256, StateSha256 = ShaText(context.Scope.State), Mvid = expected.ToString("D") });
                context.Ready.Set();
                if (!context.StartMenu.Wait(TimeSpan.FromSeconds(30))) throw new TimeoutException("UIA worker did not prepare the modal invocation.");
                if (context.Stop) throw new InvalidOperationException("Coordinator stopped before modal emission.");
                context.Fixture.ExecuteEmbeddedGitMenu(value => {
                    context.Record(value);
                    if (new JavaScriptSerializer().Serialize(value).Contains("MenuExecuteIntent"))
                    { context.MenuEmitted = true; context.MenuIntent.Set(); }
                });
                if (!context.UiDone.Wait(TimeSpan.FromSeconds(15))) throw new TimeoutException("Modal returned without terminal UIA closure observation.");
                if (!context.ModalClosed) throw new InvalidOperationException("Modal returned without the exact known UIA close proof.");
                if (context.Plan.Layout == null) context.Fixture.VerifyEmbeddedGitState(context.Scope);
                else context.Fixture.VerifyEmbeddedImportedForm(context.Scope, context.Record);
                Assert.AreEqual(context.WorkbookSha256, Sha(context.Scope.Path), "Owner-dispatched capture/checkpoint must not change the owned disk workbook.");
                nativePending = true;
                status = context.Fixture.Command("status"); Assert.IsNotNull(status);
                nativePending = false;
                Assert.AreEqual(true, status["Ok"]);
                ExcelVbeFixture.RequireMonacoCandidate(expected, typeof(VbeSession).Module.ModuleVersionId, context.Fixture.ProcessId, VbeBridgeClient.Object(status["Data"]));
                nativePending = false;
                context.Record(new { Phase = "NativeStatePreserved", StateSha256 = ShaText(context.Scope.State) });
                if (context.Plan.Persistence)
                {
                    context.Fixture.SaveEmbeddedImportedForm(context.Scope, context.Plan.Remote, context.Plan.Branch, context.Plan.Commit, value => {
                        if (value && context.Stop) { context.Fixture.PreserveMonacoNativeOutcome(); throw new InvalidOperationException("Coordinator stopped before the next persistence dispatch."); }
                        nativePending = value;
                    }, context.Record);
                    context.WorkbookSha256 = Sha(context.Scope.Path);
                    context.Record(new { Phase = "PostImportSaveVerified", Path = context.Scope.Path,
                        WorkbookSha256 = context.WorkbookSha256, SaveAttempts = 1 });
                }
            }
            catch (Exception error) { context.OwnerError = error; if (context.Fixture?.PreserveForDiagnosticRecovery == true || EmbeddedGitUiProtocol.MustRetainOwner(nativePending, context.MenuEmitted, context.ModalClosed)) context.Retain = true; }
            finally
            {
                context.Ready.Set();
                if (context.Fixture != null)
                {
                    if (context.Retain)
                    {
                        context.Fixture.PreserveMonacoNativeOutcome(); lock (Retained) if (!Retained.Contains(context)) Retained.Add(context);
                    }
                    else
                    {
                        try
                        {
                            context.Record(new { Phase = "NormalCleanupIntent", CloseQuitReplay = false });
                            context.Fixture.Dispose();
                            Assert.IsNotNull(context.Fixture.ShutdownDiagnostics);
                            Assert.AreEqual(true, context.Fixture.ShutdownDiagnostics["Exited"]);
                            Assert.AreEqual(0, Convert.ToInt32(context.Fixture.ShutdownDiagnostics["ExitCode"]));
                            Assert.AreEqual(false, context.Fixture.ShutdownDiagnostics["ForcedTermination"]);
                            if (context.WorkbookSha256 != null)
                                Assert.AreEqual(context.WorkbookSha256, Sha(context.Scope.Path), "Normal shutdown must preserve the saved owned workbook bytes.");
                            if (context.Plan.Persistence && context.OwnerError == null && context.UiError == null)
                                VerifyFreshImportedWorkbook(context, expected, hash);
                        }
                        catch (Exception cleanup)
                        {
                            context.Retain = true;
                            context.OwnerError = context.OwnerError == null ? cleanup : new AggregateException("Owner scenario and normal cleanup both failed.", context.OwnerError, cleanup);
                            lock (Retained) if (!Retained.Contains(context)) Retained.Add(context);
                        }
                    }
                }
                context.OwnerDone.Set();
            }
        }

        private static void Ui(RunContext context)
        {
            var automation = new EmbeddedGitAutomation(context.Fixture, context.Scope, context.Record, () => context.Stop);
            try
            {
                context.Record(new { Phase = "MtaUiaDiscoveryArmed", ThreadId = Thread.CurrentThread.ManagedThreadId,
                    Apartment = Thread.CurrentThread.GetApartmentState().ToString() });
                context.UiReady.Set();
                if (!context.MenuIntent.Wait(TimeSpan.FromSeconds(40)))
                    throw new TimeoutException("No modal intent was emitted; UIA has no authority to invoke or close a window.");
                automation.OpenedWindow(); automation.Link(context.Plan.Remote, context.Plan.Branch, context.Plan.Commit);
                automation.Checkpoint(context.Plan.Remote, context.Plan.Branch, context.Nonce, context.Plan.TabName);
                if (context.Plan.Layout != null) automation.MutateAndRestoreCheckpoint(context.Nonce, context.Plan.TabName);
                automation.CompareOnce();
                if (context.Plan.Persistence)
                    context.Fixture.AttestEmbeddedSelectedRepository(context.Scope, context.Plan.Remote, context.Plan.Branch, context.Plan.Commit, context.Record);
            }
            catch (Exception error) { context.UiError = error; }
            finally
            {
                if (!context.MenuEmitted) { /* Known pre-emission failure: only the owner STA may close its owned workbook normally. */ }
                else if (automation.HasObservedWindow && automation.Protocol.CanClose && !context.Stop)
                {
                    try { automation.CloseKnownTerminal(); context.ModalClosed = true; }
                    catch (Exception cleanup) { context.Retain = true; context.UiError = context.UiError == null ? cleanup : new AggregateException("UI scenario and exact modal close both failed.", context.UiError, cleanup); }
                }
                else context.Retain = true;
                context.UiDone.Set();
            }
        }

        internal sealed class TestPlan { internal string Remote, Branch, Commit, TabName, Layout; internal bool Persistence; }
        internal static TestPlan ValidateManifest(IDictionary<string, object> manifest)
        {
            string remote = Convert.ToString(manifest["repositoryUrl"]) + ".git";
            if (remote != "https://github.com/blackcancer/vbai-qualification-20260929203712-7267b1e6.git" || Convert.ToString(manifest["repositoryId"]) != "1396566119")
                throw new InvalidOperationException("Only the maintainer's exact retained synthetic repository is allowed.");
            var result = new TestPlan { Remote = remote, Branch = Convert.ToString(manifest["embeddedBranch"]),
                Commit = Convert.ToString(manifest["embeddedBranchCommit"]), TabName = Convert.ToString(manifest["checkpointTabName"]) };
            const string retainedBranch = "qualification-userform-20260929225354-93ed53dc";
            const string retainedCommit = "f5fb1a004dcb287c4c820b4bf308c673ce3b64e6";
            bool authorizedBranch = result.Branch == retainedBranch
                ? result.Commit == retainedCommit
                : Regex.IsMatch(result.Branch ?? "", "^qualification-embedded-ui-[a-z0-9-]{1,64}$");
            if (!authorizedBranch ||
                !Regex.IsMatch(result.Commit ?? "", "^[0-9a-f]{40}$") || string.IsNullOrWhiteSpace(result.TabName) || result.TabName.Length > 80 || result.TabName.Any(char.IsControl))
                throw new InvalidOperationException("Explicit authorized synthetic branch/revision and observed Checkpoints tab caption are required; no fallback.");
            return result;
        }
        private static string Sha(string path) { using (var bytes = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) using (var hash = SHA256.Create()) return BitConverter.ToString(hash.ComputeHash(bytes)).Replace("-", ""); }
        [TestMethod, TestCategory("Unit")]
        public void OwnedSavedWorkbookHashSupportsAnExistingWriterWithoutChangingItsBytes()
        {
            string directory = Path.Combine(Path.GetTempPath(), "VBAi-WorkbookHash-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            string path = Path.Combine(directory, "synthetic.bin");
            byte[] expected = Encoding.UTF8.GetBytes("Synthetic saved workbook bytes");
            try
            {
                File.WriteAllBytes(path, expected);
                using (var owner = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.Read))
                {
                    Assert.AreEqual(ShaText("Synthetic saved workbook bytes"), Sha(path));
                    Assert.AreEqual(expected.Length, owner.Length);
                }
                CollectionAssert.AreEqual(expected, File.ReadAllBytes(path));
            }
            finally { File.Delete(path); Directory.Delete(directory); }
        }

        private static string ShaText(string text) { using (var hash = SHA256.Create()) return BitConverter.ToString(hash.ComputeHash(Encoding.UTF8.GetBytes(text))).Replace("-", ""); }

        private sealed class RunContext
        {
            internal readonly string Output, Nonce = "EMBEDDED_" + Guid.NewGuid().ToString("N");
            internal readonly TestPlan Plan;
            internal readonly ManualResetEventSlim Ready = new ManualResetEventSlim(), StartMenu = new ManualResetEventSlim(), UiReady = new ManualResetEventSlim(), MenuIntent = new ManualResetEventSlim(), UiDone = new ManualResetEventSlim(), OwnerDone = new ManualResetEventSlim();
            internal ExcelVbeFixture Fixture;
            internal ExcelVbeFixture.EmbeddedGitScope Scope;
            internal RootFontObservationManifest.Configuration FontObservation;
            internal string WorkbookSha256;
            internal Exception OwnerError, UiError;
            internal volatile bool Stop, Retain, MenuEmitted, ModalClosed;
            private int sequence;
            private readonly object sync = new object();
            internal RunContext(string output, TestPlan plan) { Output = output; Plan = plan; Directory.CreateDirectory(output); }
            internal void Record(object value)
            {
                lock (sync)
                {
                    if (++sequence > (Plan.Persistence ? 240 : 160)) throw new InvalidOperationException("Bounded synthetic action evidence capacity exhausted; no further action.");
                    using (var file = new FileStream(Path.Combine(Output, sequence.ToString("D3") + ".json"), FileMode.CreateNew, FileAccess.Write, FileShare.Read))
                    using (var writer = new StreamWriter(file, new UTF8Encoding(false))) writer.Write(new JavaScriptSerializer().Serialize(value));
                }
            }
        }
    }
}
