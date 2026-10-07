using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace VBAi.Tests.Integration
{
    public sealed partial class NativeUserFormLocalGitTests
    {
        [STATestMethod]
        public void AttachedFrameFontBaselineObservedBeforeIntentionalPrewriteStop()
        { RunAttachedFonts("baseline-only"); }

        [STATestMethod]
        public void AttachedFrameFontObservedAfterOneOriginalImportDelivery()
        { RunAttachedFonts("font-delivery-returned"); }

        [STATestMethod]
        public void AttachedFrameFontBaselineObservedWithSetupPhaseReceipts()
        { RunAttachedFonts("baseline-only", setupMarkers: true); }

        private void RunAttachedFonts(string stage, bool setupMarkers = false)
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_EXCEL_TESTS") != "1" ||
                Environment.GetEnvironmentVariable("VBAi_RUN_USERFORM_LOCAL_GIT_TESTS") != "1" ||
                Environment.GetEnvironmentVariable("VBAi_RUN_ATTACHED_FONT_DIAGNOSTIC") != "1")
                Assert.Inconclusive("Explicit Excel/local Git/attached-font diagnostic opt-ins are required.");
            if (setupMarkers && Environment.GetEnvironmentVariable("VBAi_RUN_ATTACHED_SETUP_DIAGNOSTIC") != "1")
                Assert.Inconclusive("Explicit attached setup diagnostic opt-in is required.");
            string setupTrace = setupMarkers ? Environment.GetEnvironmentVariable("VBAi_TEST_ATTACHED_SETUP_TRACE") : null;
            if (setupMarkers) Assert.IsFalse(string.IsNullOrWhiteSpace(setupTrace), "A fresh setup trace path is required.");
            Assert.IsTrue(UserFormQualificationFonts.Enabled, "The actual saved/reopened qualified font fixture is required.");
            string root = Environment.GetEnvironmentVariable("VBAi_TEST_USERFORM_LOCAL_GIT_OUTPUT");
            Assert.IsFalse(string.IsNullOrWhiteSpace(root)); Assert.IsTrue(Path.IsPathRooted(root));
            string output = Path.Combine(root, "Attached-" + Guid.NewGuid().ToString("N"));
            Assert.IsFalse(Directory.Exists(output)); Directory.CreateDirectory(output);
            string nonce = Guid.NewGuid().ToString("N");
            string observationRoot = Path.Combine(output, nonce);
            string attachedPath = Path.Combine(output, nonce + ".attached-font.json");
            string previous = Environment.GetEnvironmentVariable(FormFontObservation.AttachedManifestVariable);
            Assert.IsTrue(string.IsNullOrEmpty(previous), "No overlapping attached diagnostic.");
            Assert.IsTrue(string.IsNullOrEmpty(Environment.GetEnvironmentVariable(FormFontObservation.ManifestVariable)),
                "The experimental restoration diagnostic must remain disabled.");
            var report = new Dictionary<string, object>
            {
                ["Stage"] = "STARTED", ["DiagnosticStage"] = stage, ["EvidenceDirectory"] = output,
                ["PotentialFontCacheMutation"] = true, ["GetterPerOwner"] = "Size", ["MacroExecutions"] = 0,
                ["RemoteOperations"] = 0, ["NativeStepsEmitted"] = 0, ["SourceRevisionResignedAfterObservation"] = false,
                ["AttachedManifest"] = attachedPath, ["AttachedReceipts"] = observationRoot,
                ["AssemblyMvid"] = typeof(FormFontObservation).Module.ModuleVersionId.ToString("D"),
                ["SetupTrace"] = setupTrace, ["SetupMarkers"] = setupMarkers
            };
            var previousContext = SynchronizationContext.Current;
            Environment.SetEnvironmentVariable(FormFontObservation.AttachedManifestVariable, attachedPath);
            using (var progress = NativeFixtureProgressTrace.BeginAtPath(setupTrace))
            using (var dispatcher = new Control())
            using (var owner = new OwnerGitQualificationScope(output))
            {
                dispatcher.CreateControl(); SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
                try
                {
                    if (setupMarkers) Assert.IsTrue(progress.WriterAvailable, "The fresh setup trace must be armed before native host launch.");
                    ExcelVbeFixture.Run(host =>
                    {
                        report["HostProcessId"] = host.ProcessId; report["FixtureRoot"] = host.Root;
                        string path = host.File("attached-font.xlsm"); const string form = "QualificationForm";
                        NativeFixtureProgressTrace.Run(NativeFixtureProgressTrace.Phase.PrepareLayout,
                            () => host.PrepareGitLayout(form, "FrameMultiPage", path, persistedBaseline: true));
                        host.WithGitProject(path, project =>
                        {
                            // Capture only exported files. The generic Capture helper reads font metrics and is intentionally not used.
                            var before = project.Capture(); SaveSnapshot(output, "before", before);
                            var selected = before.Manifest.Components.Single(item => item.Name == form);
                            var fonts = before.FormFonts(selected);
                            Assert.AreEqual(82700u, BitConverter.ToUInt32(fonts.Single(item => item.Type == 14).Descriptor, 6),
                                "The saved/reopened Frame descriptor must actually contain 8.27 before this diagnostic.");
                            var expected = before;
                            if (stage == "font-delivery-returned")
                            {
                                host.MutateGitLayout(form, "FrameMultiPage", persistedBaseline: true);
                                expected = project.Capture(); Assert.IsFalse(expected.SameAs(before));
                            }
                            SaveSnapshot(output, "expected", expected);
                            var selectedSnapshot = stage == "baseline-only" ? expected : before;
                            var attached = new FormFontObservation.AttachedManifest
                            {
                                ProjectPath = path, FormName = form, CandidateMvid = typeof(FormFontObservation).Module.ModuleVersionId.ToString("D"),
                                FormSha256 = OwnerGitQualificationManifest.Hash(selectedSnapshot.Files[selected.FileName]),
                                ResourceSha256 = OwnerGitQualificationManifest.Hash(selectedSnapshot.Files[form + ".frx"]),
                                Nonce = nonce, OutputRoot = observationRoot, Stage = stage, Getter = "Size", DetachedFrameClone = true,
                                OwnerPaths = new[] { "", fonts.Single(item => item.Type == 14).OwnerPath }
                            };
                            FormFontObservation.ValidateAttachedManifest(attached, path, expected, before);
                            WriteAttachedNew(attachedPath, attached);
                            report["AttachedManifestSha256"] = ExcelVbeFixture.EmbeddedRawHash(attachedPath);
                            var repository = new MacroGitRepository(Path.Combine(output, "local.git"), "qualification-layout");
                            repository.Initialize(Path.Combine(output, "unused-local-origin.git"));
                            RunLocalGit(repository, "config", "user.name", "VBAi Attached Font Diagnostic");
                            RunLocalGit(repository, "config", "user.email", "qualification@example.invalid");
                            repository.Commit(expected, null, "Frozen attached font observation input");
                            string checkpoint = repository.Checkpoint(before, "Frozen native attached font baseline").Id;
                            var step = OwnerGitQualificationScope.Step("checkpoint_restore", Path.Combine(output, "expected"),
                                Path.Combine(output, "before"), checkpoint, stage == "baseline-only" ? FormFontObservation.AttachedBaselineStop : null);
                            // Existing schema requires the ordinary three-step bank when import is admitted.
                            // Only its FIRST claim is emitted here; later steps are explicitly outside this diagnostic.
                            var steps = stage == "baseline-only" ? new[] { step } : new[]
                            {
                                step, OwnerGitQualificationScope.Step("controlled_interruption", Path.Combine(output, "before"), Path.Combine(output, "expected")),
                                OwnerGitQualificationScope.Step("rollback", Path.Combine(output, "expected"), Path.Combine(output, "before"))
                            };
                            owner.Publish(host, path, "local.git", "qualification-layout", steps);
                            report["OwnerExecutionManifest"] = owner.ManifestPath;
                            report["UnemittedStepIds"] = steps.Skip(1).Select(item => item.Id).ToArray();
                            using (var operations = new MacroGitOperations(project, repository))
                            {
                                string revision = operations.Revision(expected); report["FrozenAdmissionRevision"] = revision;
                                report["NativeStepsEmitted"] = 1; WriteReport(output, report);
                                Exception failure = null;
                                try { owner.Execute(host, step, revision); }
                                catch (Exception error) { failure = error; }
                                string terminalPath = Path.Combine(output, "owner-git-" + step.Id + ".terminal.json");
                                if (!File.Exists(terminalPath)) { host.PreserveForDiagnosticRecovery = true; throw failure ?? new IOException("No settled owner terminal."); }
                                var terminal = (IDictionary<string, object>)Json.DeserializeObject(File.ReadAllText(terminalPath));
                                OwnerGitQualificationScope.ValidateTerminal(step.Id, terminal);
                                report["OwnerTerminal"] = terminal; report["OwnerTerminalPath"] = terminalPath;
                                Assert.IsTrue(Directory.Exists(observationRoot));
                                Assert.AreEqual(1, Directory.GetFiles(observationRoot, "*-completed.json").Length,
                                    "No additional observation is allowed to make the case appear complete.");
                                if (stage == "baseline-only")
                                {
                                    Assert.IsNotNull(failure); Assert.IsTrue(failure.Message.Contains(FormFontObservation.AttachedBaselineStop));
                                    Assert.AreEqual(false, terminal["MutationStarted"]); Assert.AreEqual(false, terminal["RecoveryPending"]);
                                    Assert.IsTrue((string)terminal["Outcome"] == "FailedBeforeMutation" || (string)terminal["Outcome"] == "ExpectedPrewriteRefusal");
                                    Assert.AreEqual(terminal["BackupCommitBefore"], terminal["BackupCommitAfter"]);
                                    Assert.AreEqual(terminal["AfterImportCommitBefore"], terminal["AfterImportCommitAfter"]);
                                    report["Stage"] = "ATTACHED_BASELINE_OBSERVED_KNOWN_PREWRITE_STOP";
                                }
                                else
                                {
                                    Assert.AreEqual(true, terminal["MutationStarted"]);
                                    if (failure != null)
                                    {
                                        Assert.AreEqual("FailedAfterMutationAdmission", terminal["Outcome"]);
                                        Assert.IsTrue(failure.Message.Contains(UiText.Get("The VBE did not preserve the imported sources exactly. Use Restore or check the project.")));
                                        report["Stage"] = "ATTACHED_DELIVERY_OBSERVED_SETTLED_STRICT_REFUSAL";
                                        host.PreserveForDiagnosticRecovery = true; WriteReport(output, report); throw failure;
                                    }
                                    Assert.AreEqual("Succeeded", terminal["Outcome"]);
                                    report["Stage"] = "ATTACHED_DELIVERY_OBSERVED_EXACT_IMPORT";
                                }
                                WriteReport(output, report);
                            }
                        });
                    }, host => { report["Shutdown"] = host.ShutdownDiagnostics; WriteReport(output, report); });
                }
                catch (Exception error) { report["Error"] = error.ToString(); WriteReport(output, report); throw; }
                finally
                {
                    SynchronizationContext.SetSynchronizationContext(previousContext);
                    Environment.SetEnvironmentVariable(FormFontObservation.AttachedManifestVariable, previous);
                }
            }
        }

        private static void WriteAttachedNew(string path, object value)
        {
            byte[] bytes = new UTF8Encoding(false, true).GetBytes(Json.Serialize(value));
            using (var stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.Read))
            { stream.Write(bytes, 0, bytes.Length); stream.Flush(true); }
        }
    }
}
