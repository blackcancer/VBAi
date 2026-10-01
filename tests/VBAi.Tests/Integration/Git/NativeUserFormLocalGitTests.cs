using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    /// <summary>Opt-in native form-layout qualification through production local Git and guarded recovery.</summary>
    [TestClass, TestCategory("Excel"), TestCategory("NativeUserFormLocalGit"), DoNotParallelize]
    public sealed partial class NativeUserFormLocalGitTests
    {
        public TestContext TestContext { get; set; }
        private static readonly JavaScriptSerializer Json = new JavaScriptSerializer();

        [STATestMethod]
        [DataRow("LabelButton")]
        [DataRow("TextBox")]
        [DataRow("ComboBox")]
        [DataRow("ListBox")]
        [DataRow("CheckBox")]
        [DataRow("OptionButton")]
        [DataRow("ToggleButton")]
        [DataRow("ScrollBar")]
        [DataRow("SpinButton")]
        [DataRow("TabStrip")]
        [DataRow("Image")]
        [DataRow("FrameMultiPage")]
        public void OwnedLayoutCaptureImportRecoveryAndReopenPreserveNativeState(string layout)
        {
            RunLayout(layout, ExcelVbeFixture.Run, "ComActivation");
        }

        private void RunLayout(string layout, Action<Action<ExcelVbeFixture>, Action<ExcelVbeFixture>> run, string launchContext)
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_USERFORM_LOCAL_GIT_TESTS") != "1")
                Assert.Inconclusive("Set VBAi_RUN_USERFORM_LOCAL_GIT_TESTS=1 and VBAi_RUN_EXCEL_TESTS=1 for disposable native local-Git layout qualification.");
            if (Environment.GetEnvironmentVariable("VBAi_RUN_EXCEL_TESTS") != "1")
                Assert.Inconclusive("Owned Excel automation requires its separate explicit opt-in.");
            string root = Environment.GetEnvironmentVariable("VBAi_TEST_USERFORM_LOCAL_GIT_OUTPUT");
            if (string.IsNullOrWhiteSpace(root)) root = TestContext.TestRunResultsDirectory;
            if (string.IsNullOrWhiteSpace(root)) root = Path.Combine(Path.GetTempPath(), "VBAi-UserFormLocalGit");
            Assert.IsTrue(Path.IsPathRooted(root), "The optional durable evidence root must be absolute.");
            string output = Path.Combine(root, layout + "-" + Guid.NewGuid().ToString("N"));
            Assert.IsFalse(Directory.Exists(output));
            Directory.CreateDirectory(output);
            var report = new Dictionary<string, object> {
                ["Stage"] = "STARTED", ["Layout"] = layout, ["RemoteOperations"] = 0,
                ["LaunchContext"] = launchContext,
                ["EvidenceDirectory"] = output,
                ["MacroExecutions"] = 0, ["ForcedTermination"] = false,
                ["AssemblyMvid"] = typeof(VbeSession).Module.ModuleVersionId.ToString("D"),
                ["StartedUtc"] = DateTime.UtcNow.ToString("o"),
                ["Scope"] = "Owned Excel; production Git adapter/coordinator on external owning test STA; native readback and normal shutdown. No GitHub or embedded Git UI claim."
            };
            var previousContext = SynchronizationContext.Current;
            using (var dispatcher = new Control())
            {
                dispatcher.CreateControl();
                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
                try
                {
                    run(host => {
                        report["HostProcessId"] = host.ProcessId;
                        report["FixtureRoot"] = host.Root;
                        report["HostStatus"] = host.Command("status");
                        string path = host.File("local-git-layout.xlsm");
                        const string form = "QualificationForm";
                        Phase(output, report, "prepare-native-layout");
                        host.PrepareGitLayout(form, layout, path);
                        File.Copy(path, Path.Combine(output, "before-import.xlsm"));
                        var nativeBefore = host.ReadGitLayout(form, layout);
                        report["NativeBefore"] = nativeBefore;
                        WriteReport(output, report);
                        string rootNames = layout == "LabelButton" ? "QualificationButton|QualificationLabel" :
                            layout == "FrameMultiPage" ? "QualificationButton|QualificationExtra|QualificationLabel|QualificationMultiPage|QualificationNestedText" :
                            "QualificationButton|QualificationExtra|QualificationLabel";
                        Assert.AreEqual(rootNames, nativeBefore["Root.ControlNames"], "Exact native root collection, including descendant indexing.");
                        Assert.AreEqual(layout == "LabelButton" ? "QualificationButton|QualificationLabel" : "QualificationButton|QualificationExtra|QualificationLabel",
                            nativeBefore["Root.DirectControlNames"], "Only controls whose observed parent is the UserForm are direct root children.");
                        Assert.AreEqual(true, nativeBefore["QualificationLabel.ParentVerified"]);
                        Assert.AreEqual(true, nativeBefore["QualificationButton.ParentVerified"]);
                        if (layout != "LabelButton") Assert.AreEqual(true, nativeBefore["QualificationExtra.ParentVerified"]);
                        if (layout == "FrameMultiPage")
                        {
                            Assert.AreEqual("QualificationMultiPage|QualificationNestedText", nativeBefore["Frame.ControlNames"], "Exact indexed Frame subtree.");
                            Assert.AreEqual("QualificationMultiPage", nativeBefore["Frame.DirectControlNames"]);
                            Assert.AreEqual(true, nativeBefore["MultiPage.ParentVerified"]);
                            Assert.AreEqual(true, nativeBefore["NestedText.ParentVerified"]);
                            Assert.IsTrue(Convert.ToInt32(nativeBefore["MultiPage.PageCount"]) > 0);
                            for (int i = 0; i < Convert.ToInt32(nativeBefore["MultiPage.PageCount"]); i++)
                            {
                                string pagePrefix = "MultiPage.Page." + i;
                                Assert.AreEqual(true, nativeBefore[pagePrefix + ".ParentVerified"]);
                                Assert.AreEqual(i == 0 ? "QualificationNestedText" : "", nativeBefore[pagePrefix + ".ControlNames"]);
                                Assert.AreEqual(i == 0 ? "QualificationNestedText" : "", nativeBefore[pagePrefix + ".DirectControlNames"]);
                            }
                            Assert.AreEqual("Original nested text", nativeBefore["NestedText.Text"]);
                        }
                        host.CaptureGitFormDesigner(form, Path.Combine(output, "source-designer.png"));
                        report["DesignerCaptureAcceptance"] = "CAPTURED_PENDING_VISUAL_REVIEW";
                        host.WithGitProject(path, project => {
                            Phase(output, report, "repeat-unchanged-native-captures");
                            var before = Capture(project, output, "before");
                            Assert.IsTrue(before.Manifest.Components.Single(c => c.Name == form).HasResources);
                            for (int i = 1; i <= 3; i++)
                            {
                                var repeated = Capture(project, output, "unchanged-" + i);
                                Assert.IsTrue(before.SameAs(repeated), "Unedited native capture drift for " + layout + " at repetition " + i + ". Unsupported grammar must not weaken this comparison.");
                            }
                            report["UnchangedCapturesVerified"] = true;

                            Phase(output, report, "native-meaningful-property-change");
                            host.MutateGitLayout(form, layout);
                            var nativeChanged = host.ReadGitLayout(form, layout);
                            report["NativeChanged"] = nativeChanged;
                            Assert.IsTrue(nativeBefore.Any(pair => !Equals(pair.Value, nativeChanged[pair.Key])), "The fixture must actually change a persisted native value.");
                            var changed = Capture(project, output, "changed");
                            Assert.IsFalse(before.SameAs(changed), "The property change was lost by snapshot comparison: " + layout);
                            var changedAgain = Capture(project, output, "changed-repeat");
                            Assert.IsTrue(changed.SameAs(changedAgain), "The changed native state must also have a stable revision: " + layout);
                            report["MeaningfulChangeDetected"] = true;

                            Phase(output, report, "local-git-exact-transport");
                            var repository = new MacroGitRepository(Path.Combine(output, "local.git"), "qualification-layout");
                            // Initialize/configure only a local cache. No fetch/push operation occurs.
                            repository.Initialize(Path.Combine(output, "unused-local-origin.git"));
                            RunLocalGit(repository, "config", "user.name", "VBAi Native Layout Qualification");
                            RunLocalGit(repository, "config", "user.email", "qualification@example.invalid");
                            string commit = repository.Commit(changed, null, "Synthetic native layout " + layout);
                            var stored = repository.Read(commit);
                            AssertExactTransport(changed, stored);
                            SaveSnapshot(output, "git-readback", stored);
                            report["RawGitTransportVerified"] = true;
                            report["LocalCommit"] = commit;
                            string checkpoint = repository.Checkpoint(before, "Initial native " + layout).Id;

                            using (var operations = new MacroGitOperations(project, repository))
                            {
                                Phase(output, report, "production-checkpoint-restore");
                                try { Await(operations.ExecuteAsync("checkpoint_restore", operations.Revision(changed), name: checkpoint)); }
                                catch (InvalidOperationException error) when (error.Message.Contains("The VBE did not preserve imported sources exactly."))
                                {
                                    // This specific terminal production refusal already retains
                                    // Backup/AfterImport. Observe the actual designer once; do not
                                    // retry Apply, Save, rollback or recovery to obtain a pass.
                                    report["TerminalImportRefusal"] = error.ToString();
                                    report["RecoveryPendingAfterRefusal"] = repository.RecoveryPending;
                                    WriteReport(output, report);
                                    try
                                    {
                                        var observed = host.ReadGitLayout(form, layout);
                                        report["NativeAfterRefusedImport"] = observed;
                                        report["NativeAfterRefusedImportDifferences"] = nativeBefore.Keys.Union(observed.Keys)
                                            .Where(key => !nativeBefore.ContainsKey(key) || !observed.ContainsKey(key) || !Equals(nativeBefore[key], observed[key]))
                                            .Select(key => new { Property = key, Expected = nativeBefore.ContainsKey(key) ? nativeBefore[key] : null,
                                                Actual = observed.ContainsKey(key) ? observed[key] : null }).ToArray();
                                        WriteReport(output, report);
                                    }
                                    catch (Exception observationError)
                                    {
                                        report["PostRefusalObservationError"] = observationError.ToString();
                                        WriteReport(output, report);
                                        throw new AggregateException("Terminal strict import refusal and its read-only observation both failed.", error, observationError);
                                    }
                                    throw;
                                }
                                var restored = Capture(project, output, "checkpoint-restored");
                                Assert.IsTrue(restored.SameAs(before));
                                AssertNativeState(nativeBefore, host.ReadGitLayout(form, layout), "checkpoint restore");
                                Assert.IsFalse(repository.RecoveryPending);
                                Assert.IsTrue(repository.Read(repository.Resolve(MacroGitRepository.Backup)).SameAs(changed));
                                report["CheckpointRestoreAndBackupVerified"] = true;

                                Phase(output, report, "controlled-interruption-for-explicit-rollback");
                                // Prepare a real backup, apply once, and record the ACTUAL
                                // measured after-state. This deliberately stops before recovery
                                // completion; it is not a simulated COM error or fabricated state.
                                repository.PrepareRecovery(restored);
                                project.Apply(changed, restored);
                                var after = Capture(project, output, "measured-interrupted-state");
                                Assert.IsTrue(after.SameAs(changed));
                                AssertNativeState(nativeChanged, host.ReadGitLayout(form, layout), "native apply before rollback");
                                repository.RecordImportedState(after);
                                Assert.IsTrue(repository.RecoveryPending);
                                Assert.IsTrue(repository.Read(repository.Resolve(MacroGitRepository.AfterImport)).SameAs(after));
                                Assert.IsTrue(repository.Read(repository.Resolve(MacroGitRepository.Backup)).SameAs(before));
                                report["Interruption"] = "Explicitly prepared via production backup/apply/measured-after APIs; no uncertain mutation or failure injection";

                                Phase(output, report, "production-explicit-rollback");
                                Await(operations.ExecuteAsync("rollback", operations.Revision(after)));
                                Assert.IsTrue(Capture(project, output, "rollback-restored").SameAs(before));
                                AssertNativeState(nativeBefore, host.ReadGitLayout(form, layout), "explicit rollback");
                                Assert.IsFalse(repository.RecoveryPending);
                                report["ExplicitRollbackVerified"] = true;
                            }

                            report["ExpectedReopenedSnapshot"] = Describe(before);
                            // Retain immutable test data, never an RCW across workbook close.
                            SaveSnapshot(output, "expected-reopened", before);
                        });

                        Phase(output, report, "helper-save-and-reopen-without-macros");
                        host.SaveAndReopenGitLayout(path);
                        var reopened = host.ReadGitLayout(form, layout);
                        AssertNativeState(nativeBefore, reopened, "save/reopen");
                        report["NativeReopened"] = reopened;
                        host.WithGitProject(path, project => {
                            var expected = VbaGitSnapshot.Read(Directory.GetFiles(Path.Combine(output, "expected-reopened"))
                                .ToDictionary(file => Path.GetFileName(file), file => File.ReadAllBytes(file), StringComparer.Ordinal));
                            Assert.IsTrue(Capture(project, output, "reopened").SameAs(expected));
                        });
                        File.Copy(path, Path.Combine(output, "saved-reopened.xlsm"));
                        host.CaptureGitFormDesigner(form, Path.Combine(output, "reopened-designer.png"));
                        report["HelperSaveReopenVerified"] = true;
                        Phase(output, report, "awaiting-normal-owned-shutdown");
                    }, host => { report["Shutdown"] = host.ShutdownDiagnostics; });
                    report["NormalShutdownVerified"] = true;
                    report["Stage"] = "PASS";
                    TestContext.WriteLine("PASS native local-Git layout: " + layout + ". Designer captures require separate visual review. No remote activity.");
                }
                catch (Exception error) { report["Failure"] = error.ToString(); throw; }
                finally
                {
                    try
                    {
                        report["FinishedUtc"] = DateTime.UtcNow.ToString("o");
                        WriteReport(output, report);
                        TestContext.AddResultFile(Path.Combine(output, "local-userform-git.json"));
                    }
                    finally { SynchronizationContext.SetSynchronizationContext(previousContext); }
                }
            }
        }

        private static void AssertNativeState(IDictionary<string, object> expected, IDictionary<string, object> actual, string phase)
        {
            CollectionAssert.AreEquivalent(expected.Keys.ToArray(), actual.Keys.ToArray(), phase + " native keys");
            foreach (var item in expected) Assert.AreEqual(item.Value, actual[item.Key], phase + " native value: " + item.Key);
        }

        private static void AssertExactTransport(VbaGitSnapshot expected, VbaGitSnapshot actual)
        {
            Assert.IsNotNull(actual);
            CollectionAssert.AreEquivalent(expected.Files.Keys.ToArray(), actual.Files.Keys.ToArray());
            foreach (var item in expected.Files) CollectionAssert.AreEqual(item.Value, actual.Files[item.Key], "Raw local Git bytes: " + item.Key);
            Assert.IsTrue(expected.SameAs(actual));
        }

        private static VbaGitSnapshot Capture(VbaGitProject project, string output, string name)
        {
            var snapshot = project.Capture();
            SaveSnapshot(output, name, snapshot);
            return snapshot;
        }

        private static object[] Describe(VbaGitSnapshot snapshot)
        {
            return snapshot.Serialize().Select(file => {
                using (var hash = SHA256.Create()) return (object)new {
                    Path = file.Key, Bytes = file.Value.Length,
                    Sha256 = BitConverter.ToString(hash.ComputeHash(file.Value)).Replace("-", "")
                };
            }).ToArray();
        }

        private static void SaveSnapshot(string output, string name, VbaGitSnapshot snapshot)
        {
            string directory = Path.Combine(output, name);
            Directory.CreateDirectory(directory);
            foreach (var file in snapshot.Serialize()) File.WriteAllBytes(Path.Combine(directory, file.Key), file.Value);
            File.WriteAllText(Path.Combine(output, name + "-hashes.json"), Json.Serialize(Describe(snapshot)));
        }

        private static void RunLocalGit(MacroGitRepository repository, params string[] arguments)
        {
            typeof(MacroGitRepository).GetMethod("Run", BindingFlags.Instance | BindingFlags.NonPublic)
                .Invoke(repository, new object[] { arguments, null, true, false });
        }

        private static T Await<T>(Task<T> task)
        {
            var deadline = Stopwatch.StartNew();
            while (!task.IsCompleted)
            {
                if (deadline.Elapsed > TimeSpan.FromMinutes(3))
                    throw new TimeoutException("Native local-Git operation timed out; it was not retried.");
                Application.DoEvents(); Thread.Sleep(10);
            }
            return task.GetAwaiter().GetResult();
        }

        private static void Phase(string output, IDictionary<string, object> report, string phase)
        { report["Stage"] = phase; WriteReport(output, report); }

        private static void WriteReport(string output, IDictionary<string, object> report)
        { File.WriteAllText(Path.Combine(output, "local-userform-git.json"), Json.Serialize(report)); }
    }
}
