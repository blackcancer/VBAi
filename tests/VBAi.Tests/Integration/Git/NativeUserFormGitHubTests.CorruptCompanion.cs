using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    public sealed partial class NativeUserFormGitHubTests
    {
        /// <summary>Routes malformed local Git blobs through the production checkpoint restore before any native replacement.</summary>
        private void AssertCorruptCompanionRejectedByProductionRestore(string corruption)
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_USERFORM_CORRUPTION_TESTS") != "1" ||
                Environment.GetEnvironmentVariable("VBAi_RUN_EXCEL_TESTS") != "1")
                Assert.Inconclusive("Owned Excel and the separate corrupt UserForm opt-ins are required.");
            string root = Environment.GetEnvironmentVariable("VBAi_TEST_USERFORM_GIT_OUTPUT");
            Assert.IsTrue(!string.IsNullOrWhiteSpace(root) && Path.IsPathRooted(root));
            string output = Path.Combine(root, "corruption-" + corruption + "-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(output);
            var report = new Dictionary<string, object> {
                ["Stage"] = "STARTED", ["Corruption"] = corruption, ["RemoteMutation"] = false,
                ["AssemblyMvid"] = typeof(VbeSession).Module.ModuleVersionId.ToString("D"),
                ["NativeImportAttempts"] = 0, ["RecoveryAttempts"] = 0, ["MacroExecutions"] = 0
            };
            string sentinelPath = null, sentinelHash = null;
            var previousContext = SynchronizationContext.Current;
            using (var dispatcher = new Control())
            using (var owner = new OwnerGitQualificationScope(output))
            {
                dispatcher.CreateControl();
                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
                try
                {
                    ExcelVbeFixture.Run(host => {
                        report["HostProcessId"] = host.ProcessId;
                        const string form = "QualificationForm";
                        const string layout = "LabelButton";
                        string path = host.File("corruption-sentinel.xlsm");
                        sentinelPath = path;
                        host.PrepareGitLayout(form, layout, path, persistedBaseline: true);
                        var sourceBefore = host.ReadGitForm(form);
                        var nativeBefore = host.ReadGitLayout(form, layout);
                        var fontsBefore = host.ReadGitLayoutFonts(form, layout);
                        string diskBefore = CorruptCompanionFileHash(path);
                        sentinelHash = diskBefore;
                        File.Copy(path, Path.Combine(output, "sentinel-before-restore.xlsm"));
                        host.WithGitProject(path, project => {
                            var before = project.Capture();
                            SaveSnapshot(output, "sentinel-snapshot", before);
                            Assert.IsTrue(before.Manifest.Components.Single(x => x.Name == form).HasResources);
                            var repository = NewRepository(Path.Combine(output, "local.git"),
                                "qualification-corrupt-local-only", "https://github.com/blackcancer/vbai-qualification-20260929203712-7267b1e6.git", null);
                            string checkpoint = CorruptCompanionCheckpoint(repository, before, form, corruption);
                            string checkpointRef = "refs/codex/checkpoints/" + checkpoint;
                            string committed = repository.Resolve(checkpointRef);
                            string backupBefore = repository.Resolve(MacroGitRepository.Backup);
                            string afterImportBefore = repository.Resolve(MacroGitRepository.AfterImport);
                            Assert.IsNull(backupBefore, "The disposable repository must start without an import backup.");
                            Assert.IsNull(afterImportBefore, "The disposable repository must start without an after-import state.");
                            Assert.IsFalse(repository.RecoveryPending);
                            report["Stage"] = "production-checkpoint-restore-once";
                            report["Checkpoint"] = checkpoint;
                            report["MalformedCommit"] = committed;
                            WriteReport(output, report);
                            using (var operations = new MacroGitOperations(project, repository))
                            {
                                string expectedMessage = corruption == "Missing" ?
                                    UiText.Get("VBA sources are incomplete, unexpected or too large (32 MB maximum).") :
                                    corruption == "Empty" ? "Form resources are missing" :
                                    "Invalid or truncated UserForm OLE resource container";
                                var step = OwnerGitQualificationScope.Step("checkpoint_restore",
                                    Path.Combine(output, "sentinel-snapshot"), Path.Combine(output, "sentinel-snapshot"), checkpoint, expectedMessage);
                                owner.Publish(host, path, "local.git", "qualification-corrupt-local-only", new[] { step });
                                report["OwnerExecutionManifest"] = owner.ManifestPath;
                                report["OwnerExecutionManifestSha256"] = ExcelVbeFixture.EmbeddedRawHash(owner.ManifestPath);
                                report["OwnerExecutionReceipts"] = owner.TerminalReceipts;
                                var refusal = Assert.ThrowsException<InvalidOperationException>(() =>
                                    owner.Execute(host, step, operations.Revision(before)));
                                StringAssert.Contains(refusal.Message, expectedMessage,
                                    "The production repository preflight must be the refusal source.");
                                Assert.AreEqual(1, owner.TerminalReceipts.Count);
                                var terminal = Json.Deserialize<Dictionary<string, object>>(File.ReadAllText(owner.TerminalReceipts[0]));
                                Assert.AreEqual("ExpectedPrewriteRefusal", terminal["Outcome"]);
                                Assert.AreEqual(false, terminal["MutationStarted"]);
                                Assert.AreEqual(false, terminal["RecoveryPending"]);
                                Assert.IsFalse(host.PreserveForDiagnosticRecovery, "A proved prewrite refusal permits normal owned exit.");
                                report["PreflightError"] = refusal.Message;
                            }
                            Assert.AreEqual(committed, repository.Resolve(checkpointRef), "The malformed input must remain available as evidence.");
                            Assert.AreEqual(backupBefore, repository.Resolve(MacroGitRepository.Backup), "Refusal must precede backup creation.");
                            Assert.AreEqual(afterImportBefore, repository.Resolve(MacroGitRepository.AfterImport), "Refusal must precede recovery state creation.");
                            Assert.IsFalse(repository.RecoveryPending, "Refusal must precede the recovery marker.");
                            var after = project.Capture();
                            Assert.IsTrue(before.SameAs(after), "The sentinel snapshot changed before native import.");
                            CorruptCompanionAssertNative(sourceBefore, host.ReadGitForm(form), "source");
                            CorruptCompanionAssertNative(nativeBefore, host.ReadGitLayout(form, layout), "layout");
                            CorruptCompanionAssertNative(fontsBefore, host.ReadGitLayoutFonts(form, layout), "fonts");
                            Assert.AreEqual(diskBefore, CorruptCompanionFileHash(path), "The owned workbook changed before native import.");
                            report["SentinelSnapshotUnchanged"] = true;
                            report["SentinelSourceUnchanged"] = true;
                            report["SentinelLayoutUnchanged"] = true;
                            report["SentinelFontsUnchanged"] = true;
                            report["SentinelDiskUnchanged"] = true;
                            report["BackupAndRecoveryUnchanged"] = true;
                            report["Stage"] = "REFUSED_BEFORE_MUTATION_AWAITING_NORMAL_EXIT";
                            WriteReport(output, report);
                        });
                    }, host => {
                        report["Shutdown"] = host.ShutdownDiagnostics;
                        Assert.IsNotNull(host.ShutdownDiagnostics);
                        Assert.AreEqual(true, host.ShutdownDiagnostics["Exited"]);
                        Assert.AreEqual(0, Convert.ToInt32(host.ShutdownDiagnostics["ExitCode"]));
                        Assert.AreEqual(false, host.ShutdownDiagnostics["ForcedTermination"]);
                        Assert.AreEqual(sentinelHash, CorruptCompanionFileHash(sentinelPath),
                            "Normal shutdown must preserve the sentinel workbook bytes.");
                        WriteReport(output, report);
                    });
                    report["NormalShutdownVerified"] = true;
                    report["Stage"] = "PASS";
                }
                catch (Exception error) { report["Stage"] = "FAILED"; report["Failure"] = error.ToString(); throw; }
                finally
                {
                    report["FinishedUtc"] = DateTime.UtcNow.ToString("o");
                    WriteReport(output, report);
                    TestContext.AddResultFile(Path.Combine(output, "userform-github.json"));
                    SynchronizationContext.SetSynchronizationContext(previousContext);
                }
            }
        }

        /// <summary>Builds an actual local Git commit with malformed companion bytes, without publishing it.</summary>
        private static string CorruptCompanionCheckpoint(MacroGitRepository repository, VbaGitSnapshot before,
            string form, string corruption)
        {
            string checkpoint = repository.Checkpoint(before, "Synthetic malformed FRX preflight").Id;
            var files = before.Serialize();
            string companion = form + ".frx";
            Assert.IsTrue(files.ContainsKey(companion) && files[companion].Length > 0);
            switch (corruption)
            {
                case "Missing": files.Remove(companion); break;
                case "Empty": files[companion] = new byte[0]; break;
                case "SignatureCorrupt":
                    string frm = VbaGitSnapshot.Utf8.GetString(files[form + ".frm"]);
                    var blob = Regex.Match(frm,
                        @"^[ \t]*OleObjectBlob[ \t]*=[ \t]*""[^""\r\n]+""[ \t]*:[ \t]*([0-9a-f]+)[ \t]*$",
                        RegexOptions.Multiline | RegexOptions.IgnoreCase);
                    Assert.IsTrue(blob.Success, "The native form must contain an addressed OLE blob.");
                    int offset = checked((int)uint.Parse(blob.Groups[1].Value, NumberStyles.HexNumber, CultureInfo.InvariantCulture));
                    var corrupt = (byte[])files[companion].Clone();
                    Assert.IsTrue(offset >= 0 && offset <= corrupt.Length - 32);
                    Assert.AreEqual((byte)0xd0, corrupt[offset + 24], "The fixture must begin with the native CFB signature.");
                    corrupt[offset + 24] ^= 0xff;
                    files[companion] = corrupt;
                    break;
                default: throw new ArgumentOutOfRangeException(nameof(corruption));
            }
            var entries = files.Select(file => "100644 blob " + CorruptCompanionGitText(repository,
                new[] { "hash-object", "-w", "--stdin" }, file.Value) + "\t" + file.Key).ToArray();
            string tree = CorruptCompanionGitText(repository, new[] { "mktree", "-z" },
                VbaGitSnapshot.Utf8.GetBytes(string.Join("\0", entries) + "\0"));
            string root = CorruptCompanionGitText(repository, new[] { "mktree", "-z" },
                VbaGitSnapshot.Utf8.GetBytes("040000 tree " + tree + "\tvba\0"));
            string commit = CorruptCompanionGitText(repository,
                new[] { "commit-tree", root, "-m", "Synthetic malformed FRX preflight" });
            repository.SetRef("refs/codex/checkpoints/" + checkpoint, commit);
            return checkpoint;
        }

        private static string CorruptCompanionGitText(MacroGitRepository repository, string[] arguments, byte[] input = null)
        {
            var run = typeof(MacroGitRepository).GetMethod("Run", BindingFlags.Instance | BindingFlags.NonPublic);
            Assert.IsNotNull(run);
            var result = (MacroGitRepository.Result)run.Invoke(repository, new object[] { arguments, input, true, false });
            Assert.AreEqual(0, result.ExitCode);
            return Encoding.UTF8.GetString(result.Bytes).Trim();
        }

        private static void CorruptCompanionAssertNative(IDictionary<string, object> expected,
            IDictionary<string, object> actual, string scope)
        {
            CollectionAssert.AreEquivalent(expected.Keys.ToArray(), actual.Keys.ToArray(), scope + " keys");
            foreach (var pair in expected) Assert.AreEqual(pair.Value, actual[pair.Key], scope + ": " + pair.Key);
        }

        private static string CorruptCompanionFileHash(string path)
        {
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var hash = SHA256.Create())
                return BitConverter.ToString(hash.ComputeHash(stream)).Replace("-", "");
        }
    }
}
