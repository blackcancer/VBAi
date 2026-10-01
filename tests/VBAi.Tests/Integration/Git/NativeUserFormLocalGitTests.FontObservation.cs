using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    public sealed partial class NativeUserFormLocalGitTests
    {
        /// <summary>Diagnoses whether a read-only font getter materializes the descriptor omitted by native import.</summary>
        [STATestMethod]
        [DataRow("LabelButton")]
        [DataRow("Image")]
        [DataRow("FrameMultiPage")]
        public void ReadOnlyFontObservationAfterOneNativeImportPreservesExactSnapshot(string layout)
        {
            RunFontObservation(layout, false);
        }

        /// <summary>Diagnoses one explicit font restoration after the terminal strict import refusal, retaining exact comparison.</summary>
        [STATestMethod]
        [DataRow("LabelButton")]
        [DataRow("Image")]
        [DataRow("FrameMultiPage")]
        public void NativeFontRestorationAfterOneImportPreservesExactSnapshot(string layout)
        {
            RunFontObservation(layout, true);
        }

        /// <summary>Transfers the restored StdFont to the native owner once and verifies every comparison byte.</summary>
        [STATestMethod]
        [DataRow("LabelButton")]
        [DataRow("Image")]
        [DataRow("FrameMultiPage")]
        public void NativeOwnedFontAssignmentAfterOneImportPreservesExactSnapshot(string layout)
        {
            RunFontObservation(layout, true, true);
        }

        /// <summary>Separates the original refusal, read-only observation and optional single font correction in owned fixtures.</summary>
        private void RunFontObservation(string layout, bool restoreFonts, bool assignOwner = false)
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_EXCEL_TESTS") != "1" ||
                Environment.GetEnvironmentVariable("VBAi_RUN_USERFORM_LOCAL_GIT_TESTS") != "1" ||
                Environment.GetEnvironmentVariable("VBAi_RUN_USERFORM_EXPLICIT_BOOTSTRAP") != "1")
                Assert.Inconclusive("Disposable explicit Excel, UserForm and bootstrap qualification require all opt-ins.");
            string root = Environment.GetEnvironmentVariable("VBAi_TEST_USERFORM_LOCAL_GIT_OUTPUT");
            Assert.IsTrue(!string.IsNullOrWhiteSpace(root) && Path.IsPathRooted(root));
            string output = Path.Combine(root, "font-" + layout + "-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(output);
            const string form = "QualificationForm";
            var report = new Dictionary<string, object> {
                ["Layout"] = layout, ["Stage"] = "STARTED", ["NativeImports"] = 0,
                ["FontAssignmentsRequestedAfterImport"] = 0, ["MacroExecutions"] = 0, ["RemoteOperations"] = 0,
                ["AssemblyMvid"] = typeof(VbeSession).Module.ModuleVersionId.ToString("D"),
                ["RestoreFontsRequested"] = restoreFonts,
                ["AssignNativeFontOwnerRequested"] = assignOwner,
                ["Scope"] = "Saved/reopened baseline; one native Apply attempt; read-only Font observation and optional single explicit font restoration. No repeated import, setter retry, post-import Save or macro execution. This diagnostic does not qualify successful production import."
            };
            string evidence = Path.Combine(output, "font-observation.json");
            Action write = () => File.WriteAllText(evidence, Json.Serialize(report));
            try
            {
                RunExplicit(host => {
                    report["HostProcessId"] = host.ProcessId;
                    report["HostStatus"] = host.Command("status");
                    string path = host.File("font-observation.xlsm");
                    host.PrepareGitLayout(form, layout, path, persistedBaseline: true);
                    File.Copy(path, Path.Combine(output, "saved-baseline.xlsm"));
                    host.CaptureGitFormDesigner(form, Path.Combine(output, "before-designer.png"));
                    var fontBefore = host.ReadGitLayoutFonts(form, layout);
                    report["NativeFontsBefore"] = fontBefore;
                    host.WithGitProject(path, project => {
                        var before = Capture(project, output, "before");
                        host.MutateGitLayout(form, layout, persistedBaseline: true);
                        var changed = Capture(project, output, "changed");
                        Assert.IsFalse(before.SameAs(changed), "The single Apply must have a real native source change to undo.");
                        report["Stage"] = "one-native-import";
                        report["NativeImports"] = 1;
                        write();
                        try
                        {
                            project.Apply(before, changed);
                            report["InitialStrictComparison"] = "PASSED";
                        }
                        catch (InvalidOperationException error) when (error.Message == UiText.Get("The VBE did not preserve the imported sources exactly. Use Restore or check the project."))
                        {
                            report["InitialStrictComparison"] = "REFUSED";
                            report["InitialStrictRefusal"] = error.ToString();
                        }
                        var fontAfter = host.ReadGitLayoutFonts(form, layout);
                        report["NativeFontsAfter"] = fontAfter;
                        report["FontDifferences"] = DescribeNativeDifferences(fontBefore, fontAfter);
                        if (restoreFonts)
                        {
                            report["FontAssignmentsRequestedAfterImport"] = layout == "FrameMultiPage" ? 16 : 8;
                            report["Stage"] = "one-explicit-font-restoration";
                            write();
                            host.RestoreGitLayoutFonts(form, layout, fontBefore, assignOwner);
                            fontAfter = host.ReadGitLayoutFonts(form, layout);
                            report["NativeFontsAfterExplicitRestoration"] = fontAfter;
                            report["FontDifferencesAfterExplicitRestoration"] = DescribeNativeDifferences(fontBefore, fontAfter);
                        }
                        var observed = Capture(project, output, "after-font-observation");
                        report["ExactSnapshotAfterFontRead"] = before.SameAs(observed);
                        report["RemainingChangedFiles"] = observed.Changes(before);
                        report["NativeAfter"] = host.ReadGitLayout(form, layout);
                        report["Stage"] = "OBSERVED";
                        write();
                        AssertNativeState(fontBefore, fontAfter, "Read-only post-import font observation");
                        Assert.IsTrue(before.SameAs(observed), "Font observation did not restore the exact comparison snapshot; no import, Save or setter is repeated.");
                    });
                }, host => { report["Shutdown"] = host.ShutdownDiagnostics; write(); });
                report["Stage"] = "PASSED_DIAGNOSTIC_ONLY";
            }
            catch (Exception error) { report["Stage"] = "FAILED"; report["Failure"] = error.ToString(); throw; }
            finally { write(); TestContext.AddResultFile(evidence); }
        }
    }
}
