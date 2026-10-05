using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    public sealed partial class NativeUserFormLocalGitTests
    {
        /// <summary>Separates native Office persistence from Git import by reopening every synthetic layout without an import.</summary>
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
        public void OwnedLayoutBaselineReopenPreservesNativeStateWithoutGitImport(string layout)
        {
            RunPersistenceBaseline(layout, false);
        }

        /// <summary>Requires every saved designer baseline to remain exact across another document reopen and a fresh process.</summary>
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
        public void OwnedPersistedLayoutBaselineReopenPreservesNativeStateWithoutGitImport(string layout)
        {
            RunPersistenceBaseline(layout, true);
        }

        /// <summary>Records prepared and persisted baselines independently; all assertions remain exact.</summary>
        private void RunPersistenceBaseline(string layout, bool persistedBaseline)
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_EXCEL_TESTS") != "1" ||
                Environment.GetEnvironmentVariable("VBAi_RUN_USERFORM_LOCAL_GIT_TESTS") != "1")
                Assert.Inconclusive("Disposable Excel and native form qualification require both explicit opt-ins.");
            string root = Environment.GetEnvironmentVariable("VBAi_TEST_USERFORM_LOCAL_GIT_OUTPUT");
            Assert.IsTrue(!string.IsNullOrWhiteSpace(root) && Path.IsPathRooted(root), "A durable absolute evidence root is required.");
            string output = Path.Combine(root, layout + "-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(output);
            string saved = Path.Combine(output, "baseline.xlsm");
            const string form = "QualificationForm";
            IDictionary<string, object> before = null, sameProcess = null, freshProcess = null;
            int originalPid = 0;
            var report = new Dictionary<string, object> {
                ["Layout"] = layout, ["Stage"] = "STARTED", ["GitImports"] = 0,
                ["PersistedBaseline"] = persistedBaseline,
                ["MacroExecutions"] = 0, ["RemoteOperations"] = 0,
                ["AssemblyMvid"] = typeof(VbeSession).Module.ModuleVersionId.ToString("D"),
                ["Scope"] = "Independent native baseline persistence; one prepared Save, same-process document reopen, then read-only fresh Excel process; no Git import or repair."
            };
            string evidence = Path.Combine(output, "baseline-persistence.json");
            Action write = () => File.WriteAllText(evidence, Json.Serialize(report));
            try
            {
                ExcelVbeFixture.Run(host => {
                    originalPid = host.ProcessId;
                    report["OriginalProcessId"] = originalPid;
                    report["OriginalStatus"] = host.Command("status");
                    string path = host.File("baseline-source.xlsm");
                    host.PrepareGitLayout(form, layout, path, persistedBaseline);
                    before = host.ReadGitLayout(form, layout);
                    report["NativePrepared"] = before;
                    write();
                    if (UserFormQualificationFonts.Enabled)
                        host.WithGitProject(path, project => UserFormQualificationFonts.RequireSnapshot(project.Capture(),
                            (name, currentLayout) => host.ReadGitLayoutFonts(name, currentLayout), layout));
                    File.Copy(path, saved);
                    // PrepareGitLayout already saved once. Reopen closes without Save;
                    // neither export nor a second Save may repair the observed baseline.
                    Assert.AreEqual(0, host.ReopenAndReadProjectProtection(path));
                    sameProcess = host.ReadGitLayout(form, layout);
                    if (UserFormQualificationFonts.Enabled)
                        host.WithGitProject(path, project => UserFormQualificationFonts.RequireSnapshot(project.Capture(),
                            (name, currentLayout) => host.ReadGitLayoutFonts(name, currentLayout), layout));
                    report["NativeSameProcessReopen"] = sameProcess;
                    report["SameProcessDifferences"] = DescribeNativeDifferences(before, sameProcess);
                    report["Stage"] = "same-process-observed";
                    write();
                }, host => {
                    report["OriginalShutdown"] = host.ShutdownDiagnostics;
                    write();
                });
                // Run returned only after its retained original process handle proved
                // normal exit. The saved copy belongs to this evidence directory.
                string savedHash = BaselineFileHash(saved);
                report["SavedFileSha256"] = savedHash;
                ExcelVbeFixture.Run(host => {
                    Assert.AreNotEqual(originalPid, host.ProcessId, "A distinct Excel process is required.");
                    report["FreshProcessId"] = host.ProcessId;
                    host.OpenOwnedReadOnlyWorkbook(saved);
                    report["FreshStatus"] = host.Command("status");
                    freshProcess = host.ReadGitLayout(form, layout);
                    if (UserFormQualificationFonts.Enabled)
                        host.WithGitProject(saved, project => UserFormQualificationFonts.RequireSnapshot(project.Capture(),
                            (name, currentLayout) => host.ReadGitLayoutFonts(name, currentLayout), layout));
                    report["NativeFreshProcessReopen"] = freshProcess;
                    report["FreshProcessDifferences"] = DescribeNativeDifferences(before, freshProcess);
                    Assert.AreEqual(savedHash, BaselineFileHash(saved), "Read-only reopen/export must preserve the saved file.");
                    report["Stage"] = "fresh-process-observed";
                    write();
                }, host => {
                    report["FreshShutdown"] = host.ShutdownDiagnostics;
                    write();
                });
                AssertNativeState(before, sameProcess, "Native same-process baseline reopen");
                AssertNativeState(before, freshProcess, "Native fresh-process baseline reopen");
                report["Stage"] = "PASSED";
            }
            catch (Exception error)
            {
                report["Stage"] = "FAILED";
                report["Failure"] = error.ToString();
                throw;
            }
            finally { write(); TestContext.AddResultFile(evidence); }
        }

        /// <summary>Records all native differences before an assertion so a failed baseline does not hide other observations.</summary>
        private static object[] DescribeNativeDifferences(IDictionary<string, object> expected, IDictionary<string, object> actual)
        {
            return expected.Keys.Union(actual.Keys).Where(key => !expected.ContainsKey(key) || !actual.ContainsKey(key) ||
                !Equals(expected[key], actual[key])).Select(key => (object)new {
                    Property = key, Expected = expected.ContainsKey(key) ? expected[key] : null,
                    Actual = actual.ContainsKey(key) ? actual[key] : null
                }).ToArray();
        }

        /// <summary>Hashes the retained file while Office holds its read-only sharing handle.</summary>
        private static string BaselineFileHash(string path)
        {
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var hash = SHA256.Create())
                return BitConverter.ToString(hash.ComputeHash(stream)).Replace("-", "");
        }
    }
}
