using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Integration
{
    /// <summary>One owned, nonmutating bridge observation; visibility differences are evidence, not inferred causes.</summary>
    [TestClass, TestCategory("Excel")]
    public sealed class ExcelPathVisibilityTests
    {
        public TestContext TestContext { get; set; }
        // Retain read leases when host delivery is uncertain; do not allow cleanup to change the observed targets.
        private static readonly List<object> retained = new List<object>();

        [STATestMethod, TestCategory("ExcelPathVisibility")]
        public void OwnedOwnerStaObservesTwoTesthostCreatedGuidDirectoriesAndEffectiveToken()
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_EXCEL_PATH_VISIBILITY") != "1")
                Assert.Inconclusive("Readonly path qualification requires VBAi_RUN_EXCEL_PATH_VISIBILITY=1 and VBAi_RUN_EXCEL_TESTS=1.");
            string output = Environment.GetEnvironmentVariable("VBAi_EXCEL_RESULTS");
            if (string.IsNullOrWhiteSpace(output)) Assert.Inconclusive("A durable absolute VBAi_EXCEL_RESULTS directory is required.");
            output = ExcelOwnedBootstrapPlan.RequireLocalAbsolutePath(output);
            string evidenceRoot = Path.Combine(output, Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(evidenceRoot);
            string local = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), Guid.NewGuid().ToString("N"));
            string temp = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            string manifest = Path.Combine(evidenceRoot, Guid.NewGuid().ToString("N") + ".visibility.json");
            var leases = new List<FileStream>();
            var json = new JavaScriptSerializer();
            ExcelVbeFixture host = null;
            bool pending = false, bootstrapPending = false, terminal = false;
            var report = new Dictionary<string, object> { ["State"] = "Preparing", ["MacroRuns"] = 0, ["NativeExports"] = 0, ["SourceWrites"] = 0, ["AttributeOrTokenMutations"] = 0 };
            string reportPath = Path.Combine(evidenceRoot, "path-visibility.json");
            Action save = () => File.WriteAllText(reportPath, json.Serialize(report), new System.Text.UTF8Encoding(false));
            var evidence = new ExcelScalarQualificationEvidence(Path.Combine(evidenceRoot, "scenario-lifecycle.json"), 0, evidenceRoot,
                "Readonly synthetic path diagnostic; host identity is recorded after exact bootstrap attachment");
            evidence.Run(() => {
                foreach (string directory in new[] { local, temp })
                {
                    Assert.IsFalse(Directory.Exists(directory)); Directory.CreateDirectory(directory);
                    string file = Path.Combine(directory, PathVisibilityDiagnostic.SyntheticName);
                    File.WriteAllBytes(file, new byte[] { 86, 66, 65, 105 });
                    leases.Add(new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.Read));
                }
                string manifestJson = json.Serialize(new { Version = 1, LocalAppData = local, Temp = temp });
                string[] paths = PathVisibilityDiagnostic.ValidateManifest(manifestJson, Path.GetDirectoryName(local), Path.GetDirectoryName(temp));
                File.WriteAllText(manifest, manifestJson, new System.Text.UTF8Encoding(false));
                leases.Add(new FileStream(manifest, FileMode.Open, FileAccess.Read, FileShare.Read));
                report["Manifest"] = manifest; report["SyntheticAllowlist"] = paths;
                string[] beforeHashes = SyntheticHashes(paths);
                report["SyntheticSha256Before"] = beforeHashes;
                report["TestHostBefore"] = PathVisibilityObservation.Read(paths); report["State"] = "PreparedBeforeHostLaunch"; save();
                bootstrapPending = true;
                host = ExcelVbeFixture.StartOwnedWithTrace(Path.Combine(evidenceRoot, "unused-inspection-phases.jsonl"), manifest);
                bootstrapPending = false;
                report["HostProcessId"] = host.ProcessId; report["FixtureRoot"] = host.Root;
                object vbe = null, window = null;
                uint nativeOwnerTid, windowPid;
                try
                {
                    vbe = ((dynamic)UiInvoke.Field<object>(host, "application")).VBE;
                    window = ((dynamic)vbe).MainWindow;
                    nativeOwnerTid = GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(((dynamic)window).HWnd)), out windowPid);
                }
                finally { Release(window); Release(vbe); }
                Assert.AreEqual(host.ProcessId, (int)windowPid); Assert.AreNotEqual(0u, nativeOwnerTid);
                report["ExpectedOwnerNativeThreadId"] = nativeOwnerTid; report["State"] = "SingleOwnerObservationPending"; save();
                pending = true;
                var commandEvidence = new ExcelScalarQualificationEvidence(Path.Combine(evidenceRoot, "request-ledger.json"), host.ProcessId, host.Root,
                    "One readonly fixed-allowlist owner-STA diagnostic; no source writes, export, macros or token manipulation");
                var response = commandEvidence.Send(new { Command = PathVisibilityDiagnostic.CommandName },
                    () => VbeBridgeClient.Read("VBAi." + host.ProcessId, new { Command = PathVisibilityDiagnostic.CommandName }, 20000), value => {
                        Assert.IsNotNull(value, "No response; retain host and synthetic read leases without replay.");
                        pending = false;
                        Assert.AreEqual(true, value["Ok"], json.Serialize(value));
                    });
                var observed = VbeBridgeClient.Object(response["Data"]);
                report["OwnerSta"] = observed;
                Assert.AreEqual(host.ProcessId, Convert.ToInt32(observed["ProcessId"]));
                Assert.AreEqual(nativeOwnerTid, Convert.ToUInt32(observed["NativeThreadId"]));
                Assert.AreEqual("STA", observed["Apartment"]);
                Assert.AreEqual(typeof(VbeSession).Module.ModuleVersionId.ToString("D"), observed["AssemblyMvid"]);
                Assert.AreEqual(observed["ProcessId"], observed["FinalProcessId"]);
                Assert.AreEqual(observed["NativeThreadId"], observed["FinalNativeThreadId"]);
                foreach (string phase in new[] { "EffectiveTokenBefore", "EffectiveTokenAfter" })
                    Assert.AreEqual("READ", VbeBridgeClient.Object(observed[phase])["State"], "Partial token evidence is not complete diagnostic acceptance.");
                var rows = ((object[])observed["Paths"]).Select(VbeBridgeClient.Object).ToArray();
                CollectionAssert.AreEqual(paths, rows.Select(row => (string)row["Path"]).ToArray());
                foreach (var row in rows) Assert.IsTrue(row.ContainsKey("NativeLastError") && row.ContainsKey("NativeAttributes") &&
                    (row.ContainsKey("ManagedAttributes") || row.ContainsKey("ManagedAttributesError")));
                report["TestHostAfter"] = PathVisibilityObservation.Read(paths);
                var before = (IDictionary<string, object>)report["TestHostBefore"];
                var after = (IDictionary<string, object>)report["TestHostAfter"];
                Assert.AreEqual(before["ProcessId"], after["ProcessId"]); Assert.AreEqual(before["NativeThreadId"], after["NativeThreadId"]);
                var beforeRows = ((IEnumerable<object>)before["Paths"]).Cast<IDictionary<string, object>>().ToArray();
                var afterRows = ((IEnumerable<object>)after["Paths"]).Cast<IDictionary<string, object>>().ToArray();
                report["VisibilityComparison"] = rows.Select((row, index) => new {
                    Path = paths[index], TestHostBeforeExists = beforeRows[index]["DirectoryExists"],
                    OwnerExists = row["DirectoryExists"], TestHostAfterExists = afterRows[index]["DirectoryExists"],
                    TestHostBeforeNativeAttributes = beforeRows[index]["NativeAttributes"],
                    OwnerNativeAttributes = row["NativeAttributes"], TestHostAfterNativeAttributes = afterRows[index]["NativeAttributes"],
                    OwnerNativeError = row["NativeLastError"], OwnerNativeErrorMeaningful = row["NativeErrorMeaningful"]
                }).ToArray();
                foreach (var testRow in beforeRows.Concat(afterRows))
                    Assert.AreEqual(true, testRow["NativeSucceeded"], "The testhost itself must see every stable synthetic target.");
                string[] afterHashes = SyntheticHashes(paths);
                report["SyntheticSha256After"] = afterHashes;
                CollectionAssert.AreEqual(beforeHashes, afterHashes, "Synthetic targets must remain byte-stable throughout the observation.");
                report["State"] = "OwnerObservationRecorded; visibility differences are not asserted equal"; terminal = true; save();
            }, () => {
                if (pending || bootstrapPending)
                {
                    report["State"] = "UncertainPreserved; no replay/Close/Quit or target deletion";
                    if (host != null) host.PreserveForDiagnosticRecovery = true;
                    lock (retained) { retained.AddRange(leases); if (host != null) retained.Add(host); }
                    save(); return;
                }
                foreach (var lease in leases) lease.Dispose();
                if (host != null)
                {
                    host.Dispose(); report["Shutdown"] = host.ShutdownDiagnostics;
                    evidence.Shutdown = "Normal owned exit verified by fixture";
                }
                report["TerminalObservation"] = terminal; save();
                // Keep both GUID directories, fixed synthetic files and manifest as durable qualification evidence.
            }, () => {
                save(); TestContext.AddResultFile(reportPath);
                if (File.Exists(Path.Combine(evidenceRoot, "request-ledger.json"))) TestContext.AddResultFile(Path.Combine(evidenceRoot, "request-ledger.json"));
            });
        }

        private static string[] SyntheticHashes(string[] paths)
        {
            using (var sha = SHA256.Create()) return paths.Where((_, i) => i % 2 == 1).Select(path => {
                using (var stream = File.OpenRead(path)) return BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", "");
            }).ToArray();
        }
        private static void Release(object value) { if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
    }
}
