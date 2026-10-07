using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;

namespace VBAi.Tests.Integration
{
    /// <summary>Pairs fixed-allowlist visibility on the real owner STA with one production export in the same explicitly bootstrapped Excel.</summary>
    [TestClass, TestCategory("Excel"), TestCategory("ExcelPairedExportPathVisibility"), DoNotParallelize]
    public sealed class ExcelPairedExportPathVisibilityTests
    {
        public TestContext TestContext { get; set; }
        private static readonly List<object> retained = new List<object>();

        [STATestMethod]
        [DataRow("LocalAppData")]
        [DataRow("TEMP")]
        public void OwnerStaVisibilityImmediatelyPrecedesOneExportInOwnedGuidDirectory(string destinationKind)
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_EXCEL_PAIRED_EXPORT_PATH_VISIBILITY") != "1" ||
                Environment.GetEnvironmentVariable("VBAi_RUN_EXCEL_TESTS") != "1")
                Assert.Inconclusive("Paired native diagnosis requires both VBAi_RUN_EXCEL_PAIRED_EXPORT_PATH_VISIBILITY=1 and VBAi_RUN_EXCEL_TESTS=1.");
            string output = Environment.GetEnvironmentVariable("VBAi_EXCEL_RESULTS");
            if (string.IsNullOrWhiteSpace(output)) Assert.Inconclusive("A durable absolute VBAi_EXCEL_RESULTS directory is required.");
            output = ExcelOwnedBootstrapPlan.RequireLocalAbsolutePath(output);
            Assert.IsTrue(destinationKind == "LocalAppData" || destinationKind == "TEMP");
            const string form = "ASCIIForm";
            string trial = Path.Combine(output, "paired-" + destinationKind + "-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(trial);
            string local = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), Guid.NewGuid().ToString("N"));
            string temp = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            string manifest = Path.Combine(trial, Guid.NewGuid().ToString("N") + ".visibility.json");
            string destination = Path.Combine(destinationKind == "LocalAppData" ? local : temp, form + ".frm");
            string reportPath = Path.Combine(trial, "paired-export.json");
            var json = new JavaScriptSerializer { MaxJsonLength = 4 * 1024 * 1024 };
            var state = new PairedExportQualificationState();
            var leases = new List<FileStream>();
            ExcelVbeFixture host = null;
            bool bootstrapPending = false;
            var report = new Dictionary<string, object>
            {
                ["Scope"] = "Two owned GUID paths observed immediately before one export on the same bridge dispatcher/owner STA; no atomic token-context claim, Git normalization/import or save/reopen acceptance.",
                ["DestinationKind"] = destinationKind,
                ["Destination"] = destination,
                ["Manifest"] = manifest,
                ["MacroExecutions"] = 0,
                ["Saves"] = 0,
                ["AttributeAclOrTokenMutations"] = 0,
                ["LaunchMechanism"] = "Explicit owned /x /automation with child manifest; not COM activation",
                ["AssemblyMvid"] = typeof(VbeSession).Module.ModuleVersionId.ToString("D"),
                ["ReferencedAssemblySha256"] = Hash(typeof(VbeSession).Assembly.Location),
                ["TestAssemblySha256"] = Hash(typeof(ExcelPairedExportPathVisibilityTests).Assembly.Location)
            };
            Action save = () =>
            {
                report["Stage"] = state.Stage; report["DeliveryPending"] = state.Pending;
                report["BootstrapPending"] = bootstrapPending;
                report["NativeExportRequests"] = state.ExportRequests;
                report["TerminalExportOutcome"] = state.TerminalExportOutcome;
                report["ShutdownCompleted"] = state.ShutdownCompleted;
                File.WriteAllText(reportPath, json.Serialize(report), new UTF8Encoding(false));
            };
            var lifecycle = new ExcelScalarQualificationEvidence(Path.Combine(trial, "scenario-lifecycle.json"), 0, trial, "One export paired with owner-STA synthetic visibility, no retry or host cleanup on uncertainty");
            ExcelScalarQualificationEvidence requests = null;
            Func<object, IDictionary<string, object>> send = request =>
            {
                string command = Convert.ToString(VbeBridgeClient.Object(json.DeserializeObject(json.Serialize(request)))["Command"]);
                state.Begin(command); save();
                return requests.Send(request, () => VbeBridgeClient.Read("VBAi." + host.ProcessId, request, 20000), response =>
                {
                    state.Receive(command, response); report["LastResponse"] = response;
                    if (command == "export_component") { report["ExportResponse"] = response; report["ExportTerminalStage"] = state.Stage; }
                    save();
                    if (!Equals(true, response["Ok"])) throw new InvalidOperationException("Terminal production " + command + " failure: " + json.Serialize(response));
                });
            };
            Func<object, IDictionary<string, object>> data = request => VbeBridgeClient.Object(send(request)["Data"]);
            lifecycle.Run(() =>
            {
                foreach (string directory in new[] { local, temp })
                {
                    Assert.IsFalse(Directory.Exists(directory)); Directory.CreateDirectory(directory);
                    string synthetic = Path.Combine(directory, PathVisibilityDiagnostic.SyntheticName);
                    using (var created = new FileStream(synthetic, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                        created.Write(new byte[] { 86, 66, 65, 105 }, 0, 4);
                    leases.Add(new FileStream(synthetic, FileMode.Open, FileAccess.Read, FileShare.Read));
                }
                string manifestJson = json.Serialize(new { Version = 1, LocalAppData = local, Temp = temp });
                string[] paths = PathVisibilityDiagnostic.ValidateManifest(manifestJson, Path.GetDirectoryName(local), Path.GetDirectoryName(temp));
                using (var created = new FileStream(manifest, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                { byte[] bytes = new UTF8Encoding(false).GetBytes(manifestJson); created.Write(bytes, 0, bytes.Length); }
                leases.Add(new FileStream(manifest, FileMode.Open, FileAccess.Read, FileShare.Read));
                report["SyntheticAllowlist"] = paths; report["SyntheticSha256Before"] = new[] { Hash(paths[1]), Hash(paths[3]) };
                report["TestHostBeforeLaunch"] = PathVisibilityObservation.Read(paths); save();
                bootstrapPending = true;
                host = ExcelVbeFixture.StartOwnedWithTrace(Path.Combine(trial, "inspection-phases.jsonl"), manifest);
                bootstrapPending = false;
                report["HostProcessId"] = host.ProcessId; report["FixtureRoot"] = host.Root;
                requests = new ExcelScalarQualificationEvidence(Path.Combine(trial, "request-ledger.json"), host.ProcessId, host.Root,
                    "Guarded creation of one synthetic Form/Label, read-only evidence, parameter-free owner observation then one export");
                var loaded = data(new { Command = "status" }); report["LoadedStatus"] = loaded;
                Assert.AreEqual(host.ProcessId, Convert.ToInt32(loaded["HostProcessId"]));
                Assert.AreEqual(report["AssemblyMvid"], loaded["AssemblyModuleVersionId"]);
                report["LoadedAssemblySha256"] = Hash(Convert.ToString(loaded["AssemblyPath"]));
                Assert.AreEqual(report["ReferencedAssemblySha256"], report["LoadedAssemblySha256"]);
                var projects = (object[])send(new { Command = "list_projects" })["Data"];
                Assert.AreEqual(1, projects.Length);
                string project = Convert.ToString(VbeBridgeClient.Object(projects[0])["Name"]);
                data(new { Command = "create_form", Project = project, Form = form });
                var formState = data(new { Command = "form_state", Project = project, Form = form });
                data(new
                {
                    Command = "add_form_control",
                    Project = project,
                    Form = form,
                    ExpectedFormVersion = formState["Version"],
                    ControlType = "Forms.Label.1",
                    Control = "SyntheticLabel",
                    Caption = "VBAi synthetic export",
                    Left = 12d,
                    Top = 18d,
                    Width = 96d,
                    Height = 24d
                });
                var before = host.ReadGitExportContext(form); report["Before"] = before;
                Assert.AreEqual(project, before["ProjectName"]);
                report["NativeLabelBefore"] = host.ReadPairedExportLabel(form);
                var source = data(new { Command = "read_module", Project = project, Module = form }); report["SourceBefore"] = source;
                var tree = data(new { Command = "form_tree", Project = project, Form = form }); report["TreeBefore"] = tree;
                var component = data(new { Command = "component_properties", Project = project, Module = form }); report["ComponentBefore"] = component;
                uint tid = host.ReadPairedExportOwnerThread(); report["ExpectedOwnerNativeThreadId"] = tid;
                Assert.IsFalse(File.Exists(destination)); Assert.IsFalse(File.Exists(Path.ChangeExtension(destination, ".frx")));
                report["TestHostImmediatelyBeforeDiagnostic"] = PathVisibilityObservation.Read(paths);
                RequireVisible(VbeBridgeClient.Object(report["TestHostImmediatelyBeforeDiagnostic"]));
                var exportRequest = new
                {
                    Command = "export_component",
                    Project = project,
                    Module = form,
                    ExpectedComponentVersion = component["Version"],
                    Path = destination
                };
                var observed = data(new { Command = PathVisibilityDiagnostic.CommandName }); report["OwnerStaImmediatelyBeforeExport"] = observed;
                state.VerifyOwner(observed, host.ProcessId, tid, Convert.ToString(report["AssemblyMvid"]), paths); save();
                report["SelectedOwnerDirectoryObservation"] = ((object[])observed["Paths"]).Select(VbeBridgeClient.Object)
                    .Single(row => Equals(Path.GetDirectoryName(destination), row["Path"]));
                // No host command or COM call intervenes between this owner observation and the single production export.
                Exception exportFailure = null;
                try { report["ExportResponse"] = send(exportRequest); }
                catch (Exception error) { exportFailure = error; report["ExportFailure"] = error.ToString(); }
                if (state.Pending) { save(); ExceptionDispatchInfo.Capture(exportFailure).Throw(); }
                // A received Ok=false is terminal: record independent readback before rethrowing the original native failure.
                try
                {
                    var after = host.ReadGitExportContext(form); report["After"] = after;
                    foreach (var item in before) Assert.AreEqual(item.Value, after[item.Key], "Native identity/state changed: " + item.Key);
                    var label = host.ReadPairedExportLabel(form); report["NativeLabelAfter"] = label;
                    foreach (var item in VbeBridgeClient.Object(report["NativeLabelBefore"])) Assert.AreEqual(item.Value, label[item.Key]);
                    var sourceAfter = data(new { Command = "read_module", Project = project, Module = form }); report["SourceAfter"] = sourceAfter;
                    Assert.AreEqual(source["Sha256"], sourceAfter["Sha256"]); Assert.AreEqual(source["Code"], sourceAfter["Code"]);
                    var treeAfter = data(new { Command = "form_tree", Project = project, Form = form }); report["TreeAfter"] = treeAfter;
                    Assert.AreEqual(tree["TreeVersion"], treeAfter["TreeVersion"]);
                    report["TestHostAfterExport"] = PathVisibilityObservation.Read(paths); RequireVisible(VbeBridgeClient.Object(report["TestHostAfterExport"]));
                    CollectionAssert.AreEqual((string[])report["SyntheticSha256Before"], new[] { Hash(paths[1]), Hash(paths[3]) });
                    report["RawExportFiles"] = new[] { destination, Path.ChangeExtension(destination, ".frx") }.Select(path => new
                    {
                        Path = path,
                        Exists = File.Exists(path),
                        Bytes = File.Exists(path) ? (long?)new FileInfo(path).Length : null,
                        Sha256 = File.Exists(path) ? Hash(path) : null
                    }).ToArray();
                    save();
                }
                catch (Exception readback) { if (exportFailure != null) throw new AggregateException("Native export and independent after-readback both failed; no retry.", exportFailure, readback); throw; }
                if (exportFailure != null) ExceptionDispatchInfo.Capture(exportFailure).Throw();
                Assert.IsTrue(File.Exists(destination) && new FileInfo(destination).Length > 0, "No nonempty native FRM.");
                Assert.IsTrue(File.Exists(Path.ChangeExtension(destination, ".frx")), "Synthetic Label native FRX companion is absent.");
            }, () =>
            {
                if (state.Pending || bootstrapPending)
                {
                    report["Recovery"] = "Uncertain startup/delivery/export; exact host and synthetic/manifest read leases retained; no Close/Quit/replay/deletion";
                    if (host != null) host.PreserveForDiagnosticRecovery = true;
                    lock (retained) { retained.AddRange(leases); if (host != null) retained.Add(host); }
                    save(); return;
                }
                if (host != null)
                {
                    try
                    {
                        host.Dispose(); report["Shutdown"] = host.ShutdownDiagnostics; lifecycle.Shutdown = "Normal owned exit verified";
                        if (state.ExportRequests == 1) state.CompleteOwnedShutdown(host.ShutdownDiagnostics);
                    }
                    catch { report["Shutdown"] = host.ShutdownDiagnostics; lock (retained) { retained.AddRange(leases); retained.Add(host); } save(); throw; }
                }
                foreach (var lease in leases) lease.Dispose();
                // Keep exact synthetic GUID paths, manifest and native/partial FRM+FRX as evidence, including failed exports.
                save();
            }, () =>
            {
                save(); TestContext.AddResultFile(reportPath);
                string requestPath = Path.Combine(trial, "request-ledger.json"); if (File.Exists(requestPath)) TestContext.AddResultFile(requestPath);
            });
        }

        private static void RequireVisible(IDictionary<string, object> observation)
        {
            foreach (var row in ((IEnumerable<object>)observation["Paths"]).Select(VbeBridgeClient.Object))
                Assert.AreEqual(true, row["NativeSucceeded"], "Testhost must continue seeing every lease-protected synthetic target.");
        }
        private static string Hash(string path)
        { using (var sha = SHA256.Create()) using (var stream = File.OpenRead(path)) return BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", ""); }
    }
}
