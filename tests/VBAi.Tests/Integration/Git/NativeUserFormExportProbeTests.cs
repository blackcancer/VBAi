using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    /// <summary>Separates native export destination and dispatch failures, using one export per owned Excel trial.</summary>
    [TestClass, TestCategory("Excel"), TestCategory("NativeUserFormExportProbe"), DoNotParallelize]
    public sealed class NativeUserFormExportProbeTests
    {
        public TestContext TestContext { get; set; }

        [STATestMethod]
        [DataRow("ExternalSta", "FixtureTemporary")]
        [DataRow("ExternalSta", "GitTemporary")]
        [DataRow("ExternalSta", "EvidenceRoot")]
        [DataRow("HostBridge", "FixtureTemporary")]
        [DataRow("HostBridge", "GitTemporary")]
        [DataRow("HostBridge", "EvidenceRoot")]
        public void SingleExportPreservesExactOwnedIdentityAndProducesFormAndResources(string dispatch, string location)
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_USERFORM_EXPORT_PROBES") != "1" ||
                Environment.GetEnvironmentVariable("VBAi_RUN_EXCEL_TESTS") != "1")
                Assert.Inconclusive("Set VBAi_RUN_USERFORM_EXPORT_PROBES=1 and VBAi_RUN_EXCEL_TESTS=1 for one export per disposable Excel trial.");
            string root = Environment.GetEnvironmentVariable("VBAi_TEST_USERFORM_LOCAL_GIT_OUTPUT");
            if (string.IsNullOrWhiteSpace(root)) root = TestContext.TestRunResultsDirectory;
            if (string.IsNullOrWhiteSpace(root)) root = Path.Combine(Path.GetTempPath(), "VBAi-UserFormExports");
            Assert.IsTrue(Path.IsPathRooted(root));
            string trial = Guid.NewGuid().ToString("N");
            string output = Path.Combine(root, dispatch + "-" + location + "-" + trial);
            Assert.IsFalse(Directory.Exists(output)); Directory.CreateDirectory(output);
            string reportPath = Path.Combine(output, "native-export.json");
            var report = new Dictionary<string, object> {
                ["Stage"] = "STARTED", ["Dispatch"] = dispatch, ["Location"] = location,
                ["AssemblyMvid"] = typeof(VbeSession).Module.ModuleVersionId.ToString("D"),
                ["NativeExportRequests"] = 0, ["MacroExecutions"] = 0, ["RemoteOperations"] = 0,
                ["Scope"] = "One native UserForm Export per owned Excel process; dispatch/path diagnosis only. No Git import, capture comparison or recovery acceptance."
            };
            var json = new JavaScriptSerializer();
            try
            {
                ExcelVbeFixture.Run(host => {
                    const string form = "QualificationForm";
                    string workbook = host.File("native-export-probe.xlsm");
                    host.PrepareGitLayout(form, "LabelButton", workbook);
                    host.CaptureGitFormDesigner(form, Path.Combine(output, "source-designer.png"));
                    var before = host.ReadGitExportContext(form);
                    report["Before"] = before; report["NativeBefore"] = host.ReadGitLayout(form, "LabelButton");
                    var status = host.Command("status"); report["HostStatus"] = status;
                    var loaded = Data(status);
                    Assert.AreEqual(host.ProcessId, Convert.ToInt32(loaded["HostProcessId"]));
                    Assert.AreEqual(report["AssemblyMvid"], loaded["AssemblyModuleVersionId"]);

                    string baseDirectory = location == "GitTemporary" ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "VBAi", "GitTemporary") :
                        location == "FixtureTemporary" ? host.Root : output;
                    string destinationDirectory = Path.Combine(baseDirectory, trial);
                    Assert.IsFalse(Directory.Exists(destinationDirectory)); Directory.CreateDirectory(destinationDirectory);
                    string destination = Path.Combine(destinationDirectory, form + ".frm");
                    report["Destination"] = destination; report["DestinationLength"] = destination.Length;
                    report["DestinationContainsNonAscii"] = destination.Any(character => character > 127);
                    Exception exportFailure = null;
                    try
                    {
                        if (dispatch == "HostBridge")
                        {
                            var state = Data(host.Command(new { Command = "component_properties", Project = before["ProjectName"], Module = form }));
                            report["ComponentBefore"] = state;
                            report["NativeExportRequests"] = 1; report["Stage"] = "ONE_NATIVE_EXPORT_PENDING";
                            System.IO.File.WriteAllText(reportPath, json.Serialize(report));
                            var response = host.Command(new { Command = "export_component", Project = before["ProjectName"], Module = form,
                                ExpectedComponentVersion = state["Version"], Path = destination });
                            report["ExportResponse"] = response;
                            Data(response);
                        }
                        else
                        {
                            report["NativeExportRequests"] = 1; report["Stage"] = "ONE_NATIVE_EXPORT_PENDING";
                            System.IO.File.WriteAllText(reportPath, json.Serialize(report));
                            host.ExportGitFormOnce(form, destination);
                        }
                        var after = host.ReadGitExportContext(form); report["After"] = after;
                        foreach (var item in before) Assert.AreEqual(item.Value, after[item.Key], "Export changed identity/state: " + item.Key);
                        var nativeAfter = host.ReadGitLayout(form, "LabelButton"); report["NativeAfter"] = nativeAfter;
                        var nativeBefore = (IDictionary<string, object>)report["NativeBefore"];
                        foreach (var item in nativeBefore) Assert.AreEqual(item.Value, nativeAfter[item.Key], "Export changed native form: " + item.Key);
                        Assert.IsTrue(System.IO.File.Exists(destination), "Native Export did not create the FRM.");
                        Assert.IsTrue(System.IO.File.Exists(Path.ChangeExtension(destination, ".frx")), "Baseline native resources are missing.");
                        report["Stage"] = "EXPORT_VERIFIED_AWAITING_NORMAL_SHUTDOWN";
                    }
                    catch (Exception error) { exportFailure = error; throw; }
                    finally
                    {
                        // Keep every partial/successful raw file. Do not retry an
                        // export that failed, switch destinations or repair it.
                        try
                        {
                            report["RawFiles"] = Directory.GetFiles(destinationDirectory).Select(file => {
                                byte[] bytes = System.IO.File.ReadAllBytes(file);
                                System.IO.File.WriteAllBytes(Path.Combine(output, "raw-" + Path.GetFileName(file)), bytes);
                                using (var sha = SHA256.Create()) return (object)new { Path = file, Bytes = bytes.Length,
                                    Sha256 = BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", "") };
                            }).ToArray();
                        }
                        catch (Exception evidenceFailure)
                        {
                            if (exportFailure != null) throw new AggregateException("Native export and raw evidence retention both failed.", exportFailure, evidenceFailure);
                            throw;
                        }
                    }
                });
                report["NormalShutdownVerified"] = true; report["Stage"] = "PASS";
            }
            catch (Exception error) { report["Failure"] = error.ToString(); throw; }
            finally
            {
                System.IO.File.WriteAllText(reportPath, json.Serialize(report));
                TestContext.AddResultFile(reportPath);
            }
        }

        private static IDictionary<string, object> Data(IDictionary<string, object> response)
        {
            Assert.IsNotNull(response, "Owned host bridge did not answer.");
            object error; response.TryGetValue("Error", out error);
            Assert.AreEqual(true, response["Ok"], Convert.ToString(error));
            return VbeBridgeClient.Object(response["Data"]);
        }
    }
}
