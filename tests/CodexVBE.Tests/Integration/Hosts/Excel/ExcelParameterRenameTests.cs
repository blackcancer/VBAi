using System;
using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Integration
{
    /// <summary>Qualifie le renommage sur un vrai module VBIDE, l'exécution des appels nommés et l'annulation.</summary>
    [TestClass]
    [TestCategory("Excel")]
    public sealed class ExcelParameterRenameTests
    {
        /// <summary>Compile implicitement et exécute une fonction privée renommée puis restaure la source initiale.</summary>
        [TestMethod]
        [STATestMethod]
        public void PrivateParameterRenamePreservesNativeCallBindingAndIsUndoable()
        {
            using (var host = ExcelVbeFixture.Start())
            {
                string project = Convert.ToString(VbeBridgeClient.Object(((object[])host.Command("list_projects")["Data"]).Single())["Name"]);
                const string module = "ParameterProbe";
                var created = host.Command(new { Command = "create_module", Project = project, Module = module, ExpectedMode = 2 });
                Assert.AreEqual(true, created["Ok"], Convert.ToString(created["Error"]));
                var initial = VbeBridgeClient.Object(host.Command(new { Command = "read_module", Project = project, Module = module })["Data"]);
                const string code = "Option Explicit\r\nPrivate Function Evaluate(ByVal value As Long) As Long\r\nEvaluate = value * 2\r\nEnd Function\r\nPublic Sub Entry(Optional ByVal factor As Long = 21, Optional ByVal suffix As String = \"ok\")\r\nThisWorkbook.Worksheets(1).Range(\"A1\").Value2 = Evaluate(value:=factor)\r\nThisWorkbook.Worksheets(1).Range(\"B1\").Value2 = suffix\r\nEnd Sub";
                var written = host.Command(new { Command = "replace_lines", Project = project, Module = module,
                    ExpectedSha256 = initial["Sha256"], StartLine = 1, Count = VbeBridgeClient.Object(created["Data"])["Lines"], Text = code });
                Assert.AreEqual(true, written["Ok"], Convert.ToString(written["Error"]));
                var before = VbeBridgeClient.Object(host.Command(new { Command = "read_module", Project = project, Module = module })["Data"]);
                string original = Convert.ToString(before["Code"]);
                int column = "Private Function Evaluate(ByVal ".Length + 1;
                var preview = host.Command(new { Command = "preview_parameter_rename", Project = project, Module = module,
                    Procedure = "Evaluate", ProcKind = 0, ExpectedSha256 = before["Sha256"], StartLine = 2, StartColumn = column, Query = "value", NewName = "amount" });
                Assert.AreEqual(true, preview["Ok"], Convert.ToString(preview["Error"]));
                Assert.AreEqual(original, VbeBridgeClient.Object(host.Command(new { Command = "read_module", Project = project, Module = module })["Data"])["Code"]);
                var properties = VbeBridgeClient.Object(host.Command(new { Command = "project_properties", Project = project })["Data"]);
                string savedPath = host.File("named-arguments.xlsm");
                var saved = host.Command(new { Command = "save_host_document_as", Project = project, Path = savedPath, ExpectedProjectVersion = properties["Version"] });
                Assert.AreEqual(true, saved["Ok"], Convert.ToString(saved["Error"]));
                // Excel keeps its document open for writing; verify an isolated saved snapshot without weakening the verifier's read lock.
                string signatureSnapshot = host.File("signature-snapshot.xlsm");
                System.IO.File.Copy(savedPath, signatureSnapshot);
                var verified = host.Command(new { Command = "verify_vba_signature_file", Path = signatureSnapshot });
                Assert.AreEqual(true, verified["Ok"], Convert.ToString(verified["Error"]));
                var signature = VbeBridgeClient.Object(verified["Data"]);
                Assert.AreEqual(Convert.ToBoolean(signature["Available"]) ? "NoSignature" : "VerifierUnavailable", signature["Status"]);
                Assert.IsNull(signature["SignatureValid"], "An unsigned file or unavailable SIP cannot become a valid VBA digest.");
                var applied = host.Command(new { Command = "apply_parameter_rename", Project = project, Module = module,
                    Procedure = "Evaluate", ProcKind = 0, ExpectedSha256 = before["Sha256"], StartLine = 2, StartColumn = column, Query = "value", NewName = "amount", ExpectedMode = 2 });
                Assert.AreEqual(true, applied["Ok"], Convert.ToString(applied["Error"]));
                var after = VbeBridgeClient.Object(host.Command(new { Command = "read_module", Project = project, Module = module })["Data"]);
                StringAssert.Contains(Convert.ToString(after["Code"]), "Evaluate(amount:=factor)");
                var queued = host.Command(new { Command = "run_procedure", Project = project, Module = module, Procedure = "Entry",
                    ExpectedSha256 = after["Sha256"], ExpectedMode = 2, Arguments = new object[] { "named call" }, ArgumentNames = new[] { "suffix" } });
                Assert.AreEqual(true, queued["Ok"], Convert.ToString(queued["Error"]));
                var operation = VbeBridgeClient.Object(queued["Data"]);
                DateTime deadline = DateTime.UtcNow.AddSeconds(15);
                while (DateTime.UtcNow < deadline && Convert.ToBoolean(operation["Pending"]))
                {
                    Thread.Sleep(100);
                    var status = host.Command(new { Command = "procedure_run_status", Project = project, Query = operation["Query"] });
                    Assert.AreEqual(true, status["Ok"], Convert.ToString(status["Error"]));
                    operation = VbeBridgeClient.Object(status["Data"]);
                }
                Assert.AreEqual("Delivered", Convert.ToString(operation["State"]), Convert.ToString(operation["Error"]));
                Assert.AreEqual(42d, Convert.ToDouble(host.ReadCell("A1")), "Independent Excel readback must prove successful named-argument binding.");
                Assert.AreEqual("named call", host.ReadCell("B1"), "The named suffix must bind while the omitted factor retains its default.");
                var undone = host.Command(new { Command = "undo_code_edit", Project = project, Module = module, ExpectedSha256 = after["Sha256"], ExpectedMode = 2 });
                Assert.AreEqual(true, undone["Ok"], Convert.ToString(undone["Error"]));
                Assert.AreEqual(original, VbeBridgeClient.Object(host.Command(new { Command = "read_module", Project = project, Module = module })["Data"])["Code"]);
            }
        }
    }
}
