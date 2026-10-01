using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration.Hosts.Excel
{
    /// <summary>Separates native signature persistence from certificate trust using disposable owned workbooks.</summary>
    [TestClass, TestCategory("ExcelSignature"), DoNotParallelize]
    public sealed class ExcelSignatureQualificationTests
    {
        public TestContext TestContext { get; set; }
        private const string Module = "SignatureFixture";

        [STATestMethod]
        public void SavedUnsignedWorkbookHasNoSignatureAndPreservesSource()
        {
            RequireOptIn();
            string path = null, sourceHash = null;
            ExcelVbeFixture.Run(host => {
                Prepare(host, out path, out sourceHash);
                Assert.AreEqual(false, Data(host.Command(new { Command = "project_signature_status", Project = path }))["Signed"]);
                Assert.AreEqual(sourceHash, Data(host.Command(new { Command = "read_module", Project = path, Module }))["Sha256"]);
            });
            var verification = VerifyClosed(path, "unsigned-file-verification.json");
            Assert.AreEqual(true, verification["Available"]);
            Assert.AreEqual("NoSignature", verification["Status"]);
            Assert.AreEqual(false, verification["Trusted"]);
            Assert.IsNull(verification["SignatureValid"]);
        }

        [STATestMethod]
        public void NativeFirstSignaturePersistsThroughFreshOwnedProcessWithoutHelperSaving()
        {
            RequireOptIn();
            string manifest = Environment.GetEnvironmentVariable("VBAi_TEST_SIGNING_CERTIFICATE_MANIFEST");
            if (string.IsNullOrEmpty(manifest)) Assert.Inconclusive("An explicitly owned synthetic certificate manifest is required.");
            var certificate = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(File.ReadAllText(manifest));
            Assert.AreEqual(1, Convert.ToInt32(certificate["FormatVersion"]));
            string nonce = Convert.ToString(certificate["Nonce"]);
            Assert.IsTrue(System.Text.RegularExpressions.Regex.IsMatch(nonce, "^[A-F0-9]{32}$"));
            Assert.AreEqual("CN=VBAi Disposable Signature Qualification " + nonce, certificate["Subject"]);
            Assert.AreEqual("CurrentUser/My", certificate["Store"]);
            Assert.AreEqual(false, certificate["TrustedRootInstalled"]);
            string thumbprint = Convert.ToString(certificate["Thumbprint"]);
            Assert.IsTrue(System.Text.RegularExpressions.Regex.IsMatch(thumbprint, "^[A-F0-9]{40}$"));
            string path = null, sourceHash = null;
            int signedPid = 0;
            ExcelVbeFixture.Run(host => {
                signedPid = host.ProcessId;
                Prepare(host, out path, out sourceHash);
                var unsigned = Data(host.Command(new { Command = "project_signature_status", Project = path }));
                Assert.AreEqual(true, unsigned["Available"]); Assert.AreEqual(false, unsigned["Signed"]);
                var properties = Data(host.Command(new { Command = "project_properties", Project = path }));
                var signed = Data(host.SignOnce(new { Command = "sign_project", Project = path,
                    ExpectedProjectVersion = properties["Version"], ExpectedMode = 2, CertificateThumbprint = thumbprint }));
                Write(host.File("signing-result.json"), signed);
                Assert.AreEqual(true, VbeBridgeClient.Object(signed["Signature"])["SignatureAssigned"]);
                Assert.AreEqual(true, VbeBridgeClient.Object(signed["Persistence"])["Saved"]);
                Assert.AreEqual(false, signed["SaveRequired"]); Assert.IsNull(signed["PersistenceError"]);
                Assert.AreEqual(true, VbeBridgeClient.Object(signed["HostStatus"])["Signed"]);
                Assert.AreEqual(sourceHash, Data(host.Command(new { Command = "read_module", Project = path, Module }))["Sha256"]);
            });
            string savedFileHash = FileHash(path);
            var verification = VerifyClosed(path, "signed-file-verification.json");
            Assert.AreEqual(true, verification["Available"]);
            Assert.AreEqual("UntrustedRoot", verification["Status"], "The test certificate is deliberately absent from trusted roots.");
            Assert.AreEqual(false, verification["Trusted"]);
            Assert.IsNull(verification["SignatureValid"], "A trust failure must not be promoted to a valid signature result.");
            ExcelVbeFixture.Run(host => {
                Assert.AreNotEqual(signedPid, host.ProcessId, "Reopen requires a new owned native process.");
                host.OpenOwnedReadOnlyWorkbook(path);
                var status = Data(host.Command(new { Command = "project_signature_status", Project = path }));
                Assert.AreEqual(true, status["Available"]); Assert.AreEqual(true, status["Signed"]);
                Assert.AreEqual(sourceHash, Data(host.Command(new { Command = "read_module", Project = path, Module }))["Sha256"]);
                Assert.AreEqual(2, Convert.ToInt32(Data(host.Command(new { Command = "debug_state", Project = path }))["Mode"]));
                Write(host.File("signature-reopened.json"), new { Signed = status, SourceSha256 = sourceHash,
                    host.ProcessId, OriginalSignedPid = signedPid, HelperSaveInvoked = false, MacroExecuted = false });
            });
            Assert.AreEqual(savedFileHash, FileHash(path), "Read-only reopen and Close(false) must preserve the exact saved signed file.");
        }

        private static void RequireOptIn()
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_EXCEL_SIGNATURE_TESTS") != "1")
                Assert.Inconclusive("Set VBAi_RUN_EXCEL_SIGNATURE_TESTS=1 for disposable native signature qualification.");
            Assert.IsTrue(!string.IsNullOrEmpty(Environment.GetEnvironmentVariable("VBAi_EXCEL_RESULTS")), "Durable native evidence is required.");
        }
        private static void Prepare(ExcelVbeFixture host, out string path, out string sourceHash)
        {
            var projects = (object[])host.Command("list_projects")["Data"];
            string project = Convert.ToString(VbeBridgeClient.Object(projects.Single())["Name"]);
            var created = Data(host.Command(new { Command = "create_module", Project = project, Module, ExpectedMode = 2 }));
            var initial = Data(host.Command(new { Command = "read_module", Project = project, Module }));
            string code = "Option Explicit\r\nPrivate Const Marker As String = \"SIGNATURE_" + Guid.NewGuid().ToString("N") + "\"\r\nPublic Sub NeverExecuted()\r\nEnd Sub";
            Data(host.Command(new { Command = "replace_lines", Project = project, Module,
                ExpectedSha256 = initial["Sha256"], StartLine = 1, Count = created["Lines"], Text = code }));
            sourceHash = Convert.ToString(Data(host.Command(new { Command = "read_module", Project = project, Module }))["Sha256"]);
            var properties = Data(host.Command(new { Command = "project_properties", Project = project }));
            path = host.File("SyntheticSignature.xlsm");
            Data(host.Command(new { Command = "save_host_document_as", Project = project, Path = path, ExpectedProjectVersion = properties["Version"] }));
        }
        private IDictionary<string, object> VerifyClosed(string path, string evidenceName)
        {
            var json = new JavaScriptSerializer();
            var result = VbeBridgeClient.Object(json.DeserializeObject(json.Serialize(new VbeSignatureVerifier().Verify(path))));
            Write(Path.Combine(Path.GetDirectoryName(path), evidenceName), result);
            return result;
        }
        private void Write(string path, object value)
        {
            File.WriteAllText(path, new JavaScriptSerializer().Serialize(value), new UTF8Encoding(false));
            TestContext.AddResultFile(path);
        }
        private static string FileHash(string path)
        {
            using (var file = File.OpenRead(path)) using (var hash = SHA256.Create())
                return BitConverter.ToString(hash.ComputeHash(file)).Replace("-", "");
        }
        private static IDictionary<string, object> Data(IDictionary<string, object> response)
        {
            Assert.IsNotNull(response); Assert.AreEqual(true, response["Ok"], Convert.ToString(response["Error"]));
            return VbeBridgeClient.Object(response["Data"]);
        }
    }
}
