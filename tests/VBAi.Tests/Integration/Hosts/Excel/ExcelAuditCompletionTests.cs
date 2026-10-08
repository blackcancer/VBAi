using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Integration.Hosts.Excel
{
    /// <summary>Exercises remaining audit boundaries with disposable native projects and independent host readback.</summary>
    [TestClass, TestCategory("ExcelAuditCompletion"), DoNotParallelize]
    public sealed class ExcelAuditCompletionTests
    {
        public TestContext TestContext { get; set; }

        /// <summary>Runs a synthetic procedure once while another module owns the active native code pane.</summary>
        [STATestMethod]
        public void ColdRunSubFocusesItsExactTargetWithoutSeparateNavigationOrExecutionReplay()
        {
            RequireOptIn();
            ExcelVbeFixture.Run(host =>
            {
                string project = Project(host);
                const string module = "ColdAuditTarget", other = "ColdAuditOther";
                CreateModule(host, project, module,
                    "Option Explicit\r\nPublic Sub ColdMarker()\r\n" +
                    "    ThisWorkbook.Worksheets(1).Range(\"A2\").Value2 = 1\r\nEnd Sub");
                CreateModule(host, project, other, "Option Explicit");
                var source = Read(host, new { Command = "read_module", Project = project, Module = module });
                object workbook = UiInvoke.Field<object>(host, "workbook"), nativeProject = null,
                    components = null, component = null, code = null, pane = null, window = null;
                try
                {
                    nativeProject = ((dynamic)workbook).VBProject;
                    components = ((dynamic)nativeProject).VBComponents;
                    component = ((dynamic)components).Item(other);
                    code = ((dynamic)component).CodeModule; pane = ((dynamic)code).CodePane;
                    ((dynamic)pane).Show(); window = ((dynamic)pane).Window; ((dynamic)window).SetFocus();
                    var before = Read(host, new { Command = "debug_state", Project = project });
                    Assert.AreEqual(other, Convert.ToString(before["ActiveModule"]), "The target must actually start inactive.");
                    var result = Read(host, new
                    {
                        Command = "run_sub", Project = project, Module = module, Procedure = "ColdMarker",
                        ExpectedMode = 2, ExpectedSha256 = source["Sha256"]
                    });
                    Assert.AreEqual(true, result["Executed"]);
                    Assert.AreEqual(1.0, Convert.ToDouble(host.ReadCell("A2")));
                    Assert.AreEqual(source["Sha256"], Read(host, new { Command = "read_module", Project = project, Module = module })["Sha256"]);
                    Write(host, "cold-run.json", new { InitialActiveModule = other, NativeRunEntries = 1,
                        SeparateTargetNavigation = false, SyntheticMacroExecuted = true, Marker = host.ReadCell("A2"), Result = result });
                }
                finally { Release(window); Release(pane); Release(code); Release(component); Release(components); Release(nativeProject); }
            });
        }

        /// <summary>Uses a real workbook event to accept or cancel the single product Save, then verifies the disk boundary.</summary>
        [STATestMethod, DataRow(true), DataRow(false)]
        public void NativeBeforeSaveReportsCancellationOrVerifiedPersistenceWithoutRetry(bool cancel)
        {
            RequireOptIn();
            string path = null, savedHash = null;
            ExcelVbeFixture.Run(host =>
            {
                string project = Project(host);
                path = host.File("BeforeSaveAudit.xlsm");
                var properties = Read(host, new { Command = "project_properties", Project = project });
                Read(host, new { Command = "save_host_document_as", Project = project, Path = path, ExpectedProjectVersion = properties["Version"] });
                string baseline = FileHash(path);
                object workbook = UiInvoke.Field<object>(host, "workbook"), application = UiInvoke.Field<object>(host, "application");
                string module = Convert.ToString(((dynamic)workbook).CodeName);
                var source = Read(host, new { Command = "read_module", Project = path, Module = module });
                var component = Read(host, new { Command = "component_properties", Project = path, Module = module });
                string handler = "Option Explicit\r\nPrivate Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)\r\n" +
                    "    Me.Worksheets(1).Range(\"A1\").Value2 = Me.Worksheets(1).Range(\"A1\").Value2 + 1\r\n" +
                    "    Cancel = " + (cancel ? "True" : "False") + "\r\nEnd Sub";
                Read(host, new { Command = "replace_lines", Project = path, Module = module,
                    ExpectedSha256 = source["Sha256"], StartLine = 1, Count = component["CodeLines"], Text = handler });
                string pending = Convert.ToString(Read(host, new { Command = "read_module", Project = path, Module = module })["Sha256"]);
                ((dynamic)application).EnableEvents = true;
                Assert.IsFalse(Convert.ToBoolean(((dynamic)workbook).Saved));
                properties = Read(host, new { Command = "project_properties", Project = path });
                var reply = host.Command(new { Command = "save_host_document", Project = path, ExpectedHostPath = path, ExpectedProjectVersion = properties["Version"] });
                RequireTerminal(host, reply);
                Assert.AreEqual(!cancel, reply["Ok"]);
                Assert.AreEqual(1.0, Convert.ToDouble(host.ReadCell("A1")), "The real BeforeSave handler must execute exactly once.");
                Assert.AreEqual(!cancel, Convert.ToBoolean(((dynamic)workbook).Saved));
                Assert.AreEqual(pending, Read(host, new { Command = "read_module", Project = path, Module = module })["Sha256"]);
                savedHash = FileHash(path);
                if (cancel)
                {
                    StringAssert.Contains(Convert.ToString(reply["Error"]), "BeforeSave");
                    Assert.AreEqual(baseline, savedHash, "Cancellation must leave the exact disk baseline untouched.");
                }
                else Assert.AreNotEqual(baseline, savedHash, "Successful product Save must persist the added event code and cell marker.");
                ((dynamic)application).EnableEvents = false;
                Write(host, "before-save.json", new { CancelRequested = cancel, NativeSaveEntries = 1,
                    SyntheticBeforeSaveExecuted = true, EventCount = host.ReadCell("A1"), DiskBaseline = baseline,
                    DiskAfter = savedHash, Reply = reply, HelperSaveInvoked = false });
            });
            if (!cancel)
                ExcelVbeFixture.Run(host =>
                {
                    host.OpenOwnedReadOnlyWorkbook(path);
                    Assert.AreEqual(1.0, Convert.ToDouble(host.ReadCell("A1")));
                    Assert.AreEqual(savedHash, FileHash(path));
                    Write(host, "before-save-reopened.json", new { Path = path, EventCount = host.ReadCell("A1"),
                        MacrosDisabledForReopen = true, HelperSaveInvoked = false });
                });
        }

        private static void RequireOptIn()
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_EXCEL_AUDIT_COMPLETION") != "1")
                Assert.Inconclusive("Explicit disposable Excel audit completion opt-in is required.");
        }

        private static string Project(ExcelVbeFixture host)
            => Convert.ToString(VbeBridgeClient.Object(((object[])host.Command("list_projects")["Data"]).Single())["Name"]);

        private static void CreateModule(ExcelVbeFixture host, string project, string module, string text)
        {
            var created = Read(host, new { Command = "create_module", Project = project, Module = module, ExpectedMode = 2 });
            var source = Read(host, new { Command = "read_module", Project = project, Module = module });
            Read(host, new { Command = "replace_lines", Project = project, Module = module,
                ExpectedSha256 = source["Sha256"], StartLine = 1, Count = created["Lines"], Text = text });
        }

        private static void RequireTerminal(ExcelVbeFixture host, IDictionary<string, object> reply)
        {
            if (reply == null || !reply.ContainsKey("Ok"))
            {
                host.PreserveForDiagnosticRecovery = true;
                Assert.Fail("Native outcome unknown; preserve the exact fixture and do not replay.");
            }
        }

        private static IDictionary<string, object> Read(ExcelVbeFixture host, object request)
        {
            var reply = host.Command(request); RequireTerminal(host, reply);
            Assert.AreEqual(true, reply["Ok"], Convert.ToString(reply["Error"]));
            return VbeBridgeClient.Object(reply["Data"]);
        }

        private void Write(ExcelVbeFixture host, string name, object value)
        {
            string path = host.File(name);
            File.WriteAllText(path, new JavaScriptSerializer().Serialize(value), new UTF8Encoding(false));
            TestContext.AddResultFile(path);
        }

        private static string FileHash(string path)
        {
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (var hash = SHA256.Create())
                return BitConverter.ToString(hash.ComputeHash(stream)).Replace("-", "");
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
        }
    }
}
