using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Web.Script.Serialization;

namespace VBAi.Tests.Integration
{
    /// <summary>Product save followed by discard-close, normal process exit and independent disk readback.</summary>
    [TestClass, TestCategory("Excel")]
    public sealed class ExcelQualificationPersistenceTests
    {
        public TestContext TestContext { get; set; }

        [STATestMethod]
        public void ProductSavePreservesModuleClassAndFormInFreshOwnedProcessWithoutHelperSaving()
        {
            string path = null;
            var expected = new Dictionary<string, string>();
            int savedPid = 0;
            ExcelVbeFixture.Run(host =>
            {
                savedPid = host.ProcessId;
                string project = Project(host);
                path = host.File("Q006Saved.xlsm");
                var properties = Data(host.Command(new { Command = "project_properties", Project = project }));
                Data(host.Command(new { Command = "save_host_document_as", Project = project, Path = path, ExpectedProjectVersion = properties["Version"] }));
                foreach (var item in new[] { Tuple.Create("Q006Module", 1), Tuple.Create("Q006Class", 2), Tuple.Create("Q006Form", 3) })
                {
                    if (item.Item2 == 3) Data(host.Command(new { Command = "create_form", Project = path, Form = item.Item1 }));
                    else Data(host.Command(new { Command = item.Item2 == 2 ? "create_class" : "create_module", Project = path, Module = item.Item1, ExpectedMode = 2 }));
                    var before = Data(host.Command(new { Command = "read_module", Project = path, Module = item.Item1 }));
                    var component = Data(host.Command(new { Command = "component_properties", Project = path, Module = item.Item1 }));
                    Data(host.Command(new
                    {
                        Command = "replace_lines",
                        Project = path,
                        Module = item.Item1,
                        ExpectedSha256 = before["Sha256"],
                        StartLine = 1,
                        Count = component["CodeLines"],
                        Text = "Option Explicit\r\nPublic Sub Q006Marker()\r\n    Debug.Print \"synthetic Q006\"\r\nEnd Sub"
                    }));
                    expected[item.Item1] = (string)Data(host.Command(new { Command = "read_module", Project = path, Module = item.Item1 }))["Code"];
                }
                var form = Data(host.Command(new { Command = "form_state", Project = path, Form = "Q006Form" }));
                Data(host.Command(new
                {
                    Command = "add_form_control",
                    Project = path,
                    Form = "Q006Form",
                    ExpectedFormVersion = form["Version"],
                    Control = "Q006Label",
                    ControlType = "Forms.Label.1",
                    Caption = "Q006 persisted label",
                    Left = 12,
                    Top = 12,
                    Width = 140,
                    Height = 24
                }));
                var beforeSave = Data(host.Command(new { Command = "project_persistence_status", Project = path }));
                Assert.AreEqual(false, beforeSave["HostSaved"], "Pending module/class/form edits must make the host dirty.");
                properties = Data(host.Command(new { Command = "project_properties", Project = path }));
                var saved = Data(host.Command(new { Command = "save_host_document", Project = path, ExpectedHostPath = path, ExpectedProjectVersion = properties["Version"] }));
                Assert.AreEqual(true, saved["SaveInvoked"]);
                Assert.AreEqual(true, Data(host.Command(new { Command = "project_persistence_status", Project = path }))["HostSaved"]);
                Assert.IsTrue(File.Exists(path));
                // Fixture Close(false)/Quit must not perform a helper Save.
            });
            ExcelVbeFixture.Run(reopened =>
            {
                Assert.AreNotEqual(savedPid, reopened.ProcessId);
                object document = null, project = null, components = null;
                try
                {
                    reopened.OpenOwnedReadOnlyWorkbook(path);
                    document = VBAi.Tests.Infrastructure.UiInvoke.Field<object>(reopened, "workbook");
                    project = ((dynamic)document).VBProject;
                    components = ((dynamic)project).VBComponents;
                    foreach (var pair in expected)
                    {
                        object component = null, code = null;
                        try
                        {
                            component = ((dynamic)components).Item(pair.Key);
                            code = ((dynamic)component).CodeModule;
                            Assert.AreEqual(pair.Value, Convert.ToString(((dynamic)code).Lines[1, ((dynamic)code).CountOfLines]), pair.Key);
                        }
                        finally { Release(code); Release(component); }
                    }
                    object form = null, designer = null, controls = null, label = null;
                    try
                    {
                        form = ((dynamic)components).Item("Q006Form");
                        Assert.AreEqual(3, Convert.ToInt32(((dynamic)form).Type));
                        designer = ((dynamic)form).Designer; controls = ((dynamic)designer).Controls;
                        Assert.AreEqual(1, Convert.ToInt32(((dynamic)controls).Count));
                        label = ((dynamic)controls).Item("Q006Label");
                        Assert.AreEqual("Q006 persisted label", Convert.ToString(((dynamic)label).Caption));
                    }
                    finally { Release(label); Release(controls); Release(designer); Release(form); }
                    string reportDirectory = string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("VBAi_EXCEL_RESULTS"))
                        ? TestContext.TestResultsDirectory : reopened.Root;
                    string report = Path.Combine(reportDirectory, "q006-persistence.json");
                    File.WriteAllText(report,
                        new JavaScriptSerializer().Serialize(new
                        {
                            SavedPid = savedPid,
                            ReopenedPid = reopened.ProcessId,
                            Path = path,
                            ModuleSources = expected,
                            Form = "Q006Form",
                            Label = "Q006 persisted label",
                            HelperSaveInvoked = false,
                            MacrosDisabledOnReopen = true,
                            AssemblyMvid = typeof(VbeSession).Module.ModuleVersionId
                        }));
                    TestContext.AddResultFile(report);
                }
                finally
                {
                    Release(components); Release(project);
                    // The document belongs to the fixture, which performs its one Close/Quit.
                }
            });
        }

        private static string Project(ExcelVbeFixture host) => Convert.ToString(VbeBridgeClient.Object(((object[])host.Command("list_projects")["Data"]).Single())["Name"]);
        private static IDictionary<string, object> Data(IDictionary<string, object> response)
        {
            Assert.IsNotNull(response);
            Assert.AreEqual(true, response["Ok"], new JavaScriptSerializer().Serialize(response));
            return VbeBridgeClient.Object(response["Data"]);
        }
        private static void Release(object value) { if (value != null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value); }
    }
}
