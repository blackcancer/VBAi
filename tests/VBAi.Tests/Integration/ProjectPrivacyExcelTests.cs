using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Web.Script.Serialization;
using VBAi;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    [TestClass, TestCategory("MonacoExcel")]
    public sealed class ProjectPrivacyExcelTests
    {
        [STATestMethod]
        public void RealExcelProjectsStayPrivateUntilReadGrantAndNeverBecomeWritable()
        {
            ExcelScenarioLifetime.Run(RunScenario);
        }

        [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
        private static void RunScenario(ExcelScenarioLifetime lifetime)
        {
            if (Environment.GetEnvironmentVariable("VBAI_EDITOR_EXCEL_TEST") != "1") Assert.Inconclusive("Explicit disposable Excel opt-in required.");
            if (Process.GetProcessesByName("EXCEL").Length != 0) Assert.Inconclusive("Close existing Excel processes before this isolated test.");
            dynamic excel = null, first = null, second = null;
            const string projectA = "PrivacyFixtureA", projectB = "PrivacyFixtureSecretB";
            const string secret = "SECRET_PRIVACY_SENTINEL_B";
            var json = new JavaScriptSerializer();
            try
            {
                excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                lifetime.Capture((object)excel);
                excel.Visible = true; excel.DisplayAlerts = false; excel.EnableEvents = false;
                first = excel.Workbooks.Add(); second = excel.Workbooks.Add();
                first.VBProject.Name = projectA; second.VBProject.Name = projectB;
                dynamic moduleA = first.VBProject.VBComponents.Add(1);
                dynamic moduleB = second.VBProject.VBComponents.Add(1);
                moduleA.Name = "AllowedCode"; moduleB.Name = "SecretCode";
                moduleA.CodeModule.AddFromString("Option Explicit\r\nPublic Const Allowed = 1");
                moduleB.CodeModule.AddFromString("Option Explicit\r\nPublic Const SecretValue = \"" + secret + "\"");
                string sourceB = moduleB.CodeModule.Lines[1, moduleB.CodeModule.CountOfLines];
                string hostB = second.FullName;
                object vbe = excel.VBE;
                var tools = new LlmVbeTools(new VbeSession(vbe), null,
                    new LlmSettings { VbeEditApproval = "Automatic" }) { BoundProject = projectA };
                string readA = json.Serialize(new { Project = projectA, Module = "AllowedCode" });
                string readB = json.Serialize(new { Project = projectB, Module = "SecretCode" });
                Assert.IsTrue(json.Deserialize<Response>(tools.Invoke("read_module", readA)).Ok);
                string denied = tools.Invoke("read_module", readB);
                Assert.IsFalse(json.Deserialize<Response>(denied).Ok);
                Assert.IsFalse(denied.Contains(secret));
                Assert.IsFalse(json.Deserialize<Response>(tools.Invoke("code_panes", "{}")).Ok);
                foreach (string inventory in new[] { tools.Invoke("list_projects", "{}"), tools.Invoke("status", "{}"), tools.LiveContextJson() })
                {
                    Assert.IsFalse(inventory.Contains(projectB), "Unapproved project name escaped filtering.");
                    Assert.IsFalse(inventory.Contains(hostB), "Unapproved host identity escaped filtering.");
                    Assert.IsFalse(inventory.Contains(secret));
                }
                // Tree project selection and active code pane can identify different projects.
                moduleB.CodeModule.CodePane.Show();
                ((dynamic)vbe).ActiveVBProject = first.VBProject;
                string state = tools.Invoke("debug_state", json.Serialize(new { Project = projectA }));
                Assert.IsTrue(json.Deserialize<Response>(state).Ok);
                Assert.IsFalse(state.Contains("SecretCode"));
                Assert.IsFalse(state.Contains(projectB));
                tools.SetReadAccess(new[] { projectB }, false);
                string granted = tools.Invoke("read_module", readB);
                Assert.IsTrue(json.Deserialize<Response>(granted).Ok);
                StringAssert.Contains(granted, secret);
                int countBefore = second.VBProject.VBComponents.Count;
                Assert.IsFalse(json.Deserialize<Response>(tools.Invoke("create_module", json.Serialize(new {
                    Project = projectB, Module = "UnauthorizedWrite", ExpectedMode = 2 }))).Ok);
                Assert.AreEqual(countBefore, (int)second.VBProject.VBComponents.Count);
                Assert.AreEqual(sourceB, (string)moduleB.CodeModule.Lines[1, moduleB.CodeModule.CountOfLines]);
                Assert.IsFalse(json.Deserialize<Response>(tools.Invoke("code_panes", "{}")).Ok,
                    "A project read grant must not grant shared VBE context.");
                Assert.AreEqual("", (string)first.Path); Assert.AreEqual("", (string)second.Path);
                Assert.AreEqual(2, (int)first.VBProject.Mode); Assert.AreEqual(2, (int)second.VBProject.Mode);
            }
            finally
            {
                if ((object)first != null) { try { first.Close(false); } catch (COMException) { } Marshal.FinalReleaseComObject((object)first); }
                if ((object)second != null) { try { second.Close(false); } catch (COMException) { } Marshal.FinalReleaseComObject((object)second); }
                if ((object)excel != null) { try { if (lifetime.OwnsApplication) excel.Quit(); } catch (COMException) { } Marshal.FinalReleaseComObject((object)excel); }
            }
        }
    }
}
