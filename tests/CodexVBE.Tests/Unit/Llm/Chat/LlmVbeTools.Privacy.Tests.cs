using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class LlmProjectPrivacyTests
    {
        private readonly JavaScriptSerializer json = new JavaScriptSerializer();
        private LlmVbeTools Bound() => new LlmVbeTools(null, null, new LlmSettings { VbeEditApproval = "Automatic" }) { BoundProject = "A" };

        [TestMethod]
        public async Task OtherProjectReadIsDeniedBeforeDispatchForSyncAndAsyncTools()
        {
            var tools = Bound();
            int dispatched = 0;
            tools.Execute = request => { if (request.Command == "list_projects") return Response.Success(new object[0]); dispatched++; return Response.Success(new { Code = "PRIVATE-B" }); };
            var denied = json.Deserialize<Response>(tools.Invoke("read_module", "{\"Project\":\"B\",\"Module\":\"M\"}"));
            Assert.IsFalse(denied.Ok);
            var modern = json.Deserialize<Response>(await tools.InvokeAsync("monaco_read", "{\"Project\":\"B\",\"Module\":\"M\"}"));
            Assert.IsFalse(modern.Ok);
            Assert.AreEqual(0, dispatched);
            Assert.IsFalse(json.Serialize(denied).Contains("PRIVATE-B"));
        }

        [TestMethod]
        public void ExplicitProjectGrantAllowsReadsButNeverWrites()
        {
            var tools = Bound();
            int dispatched = 0;
            tools.Execute = request => { dispatched++; return Response.Success(new { Code = "AUTHORIZED-B" }); };
            tools.SetReadAccess(new[] { "B" }, false);
            var read = json.Deserialize<Response>(tools.Invoke("read_module", "{\"Project\":\"B\",\"Module\":\"M\"}"));
            Assert.IsTrue(read.Ok);
            StringAssert.Contains(json.Serialize(read.Data), "AUTHORIZED-B");
            Assert.IsFalse(json.Deserialize<Response>(tools.Invoke("create_module", "{\"Project\":\"B\",\"Module\":\"M\",\"ExpectedMode\":2}")).Ok);
            Assert.AreEqual(1, dispatched);
            tools.SetReadAccess(new string[0], false);
            Assert.ThrowsException<InvalidOperationException>(() => tools.RequireProjectRead("B"));
        }

        [TestMethod]
        public void InventoriesAndLiveContextNeverIncludeUnapprovedProjects()
        {
            var tools = Bound();
            tools.Execute = request => Response.Success(new[] {
                new { Name = "A", FileName = "C:\\A.xlsm", Mode = 2 },
                new { Name = "PRIVATE-B", FileName = "C:\\PRIVATE-B.xlsm", Mode = 2 } });
            var list = tools.Invoke("list_projects", "{}");
            StringAssert.Contains(list, "A.xlsm");
            Assert.IsFalse(list.Contains("PRIVATE-B"));
            Assert.IsFalse(tools.Invoke("status", "{}").Contains("PRIVATE-B"));
            Assert.IsFalse(tools.LiveContextJson().Contains("PRIVATE-B"));
        }

        [TestMethod]
        public void DuplicateNamesCannotExpandInventoryAuthorization()
        {
            var tools = Bound();
            tools.Execute = request => Response.Success(new[] {
                new { Name = "A", FileName = "C:\\A.xlsm", Mode = 2 },
                new { Name = "A", FileName = "C:\\SECRET.xlsm", Mode = 2 } });
            Assert.AreEqual(0, ((object[])json.Deserialize<Response>(tools.Invoke("list_projects", "{}")).Data).Length);
            tools.BoundProject = "C:\\A.xlsm";
            Assert.IsFalse(tools.Invoke("list_projects", "{}").Contains("SECRET"));
        }

        [TestMethod]
        public async Task GlobalReadToolsNeedTheirOwnGrantEvenWithProjectGrants()
        {
            var tools = Bound();
            tools.SetReadAccess(new[] { "B" }, false);
            foreach (string name in new[] { "debug_windows", "debug_dialog", "read_runtime_forms", "read_navigation_surface", "project_collection_state", "code_panes", "vbe_windows", "vbe_environment", "read_code_clipboard", "read_object_browser" })
            {
                var response = json.Deserialize<Response>(await tools.InvokeAsync(name, "{}"));
                Assert.IsFalse(response.Ok, name);
                StringAssert.Contains(response.Error, UiText.Get("This tool uses shared VBE context. Authorize shared context in Project access before using it."), name);
            }
            tools.Execute = request => Response.Success(new { Clipboard = "explicit shared grant" });
            tools.SetReadAccess(new string[0], true);
            Assert.IsTrue(json.Deserialize<Response>(tools.Invoke("read_code_clipboard", "{}")).Ok);
            Assert.ThrowsException<InvalidOperationException>(() => tools.RequireProjectRead("B"));
        }

        [TestMethod]
        public void DebugStateRemovesActiveContextFromAnotherProject()
        {
            var tools = Bound();
            tools.Execute = request => Response.Success(new { Project = "A", Mode = 1, SelectedProject = "B", SelectedProjectPath = "C:\\B.xlsm", ActiveModule = "PRIVATE-B", Selection = new { StartLine = 21 } });
            string result = tools.Invoke("debug_state", "{\"Project\":\"A\"}");
            Assert.IsTrue(json.Deserialize<Response>(result).Ok);
            Assert.IsFalse(result.Contains("PRIVATE-B"));
            Assert.IsFalse(result.Contains("B.xlsm"));
        }

        [TestMethod]
        public void ProjectScopedDebugStatePreservesItsOwnLocation()
        {
            var tools = Bound();
            tools.Execute = request => Response.Success(new { Project = "A", Mode = 1, SelectedProject = "A", SelectedProjectPath = "C:\\A.xlsm", ActiveModule = "MyModule", Selection = new { StartLine = 21 } });
            StringAssert.Contains(tools.Invoke("debug_state", "{\"Project\":\"A\"}"), "MyModule");
        }

        [TestMethod]
        public void AttachmentPermissionChecksFollowExactSelectors()
        {
            var tools = Bound();
            tools.SetReadAccess(new[] { "C:\\B.xlsm" }, false);
            tools.RequireProjectRead("C:\\B.xlsm");
            Assert.ThrowsException<InvalidOperationException>(() => tools.RequireProjectRead("B"));
            Assert.ThrowsException<InvalidOperationException>(() => tools.RequireProjectRead(null));
        }

        [TestMethod]
        public void ReferenceNameAliasesResolveOnlyWhenUniqueAndAuthorizedByPath()
        {
            var tools = Bound();
            tools.BoundProject = "C:\\A.xlsm";
            tools.Execute = request => Response.Success(new[] {
                new { Name = "A", FileName = "C:\\A.xlsm", Mode = 2 },
                new { Name = "B", FileName = "C:\\B.xlsm", Mode = 2 } });
            tools.RequireProjectRead("A");
            Assert.ThrowsException<InvalidOperationException>(() => tools.RequireProjectRead("B"));
            tools.SetReadAccess(new[] { "C:\\B.xlsm" }, false);
            tools.RequireProjectRead("B");
            tools.Execute = request => Response.Success(new[] {
                new { Name = "B", FileName = "C:\\B.xlsm", Mode = 2 },
                new { Name = "B", FileName = "C:\\SECRET.xlsm", Mode = 2 } });
            Assert.ThrowsException<InvalidOperationException>(() => tools.RequireProjectRead("B"));
        }
    }
}
