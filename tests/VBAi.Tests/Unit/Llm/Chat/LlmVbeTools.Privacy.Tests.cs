using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace VBAi.Tests.Unit
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
        public async Task ReadImmediateRequiresBoundProjectAndSharedContextGrant()
        {
            var tools = Bound();
            int nativeCalls = 0;
            tools.Execute = request =>
            {
                if (request.Command == "debug_state")
                {
                    var state = new System.Dynamic.ExpandoObject();
                    ((IDictionary<string, object>)state)["Mode"] = 1;
                    return Response.Success(state);
                }
                throw new InvalidOperationException("Unexpected synchronous command.");
            };
            tools.ReadImmediateNative = request =>
            {
                nativeCalls++;
                return Task.FromResult<object>(new { Text = "private VBE-wide output" });
            };
            const string own = "{\"Project\":\"A\",\"ExpectedMode\":1}";
            const string other = "{\"Project\":\"B\",\"ExpectedMode\":1}";
            Assert.IsFalse(json.Deserialize<Response>(await tools.InvokeAsync("read_immediate", own)).Ok);
            Assert.AreEqual(0, nativeCalls);
            tools.SetReadAccess(new[] { "B" }, true);
            Assert.IsFalse(json.Deserialize<Response>(await tools.InvokeAsync("read_immediate", other)).Ok);
            Assert.AreEqual(0, nativeCalls);
            var granted = json.Deserialize<Response>(await tools.InvokeAsync("read_immediate", own));
            Assert.IsTrue(granted.Ok, granted.Error);
            Assert.AreEqual(1, nativeCalls);
        }

        [TestMethod]
        public async Task LocalScalarInspectionRequiresBoundProjectAndSharedContextGrant()
        {
            var tools = Bound();
            int nativeCalls = 0;
            tools.Execute = request =>
            {
                var state = new System.Dynamic.ExpandoObject();
                ((IDictionary<string, object>)state)["Mode"] = 1;
                return Response.Success(state);
            };
            tools.InspectLocalScalarsNative = request =>
            {
                nativeCalls++;
                return Task.FromResult<object>(new { Partial = true });
            };
            const string own = "{\"Project\":\"A\",\"Module\":\"M\",\"Procedure\":\"Run\",\"ExpectedSha256\":\"aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa\",\"ExpectedMode\":1}";
            const string other = "{\"Project\":\"B\",\"Module\":\"M\",\"Procedure\":\"Run\",\"ExpectedSha256\":\"aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa\",\"ExpectedMode\":1}";
            Assert.IsFalse(json.Deserialize<Response>(await tools.InvokeAsync("inspect_local_scalars", own)).Ok);
            tools.SetReadAccess(new[] { "B" }, true);
            Assert.IsFalse(json.Deserialize<Response>(await tools.InvokeAsync("inspect_local_scalars", other)).Ok);
            Assert.AreEqual(0, nativeCalls);
            var granted = json.Deserialize<Response>(await tools.InvokeAsync("inspect_local_scalars", own));
            Assert.IsTrue(granted.Ok, granted.Error);
            Assert.AreEqual(1, nativeCalls);
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
namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ProjectPrivacyBoundaryTests
    {
        private readonly JavaScriptSerializer json = new JavaScriptSerializer();
        private LlmVbeTools Tools(string bound = "A") => new LlmVbeTools(null, null, new LlmSettings()) { BoundProject = bound };
        [TestMethod]
        public void PrivacyGuardSeparatesProjectReadEditSharedAndIndependentCommands()
        {
            var tools = Tools();
            tools.Execute = r => Response.Success(new object[0]);
            tools.SetReadAccess(new[] { null, " ", "B" }, false);
            tools.RequireProjectRead("B");
            foreach (var arguments in new[] { "null", "[]", "{}", "{\"Project\":null}", "{\"Project\":4}", "{\"Project\":\" \"}" })
            {
                if (arguments == "{}" || arguments == "null" || arguments == "[]") AccessContract.Denied(tools, "GuardProjectPrivacy", "debug_windows", arguments);
                else AccessContract.Call(tools, "GuardProjectPrivacy", "read_module", arguments);
            }
            foreach (var command in new[] { "status", "git_status" }) AccessContract.Call(tools, "GuardProjectPrivacy", command, "{}");
            foreach (var command in new[] { "compile_project", "create_module" })
            {
                AccessContract.Call(tools, "GuardProjectPrivacy", command, "{\"Project\":\"A\"}");
                AccessContract.Denied(tools, "GuardProjectPrivacy", command, "{\"Project\":\"B\"}");
            }
            AccessContract.Denied(tools, "GuardProjectPrivacy", "debug_global", "{\"Project\":\"A\"}");
            tools.SetReadAccess(null, true); AccessContract.Call(tools, "GuardProjectPrivacy", "debug_global", "{\"Project\":\"A\"}");
            AccessContract.Call(tools, "GuardProjectPrivacy", "debug_windows", "{}");
            tools.BoundProject = null; tools.RequireProjectRead("Other"); AccessContract.Call(tools, "GuardProjectPrivacy", "debug_windows", "{}");
            Assert.IsFalse((bool)AccessContract.Call(Tools(), "IsAuthorizedAlias", new object[] { null }));
        }

        [TestMethod]
        public void AliasResolutionFailsClosedForMissingAmbiguousMalformedOrUnavailableInventories()
        {
            var tools = Tools(@"C:\A.xlsm");
            foreach (var data in new object[] { null, new { }, new object[] { new { Name = "B", FileName = @"C:\B.xlsm" }, new { Name = "B", FileName = @"C:\C.xlsm" } }, new object[] { new { Name = "A", FileName = @"C:\A.xlsm" }, new { Name = "B", FileName = @"C:\B.xlsm" } } })
            {
                tools.Execute = r => Response.Success(data);
                Assert.IsFalse((bool)AccessContract.Call(tools, "IsAuthorizedAlias", "B"));
            }
            tools.Execute = r => Response.Failure("unavailable"); Assert.IsFalse((bool)AccessContract.Call(tools, "IsAuthorizedAlias", "B"));
            tools.Execute = r => throw new InvalidOperationException("offline"); Assert.IsFalse((bool)AccessContract.Call(tools, "IsAuthorizedAlias", "B"));
            tools.BoundProject = "A"; tools.SetReadAccess(new[] { "B" }, false);
            tools.Execute = r => Response.Success(new[] { new { Name = "A", FileName = @"C:\A.xlsm" }, new { Name = "B", FileName = @"C:\B.xlsm" } });
            Assert.IsTrue((bool)AccessContract.Call(tools, "IsAuthorizedAlias", @"C:\B.xlsm"));
            tools.Execute = r => Response.Success(new object[] { new { FileName = @"C:\B.xlsm" } });
            Assert.IsFalse((bool)AccessContract.Call(tools, "IsAuthorizedAlias", "B"));
        }

        [TestMethod]
        public void ProjectFilteringHandlesMissingFieldsNameUniquenessPathsAndExactDebugContext()
        {
            var tools = Tools();
            foreach (var data in new object[] { null, new { }, new object[] { new { Name = "A", FileName = "" }, new { Name = "X" } }, new object[] { new { FileName = @"C:\A.xlsm" }, new { Name = "X", FileName = "" } } })
            {
                var result = (System.Array)AccessContract.Call(tools, "FilterProjects", data);
                Assert.IsNotNull(result);
            }
            tools.BoundProject = @"C:\A.xlsm";
            var visible = (System.Array)AccessContract.Call(tools, "FilterProjects", (object)new[] { new { Name = "A", FileName = @"C:\A.xlsm" }, new { Name = "A", FileName = @"C:\B.xlsm" } });
            Assert.AreEqual(1, visible.Length);
            var failure = Response.Failure("native error");
            Assert.AreSame(failure, AccessContract.Call(tools, "FilterProjectResponse", "debug_state", failure, "A"));
            var pass = Response.Success(new { Data = "safe" });
            Assert.AreSame(pass, AccessContract.Call(tools, "FilterProjectResponse", "read_module", pass, "A"));
            tools.BoundProject = null; Assert.AreSame(pass, AccessContract.Call(tools, "FilterProjectResponse", "debug_state", pass, "A"));
            Assert.AreSame(pass.Data, AccessContract.Call(tools, "FilterProjects", pass.Data));
            tools.BoundProject = "A";
            foreach (var data in new object[]{
                new{Project="A",SelectedProject="A",SelectedProjectPath="",ActiveModule="M"},
                new{Project="B",SelectedProject="A",SelectedProjectPath=@"C:\A.xlsm",ActiveModule="PRIVATE"},
                new{Project="A",SelectedProject="",SelectedProjectPath="",ActiveModule="PRIVATE"},
                new{Project="A",SelectedProject="A",SelectedProjectPath=@"C:\A.xlsm",ActiveModule="M"}})
            {
                var result = (Response)AccessContract.Call(tools, "FilterProjectResponse", "debug_state", Response.Success(data), "A");
                var fields = (IDictionary<string, object>)result.Data;
                Assert.IsFalse(Convert.ToString(fields["ActiveModule"]) == "PRIVATE");
            }
            tools.Execute = r => Response.Failure("missing inventory");
            var failed = json.Serialize(AccessContract.Call(tools, "ScopedLiveSnapshot")); StringAssert.Contains(failed, "missing inventory");
            tools.Execute = r => throw new InvalidOperationException("offline");
            StringAssert.Contains(json.Serialize(AccessContract.Call(tools, "ScopedLiveSnapshot")), "Project inventory unavailable");
        }
    }
}
