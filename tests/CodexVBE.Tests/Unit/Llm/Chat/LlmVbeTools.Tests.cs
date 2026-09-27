namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Threading;
    using System.Web.Script.Serialization;
    using System.Threading.Tasks;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class LlmVbeAsyncValidationTests
    {
        [TestMethod]
        public async Task NativeDebuggerToolsRejectMalformedArgumentsBeforeUiAccess()
        {
            var settings = new LlmSettings
            {
                VbeEditApproval = "Automatic"
            };
            var tools = new LlmVbeTools(null, null, settings);
            await Failure(tools, "debug_item", "{}", "Pane, Action and PathSegments");
            await Failure(tools, "debug_item", "{\"Pane\":\"locals\",\"Action\":\"expand\",\"PathSegments\":[\"x\"],\"Extra\":1}", "Unexpected debug_item argument");
            await Failure(tools, "immediate_execute", "{}", "Project, ExpectedMode and Text");
            await Failure(tools, "immediate_execute", "{\"Project\":\"P\",\"ExpectedMode\":0,\"Text\":\"Debug.Print 1\"}", "Project, ExpectedMode and Text");
            await Failure(tools, "respond_debug_dialog", "{}", "Exact Diagnostic and Button");
            await Failure(tools, "debug_dialog", "{\"Unexpected\":1}", "debug_dialog has no arguments");
            await Failure(tools, "compile_project", "{}", "Project and ExpectedMode=2");
            await Failure(tools, "compile_project", "{\"Project\":\"P\",\"ExpectedMode\":1}", "Project and ExpectedMode=2");
        }

        [TestMethod]
        public async Task NativeEvaluationAndMutationRequireAutomaticPolicy()
        {
            var settings = new LlmSettings
            {
                VbeEditApproval = "ReadOnly"
            };
            var tools = new LlmVbeTools(null, null, settings);
            await Failure(tools, "quick_watch", "{}", "Automatic VBE edit policy");
            await Failure(tools, "edit_watch", "{}", "Automatic VBE edit policy");
            await Failure(tools, "immediate_execute", "{}", "Automatic VBE edit policy");
            await Failure(tools, "respond_debug_dialog", "{}", "Automatic VBE edit policy");
        }

        [TestMethod]
        public async Task ScopeAndProjectBindingApplyToAsyncInvocationBeforeHostAccess()
        {
            var tools = new LlmVbeTools(null, null, new LlmSettings { VbeEditApproval = "Automatic" })
            {
                BoundProject = "WorkbookA"
            };
            await Failure(tools, "compile_project", "{\"Project\":\"WorkbookB\",\"ExpectedMode\":2}", "autre projet");
            tools.ValidateScope = () =>
            {
                throw new InvalidOperationException("stale conversation scope");
            };
            await Failure(tools, "debug_item", "{}", "stale conversation scope");
        }

        [TestMethod]
        public async Task AsyncDispatchReturnsReadOnlyStatusFromInMemorySession()
        {
            var tools = new LlmVbeTools(new VbeSession(new VbeSessionTests.FakeVbe()), null, new LlmSettings());
            var response = Json.Deserialize<Response>(await tools.InvokeAsync("status", "{}"));
            Assert.IsTrue(response.Ok);
            Assert.IsNotNull(response.Data);
        }

        [TestMethod]
        public async Task AsyncNativePreflightRejectsWrongShapesWithoutOpeningDialogs()
        {
            var tools = new LlmVbeTools(null, null, new LlmSettings { VbeEditApproval = "Automatic" });
            await Failure(tools, "debug_windows", "[]", "Tool arguments must be an object");
            await Failure(tools, "debug_windows", "{\"IncludeCallStack\":1}", "must be a boolean");
            await Failure(tools, "debug_windows", "{\"Unexpected\":true}", "Unexpected argument");
            await Failure(tools, "quick_watch", "[]", "Tool arguments must be an object");
            await Failure(tools, "edit_watch", "[]", "Tool arguments must be an object");
            await Failure(tools, "remove_watch", "[]", "Tool arguments must be an object");
        }

        [TestMethod]
        public async Task CompileRequiresUiContextBeforeNativeDialogInspection()
        {
            var tools = new LlmVbeTools(null, null, new LlmSettings());
            var prior = SynchronizationContext.Current;
            try
            {
                SynchronizationContext.SetSynchronizationContext(null);
                await Failure(tools, "compile_project", "{\"Project\":\"P\",\"ExpectedMode\":2}", "VBE UI context is unavailable");
            }
            finally
            {
                SynchronizationContext.SetSynchronizationContext(prior);
            }
        }

        [TestMethod]
        public async Task ImmediateExecuteRejectsModeChangedBeforeNativeExecution()
        {
            var host = new VbeSessionTests.FakeVbe();
            host.VBProjects.Add(new VbeSessionTests.FakeProject { Name = "P", FileName = @"C:\Temp\P.xlsm", Mode = 2 });
            var tools = new LlmVbeTools(new VbeSession(host), null, new LlmSettings { VbeEditApproval = "Automatic" });
            await Failure(tools, "immediate_execute", "{\"Project\":\"P\",\"ExpectedMode\":1,\"Text\":\"Debug.Print 1\"}", "Project mode changed");
        }
    }
}

namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Web.Script.Serialization;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class LlmVbeToolContractTests
    {
        [TestMethod]
        public void PublishedToolSchemasHaveUniqueNamesAndRequiredFieldsExist()
        {
            var names = new HashSet<string>(StringComparer.Ordinal);
            foreach (object definition in LlmVbeTools.Definitions)
            {
                var outer = Dict(Json.DeserializeObject(Json.Serialize(definition)));
                Assert.AreEqual("function", outer["type"]);
                var function = Dict(outer["function"]);
                var name = (string)function["name"];
                Assert.IsTrue(names.Add(name), "Duplicate LLM tool: " + name);
                var parameters = Dict(function["parameters"]);
                Assert.AreEqual(false, parameters["additionalProperties"]);
                var properties = Dict(parameters["properties"]);
                foreach (object required in (object[])parameters["required"])
                    Assert.IsTrue(properties.ContainsKey((string)required), name + ": " + required);
            }

            Assert.IsTrue(names.Contains("create_class"));
            Assert.IsTrue(names.Contains("form_tree"));
            Assert.IsTrue(names.Contains("debug_state"));
            Assert.IsTrue(names.Contains("list_reference_types"));
        }

        [TestMethod]
        public void InvocationRejectsMalformedAndUnexpectedArgumentsBeforeHostAccess()
        {
            var tools = new LlmVbeTools(null, null, new LlmSettings());
            IsFailure(tools.Invoke("missing_tool", "{}"), "Unknown tool");
            IsFailure(tools.Invoke("list_modules", "{}"), "Project is required");
            IsFailure(tools.Invoke("list_modules", "[]"), "object");
            IsFailure(tools.Invoke("list_modules", "{\"Project\":\"P\",\"Extra\":1}"), "Unexpected argument");
            IsFailure(tools.Invoke("list_modules", "{\"Project\":12}"), "must be a string");
            IsFailure(tools.Invoke("git_status", "{}"), "InvokeAsync");
            IsFailure(tools.Invoke("form_tree", "{\"Project\":\"P\"}"), "Form is required");
        }

        [TestMethod]
        public void EditingModesAndPoliciesRejectChangesBeforeHostAccess()
        {
            var settings = new LlmSettings
            {
                VbeEditApproval = "Automatic"
            };
            var tools = new LlmVbeTools(null, null, settings)
            {
                Mode = ChatMode.Plan
            };
            IsFailure(tools.Invoke("create_module", "{}"), "create_module");
            tools.Mode = ChatMode.Agent;
            settings.VbeEditApproval = "ReadOnly";
            IsFailure(tools.Invoke("create_module", "{\"Project\":\"P\",\"Module\":\"M\",\"ExpectedMode\":2}"), "VBE");
            settings.VbeEditApproval = "invalid";
            IsFailure(tools.Invoke("create_module", "{\"Project\":\"P\",\"Module\":\"M\",\"ExpectedMode\":2}"), "VBE");
            tools.ValidateScope = () =>
            {
                throw new InvalidOperationException("scope changed");
            };
            IsFailure(tools.Invoke("list_projects", "{}"), "scope changed");
        }

        [TestMethod]
        public void BoundProjectAndFilePathAreCheckedBeforeHostAccess()
        {
            var tools = new LlmVbeTools(null, null, new LlmSettings())
            {
                BoundProject = "WorkbookA"
            };
            IsFailure(tools.Invoke("create_module", "{\"Project\":\"WorkbookB\",\"Module\":\"M\",\"ExpectedMode\":2}"), "autre projet");
            IsFailure(tools.Invoke("inspect_code_file", "{\"Path\":\"C:\\\\Temp\\\\code.bas\"}"), "explicitement");
            IsFailure(tools.Invoke("set_form_node_property", "{\"Project\":\"WorkbookA\",\"Form\":\"F\",\"ControlPath\":\"X\",\"ExpectedTreeVersion\":\"v\",\"Property\":\"Caption\",\"Value\":[]}"), "Value must be");
        }

        [TestMethod]
        public void PlanAndReadOnlyModesCanDiscoverLiveProjectsThroughTheSameToolProtocol()
        {
            var host = new VbeSessionTests.FakeVbe();
            host.VBProjects.Add(new VbeSessionTests.FakeProject { Name = "ProjetÉté", FileName = @"C:\Temp\ProjetÉté.xlsm", Mode = 2 });
            var tools = new LlmVbeTools(new VbeSession(host), null, new LlmSettings { VbeEditApproval = "ReadOnly" })
            {
                Mode = ChatMode.Plan,
                BoundProject = "ProjetÉté"
            };
            var response = Json.Deserialize<Response>(tools.Invoke("list_projects", "{}"));
            Assert.IsTrue(response.Ok);
            var projects = (object[])response.Data;
            Assert.AreEqual(1, projects.Length);
            Assert.AreEqual("ProjetÉté", (string)Dict(projects[0])["Name"]);
            var status = Json.Deserialize<Response>(tools.Invoke("status", "{}"));
            Assert.IsTrue(status.Ok);
        }

        [TestMethod]
        public void FileReadRequiresTheExactUserProvidedAbsolutePathBeforeShowingApproval()
        {
            var tools = new LlmVbeTools(null, null, new LlmSettings());
            string supplied = @"C:\Temp\CodexVBE-user-file-does-not-exist.txt";
            IsFailure(tools.Invoke("read_user_file", Json.Serialize(new { Path = supplied })), "n'a pas fourni");
            tools.NoteUserRequest("Please inspect " + supplied);
            IsFailure(tools.Invoke("read_user_file", Json.Serialize(new { Path = @"C:\Temp\other.txt" })), "n'a pas fourni");
            IsFailure(tools.Invoke("read_user_file", Json.Serialize(new { Path = supplied })), "introuvable");
        }

        [TestMethod]
        public void AutomaticCodeEditPublishesVerifiedChangeAndCanRestoreIt()
        {
            var module = new VbeSessionTests.FakeModule("Alpha\r\nBeta");
            var project = new VbeSessionTests.FakeProject
            {
                Name = "VBAProject",
                FileName = @"C:\Temp\Host.xlsm",
                Mode = 2
            };
            project.VBComponents.Items.Add(new VbeSessionTests.FakeComponent { Name = "Module1", Type = 1, CodeModule = module });
            var host = new VbeSessionTests.FakeVbe();
            host.VBProjects.Add(project);
            var session = new VbeSession(host);
            dynamic before = session.Execute(new Request { Command = "read_module", Project = project.Name, Module = "Module1" }).Data;
            var tools = new LlmVbeTools(session, null, new LlmSettings { VbeEditApproval = "Automatic" });
            CodeChange observed = null;
            tools.CodeEdited += change => observed = change;
            var result = Json.Deserialize<Response>(tools.Invoke("replace_lines", Json.Serialize(new { Project = project.Name, Module = "Module1", ExpectedSha256 = (string)before.Sha256, StartLine = 2, Count = 1, Text = "Gamma" })));
            Assert.IsTrue(result.Ok, result.Error);
            Assert.AreEqual("Alpha\r\nGamma", module.Code);
            Assert.IsNotNull(observed);
            Assert.AreEqual("Alpha\r\nBeta", observed.Before);
            Assert.AreEqual("Alpha\r\nGamma", observed.After);
            Assert.IsTrue(tools.RestoreCodeChange(observed).Ok);
            Assert.AreEqual("Alpha\r\nBeta", module.Code);
            Assert.IsTrue(observed.Restored);
            Assert.IsFalse(tools.RestoreCodeChange(observed).Ok);
        }

        [TestMethod]
        public void MultiModuleUndoPreflightsEveryConflictBeforeWritingAnyModule()
        {
            var project = new VbeSessionTests.FakeProject
            {
                Name = "VBAProject",
                FileName = @"C:\Temp\Host.xlsm",
                Mode = 2
            };
            var first = new VbeSessionTests.FakeComponent
            {
                Name = "First",
                Type = 1,
                CodeModule = new VbeSessionTests.FakeModule("A2")
            };
            var second = new VbeSessionTests.FakeComponent
            {
                Name = "Second",
                Type = 1,
                CodeModule = new VbeSessionTests.FakeModule("unexpected")
            };
            project.VBComponents.Items.Add(first);
            project.VBComponents.Items.Add(second);
            var host = new VbeSessionTests.FakeVbe();
            host.VBProjects.Add(project);
            var tools = new LlmVbeTools(new VbeSession(host), null, new LlmSettings { VbeEditApproval = "Automatic" });
            var stale = new CodeChange(project.Name, second.Name, "B", "", "B2", "", 1);
            var valid = new CodeChange(project.Name, first.Name, "A", "", "A2", "", 1);
            var rejected = tools.RestoreChanges(new[] { stale, valid }, null);
            Assert.IsFalse(rejected.Ok);
            Assert.AreEqual("A2", first.CodeModule.Code);
            Assert.AreEqual("unexpected", second.CodeModule.Code);
            Assert.IsFalse(valid.Restored);
            second.CodeModule = new VbeSessionTests.FakeModule("B2");
            var restored = tools.RestoreChanges(new[] { stale, valid }, null);
            Assert.IsTrue(restored.Ok, restored.Error);
            Assert.AreEqual("A", first.CodeModule.Code);
            Assert.AreEqual("B", second.CodeModule.Code);
            Assert.IsTrue(valid.Restored);
            Assert.IsTrue(stale.Restored);
        }
    }
}
