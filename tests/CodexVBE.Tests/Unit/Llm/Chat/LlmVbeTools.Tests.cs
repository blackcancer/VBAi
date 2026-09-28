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

namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Web.Script.Serialization;
    using System.Windows.Forms;
    using CodexVBE;
    using CodexVBE.Tests.Infrastructure;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class LlmVbeToolsBoundaryTests
    {
        [TestMethod]
        [STATestMethod]
        public void ContractMatrixChecksRequiredTypesWhitespaceOptionalFieldsAndHostDispatch()
        {
            Assert.ThrowsException<ArgumentNullException>(() => new LlmVbeTools(null, null, null));
            var tools = Create();
            tools.NoteUserRequest(null); tools.NoteUserRequest(" ");
            tools.NoteUserRequest(@"The supplied path is C:\Temp\fixture.bas");
            foreach (object definition in LlmVbeTools.Definitions)
            {
                var function = Dict(Dict(Json.DeserializeObject(Json.Serialize(definition)))["function"]);
                string name = (string)function["name"];
                if (name.StartsWith("git_") || name == "read_user_file" || name == "replace_lines") continue;
                var parameters = Dict(function["parameters"]);
                var required = (object[])parameters["required"];
                var fields = Dict(parameters["properties"]);
                var values = Arguments(name);
                Success(tools.Invoke(name, Json.Serialize(values)), name);
                foreach (string field in fields.Keys)
                {
                    object original = values[field];
                    values[field] = null;
                    Failed(tools.Invoke(name, Json.Serialize(values)), field);
                    values[field] = new object[] { new object() };
                    Failed(tools.Invoke(name, Json.Serialize(values)), name + ":" + field);
                    values[field] = original;
                }
                foreach (string field in required.Cast<string>())
                {
                    object original = values[field]; values.Remove(field);
                    Failed(tools.Invoke(name, Json.Serialize(values)), name + ":missing:" + field);
                    values[field] = original;
                    if (original is string && field != "Text" && field != "Caption" && field != "Value")
                    {
                        values[field] = " "; Failed(tools.Invoke(name, Json.Serialize(values)), field); values[field] = original;
                    }
                }
            }
            foreach (object value in new object[] { "", true, 1, 2147483648L, 1.25m, 1e50 })
            {
                var values = Arguments("set_form_node_property"); values["Value"] = value;
                Success(tools.Invoke("set_form_node_property", Json.Serialize(values)), "scalar " + value);
            }
            foreach (object value in new object[] { 1, 2147483648L, 1.25m, 1e50 })
            {
                var values = Arguments("add_form_control");
                // Every numeric field is checked through its published schema.
                foreach (string field in values.Keys.ToArray())
                    if (field == "Left" || field == "Top" || field == "Width" || field == "Height") values[field] = value;
                Success(tools.Invoke("add_form_control", Json.Serialize(values)), "numeric " + value);
            }
            tools.BoundProject = "P";
            Failed(tools.Invoke("create_module", "[]"), "bound array");
            Failed(tools.Invoke("create_module", "{}"), "bound missing project");
            Success(tools.Invoke("create_module", Json.Serialize(Arguments("create_module"))), "bound same project");
            tools.Mode = ChatMode.Plan;
            Failed(tools.Invoke("create_module", Json.Serialize(Arguments("create_module"))), "plan edit");
            Success(tools.Invoke("list_projects", "{}"), "plan read");
            tools.Mode = ChatMode.Agent;
            tools.Settings.VbeEditApproval = "AskEachTime";
            tools.ShowApproval = (dialog, owner) => DialogResult.No;
            Failed(tools.Invoke("create_module", Json.Serialize(Arguments("create_module"))), "approval refusal");
            tools.ShowApproval = (dialog, owner) => DialogResult.Yes;
            Success(tools.Invoke("create_module", Json.Serialize(Arguments("create_module"))), "approval yes");
        }

        [TestMethod]
        public void ItemsMatrixAcceptsBoundaryLengthsAndRejectsMultilineWrongTypesAndOverflow()
        {
            var tools = Create();
            string name = LlmVbeTools.Definitions.Select(d => Dict(Dict(Json.DeserializeObject(Json.Serialize(d)))["function"]))
                .Where(f => Dict(Dict(f["parameters"])["properties"]).ContainsKey("Items")).Select(f => (string)f["name"]).First();
            var values = Arguments(name);
            foreach (object items in new object[] { new string[0], new[] { "", new string('x', 256) }, Enumerable.Repeat("x", 64).ToArray() })
            { values["Items"] = items; Success(tools.Invoke(name, Json.Serialize(values)), "valid items"); }
            foreach (object items in new object[] { "x", Enumerable.Repeat("x", 65).ToArray(), new object[] { 2 }, new[] { new string('x',257) }, new[] { "line\nline" } })
            { values["Items"] = items; Failed(tools.Invoke(name, Json.Serialize(values)), "invalid items"); }
        }

        [TestMethod]
        public void FileReadMatrixRequiresLiteralAbsolutePathConfirmationAndTextSizeLimit()
        {
            using (var scope = new LlmBoundaryScope())
            {
                var tools = Create();
                string file = Path.Combine(scope.Root, "document.txt");
                File.WriteAllText(file, "étè", new System.Text.UTF8Encoding(false));
                tools.NoteUserRequest("Read " + file);
                tools.ConfirmFile = (owner,text,title) => { StringAssert.Contains(text,file); return DialogResult.No; };
                Failed(tools.Invoke("read_user_file", Json.Serialize(new { Path = file })), "declined");
                tools.ConfirmFile = (owner,text,title) => DialogResult.Yes;
                var data = Data(tools.Invoke("read_user_file", Json.Serialize(new { Path = file })));
                Assert.AreEqual("étè", data["Text"]); Assert.AreEqual(false,data["Truncated"]);
                tools.CurrentProviderName = "FixtureProvider";
                tools.ConfirmFile = (owner,text,title) => { StringAssert.Contains(text,"FixtureProvider"); return DialogResult.Yes; };
                File.WriteAllBytes(file, new byte[] { 65, 0, 66 });
                Failed(tools.Invoke("read_user_file", Json.Serialize(new { Path = file })), "binary");
                File.WriteAllText(file, new string('a',65537), new System.Text.UTF8Encoding(false));
                data = Data(tools.Invoke("read_user_file", Json.Serialize(new { Path = file })));
                Assert.AreEqual(65536, ((string)data["Text"]).Length); Assert.AreEqual(true,data["Truncated"]);
                Assert.AreEqual(65537,Convert.ToInt32(data["ByteLength"]));
                Failed(tools.Invoke("inspect_code_file", Json.Serialize(new { Path = "relative.bas" })), "relative path");
            }
        }

        [TestMethod]
        public async Task AsyncDispatchMatrixExecutesNativeBoundariesAndReturnsHostFailures()
        {
            var tools = Create();
            foreach (string name in new[] { "sign_project", "read_project_signature_dialog", "read_debug_options", "read_vbe_options",
                "quick_watch", "edit_watch", "debug_item", "immediate_execute", "debug_dialog", "respond_debug_dialog",
                "remove_watch", "add_watch", "debug_windows" })
            {
                var values = AsyncArguments(name);
                Success(await tools.InvokeAsync(name, Json.Serialize(values)),name);
                tools.Execute = request => Response.Failure("host declined");
                if (name != "debug_windows" && name != "debug_item" && name != "debug_dialog" && name != "respond_debug_dialog")
                    Failed(await tools.InvokeAsync(name, Json.Serialize(values)), name + ":host");
                tools.Execute = VbeToolBoundaryFixture.Execute;
            }
            Failed(await tools.InvokeAsync("debug_dialog","[]"),"dialog array");
            tools.Native.ReadSignatureDialog = p => { throw new InvalidOperationException("signature dialog unavailable"); };
            Failed(await tools.InvokeAsync("read_project_signature_dialog",Json.Serialize(Arguments("read_project_signature_dialog"))),"signature native exception");
            tools.Native.ReadDebugOptions = () => { throw new InvalidOperationException("options unavailable"); };
            Failed(await tools.InvokeAsync("read_debug_options","{}"),"options native exception");
            tools.Native.VerifyWatchRemoved = r => { throw new InvalidOperationException("watch native unavailable"); };
            Failed(await tools.InvokeAsync("remove_watch",Json.Serialize(Arguments("remove_watch"))),"remove native exception");
            tools.Native.ReadDebugDialog = () => { throw new InvalidOperationException("native declined"); };
            Failed(await tools.InvokeAsync("debug_dialog", "{}"), "native exception");
            tools.Native.CompleteAddWatch = r => { throw new InvalidOperationException("native add declined"); };
            Failed(await tools.InvokeAsync("add_watch", Json.Serialize(Arguments("add_watch"))), "native add exception");
            tools.Native.Capture = stack => new { Stack = stack };
            foreach (string args in new[] { "{}", "{\"IncludeCallStack\":true}", "{\"IncludeCallStack\":false}" })
                Success(await tools.InvokeAsync("debug_windows", args), args);
            foreach (string name in new[] { "sign_project", "read_project_signature_dialog", "quick_watch", "edit_watch", "remove_watch", "add_watch" })
                Failed(await tools.InvokeAsync(name, "{}"), name + ":missing");
            foreach (string args in new[] { "[]", "{\"Diagnostic\":\"x\"}", "{\"Diagnostic\":1,\"Button\":\"ok\"}", "{\"Diagnostic\":\"x\",\"Button\":1}" })
                Failed(await tools.InvokeAsync("respond_debug_dialog", args), args);
            foreach (string args in new[] { "[]", "{}", "{\"Action\":\"expand\",\"PathSegments\":[\"x\"]}", "{\"Pane\":\"locals\",\"PathSegments\":[\"x\"]}", "{\"Pane\":\"locals\",\"Action\":\"expand\"}", "{\"Pane\":\"locals\",\"Action\":\"expand\",\"PathSegments\":[\"x\"],\"Context\":\"p\"}" })
            {
                if (args.Contains("Context")) Success(await tools.InvokeAsync("debug_item",args),args);
                else Failed(await tools.InvokeAsync("debug_item",args),args);
            }
            var immediate = AsyncArguments("immediate_execute");
            foreach (string field in immediate.Keys.ToArray())
            {
                object original = immediate[field]; immediate.Remove(field);
                Failed(await tools.InvokeAsync("immediate_execute",Json.Serialize(immediate)), "missing "+field);
                immediate[field] = original is string ? (object)true : "2";
                Failed(await tools.InvokeAsync("immediate_execute",Json.Serialize(immediate)), "type "+field);
                immediate[field] = original;
            }
            immediate["ExpectedMode"] = 1;
            Failed(await tools.InvokeAsync("immediate_execute",Json.Serialize(immediate)),"changed mode");
            immediate["ExpectedMode"] = 3;
            Failed(await tools.InvokeAsync("immediate_execute",Json.Serialize(immediate)),"invalid mode");
            Failed(await tools.InvokeAsync("immediate_execute","[]"),"immediate array");
            tools.Settings.VbeEditApproval = "ReadOnly";
            Failed(await tools.InvokeAsync("remove_watch",Json.Serialize(Arguments("remove_watch"))),"remove policy");
        }

        [TestMethod]
        public async Task SignaturePersistenceMatrixHandlesMissingCertificateAndSaveRetries()
        {
            var tools = Create(); string args = Json.Serialize(Arguments("sign_project"));
            tools.Execute = r => Response.Success(new { Missing = true });
            Failed(await tools.InvokeAsync("sign_project",args),"missing certificate");
            tools.Execute = r => Response.Success("unexpected host shape");
            Failed(await tools.InvokeAsync("sign_project",args),"host shape");
            tools.Execute = VbeToolBoundaryFixture.Execute;
            foreach (bool saved in new[] { false, true })
            {
                tools.PersistSignature = p => new CodexVBE.Tests.Infrastructure.VbeToolPersistence { Saved = saved };
                Assert.AreEqual(!saved,Data(await tools.InvokeAsync("sign_project",args))["SaveRequired"]);
            }
            tools.PersistSignature = p => null;
            Assert.AreEqual(true,Data(await tools.InvokeAsync("sign_project",args))["SaveRequired"]);
            tools.PersistSignature = p => { throw new InvalidOperationException("save declined"); };
            tools.Execute = r => r.Command == "project_signature_status" ? Response.Failure("status declined") : VbeToolBoundaryFixture.Execute(r);
            var data = Data(await tools.InvokeAsync("sign_project",args));
            Assert.AreEqual("save declined",data["PersistenceError"]); Assert.AreEqual("status declined",data["HostStatusError"]);
            int attempts = 0;
            tools.PersistSignature = p => { if (++attempts == 2) return new CodexVBE.Tests.Infrastructure.VbeToolPersistence { Saved = true }; throw new InvalidOperationException("0x800AC472 busy"); };
            Assert.AreEqual(false,Data(await tools.InvokeAsync("sign_project",args))["SaveRequired"]); Assert.AreEqual(2,attempts);
            attempts = 0; tools.PersistSignature = p => { attempts++; throw new InvalidOperationException("0x800AC472 busy"); };
            data = Data(await tools.InvokeAsync("sign_project",args)); Assert.AreEqual(12,attempts);
            StringAssert.Contains((string)data["PersistenceError"],"busy");
        }

        [TestMethod]
        [STATestMethod]
        public void CompileMatrixUsesPostedUiWorkAndDistinguishesTimeoutDiagnosisAndErrors()
        {
            var prior = SynchronizationContext.Current;
            try
            {
                var tools = Create();
                SynchronizationContext.SetSynchronizationContext(new ImmediateContext());
                const string args = "{\"Project\":\"P\",\"ExpectedMode\":2}";
                Assert.AreEqual(true,Data(tools.InvokeAsync("compile_project",args).GetAwaiter().GetResult())["Compiled"]);
                tools.Native.AwaitCompileDialog = completed => { Assert.IsTrue(completed.Wait(5000)); return "compile diagnostic"; };
                var data = Data(tools.InvokeAsync("compile_project",args).GetAwaiter().GetResult());
                Assert.AreEqual(false,data["Compiled"]); Assert.AreEqual("NativeDiagnosticCaptured",data["Verification"]);
                tools.Execute = r => Response.Failure("compile failure"); Failed(tools.InvokeAsync("compile_project",args).GetAwaiter().GetResult(),"compile failure");
                tools.Execute = r => { throw new InvalidOperationException("compile exception"); };
                Failed(tools.InvokeAsync("compile_project",args).GetAwaiter().GetResult(),"compile exception");
                SynchronizationContext.SetSynchronizationContext(new DeferredContext());
                tools.Native.AwaitCompileDialog = completed => null;
                Failed(tools.InvokeAsync("compile_project",args).GetAwaiter().GetResult(),"timeout");
                SynchronizationContext.SetSynchronizationContext(new ImmediateContext());
                foreach (string invalid in new[] { "[]", "{}", "{\"ExpectedMode\":2}", "{\"Project\":\"P\"}", "{\"Project\":1,\"ExpectedMode\":2}", "{\"Project\":\" \",\"ExpectedMode\":2}", "{\"Project\":\"P\",\"ExpectedMode\":\"2\"}", "{\"Project\":\"P\",\"ExpectedMode\":1}" })
                    Failed(tools.InvokeAsync("compile_project",invalid).GetAwaiter().GetResult(),invalid);
            }
            finally { SynchronizationContext.SetSynchronizationContext(prior); }
        }
        [TestMethod]
        public void CodeEditReadbackMatrixRejectsStaleReadsAndHandlesUnchangedCodeOrSubscriberErrors()
        {
            var host=new VbeSessionTests.FakeVbe();
            var module=new VbeSessionTests.FakeModule("A\r\nB");
            var project=new VbeSessionTests.FakeProject {Name="P",FileName=@"C:\Temp\fixture.xlsm",Mode=2};
            project.VBComponents.Items.Add(new VbeSessionTests.FakeComponent {Name="M",Type=1,CodeModule=module});host.VBProjects.Add(project);
            var session=new VbeSession(host);
            var tools=new LlmVbeTools(session,null,new LlmSettings { VbeEditApproval="Automatic" });
            var logs=new List<string>(); tools.WriteLog=logs.Add;
            Func<string,string> arguments=text=>Json.Serialize(new {Project="P",Module="M",ExpectedSha256=(string)((dynamic)session.Execute(new Request {Command="read_module",Project="P",Module="M"}).Data).Sha256,StartLine=1,Count=2,Text=text});
            string args=arguments("A\r\nC");
            module.DeleteLines(1,module.CountOfLines); module.InsertLines(1,"stale");
            Failed(tools.Invoke("replace_lines",args),"stale code");
            module.DeleteLines(1,module.CountOfLines); module.InsertLines(1,"A\r\nB");
            tools.Execute=r=>Response.Failure("read unavailable");
            Failed(tools.Invoke("replace_lines",args),"read unavailable");
            tools.Execute=session.Execute;
            Success(tools.Invoke("replace_lines",arguments("A\r\nB")),"unchanged code");
            Success(tools.Invoke("replace_lines",arguments("A\r\nC")),"edit without subscriber");
            tools.CodeEdited+=change=> { throw new InvalidOperationException("subscriber failed"); };
            Success(tools.Invoke("replace_lines",arguments("A\r\nD")),"subscriber error is contained");
            int reads=0;
            tools.Execute=r=> r.Command=="read_module" && ++reads==2 ? Response.Failure("readback failed") : session.Execute(r);
            Success(tools.Invoke("replace_lines",arguments("A\r\nE")),"readback error is contained");
            Assert.AreEqual(2,logs.Count);
            StringAssert.Contains(logs[0],"subscriber failed");
            StringAssert.Contains(logs[1],"readback failed");
            tools.Execute=r=>r.Command=="replace_lines" ? Response.Failure("write refused") : session.Execute(r);
            Failed(tools.Invoke("replace_lines",arguments("A\r\nF")),"write refusal");
        }

        [TestMethod]
        public void RestorationMatrixHandlesNullEntriesSharedModuleHunksAndWriteFailure()
        {
            var host=new VbeSessionTests.FakeVbe();
            var module=new VbeSessionTests.FakeModule("A2\r\nKeep\r\nB2");
            var project=new VbeSessionTests.FakeProject {Name="P",FileName=@"C:\Temp\fixture.xlsm",Mode=2};
            project.VBComponents.Items.Add(new VbeSessionTests.FakeComponent {Name="M",Type=1,CodeModule=module});host.VBProjects.Add(project);
            var session=new VbeSession(host);
            var settings=new LlmSettings {VbeEditApproval="Automatic"};
            var tools=new LlmVbeTools(session,null,settings) {Mode=ChatMode.Plan};
            Assert.IsFalse(tools.RestoreChanges(new CodeChange[] {null},null).Ok);
            var change=new CodeChange("P","M","A\r\nKeep\r\nB","","A2\r\nKeep\r\nB2","",3);
            tools.Execute=r=>r.Command=="replace_lines" ? Response.Failure("write refused") : session.Execute(r);
            var failure=tools.RestoreChanges(new[] {change},null);
            Assert.IsFalse(failure.Ok);StringAssert.Contains(failure.Error,"write refused");Assert.IsFalse(change.Restored);
            tools.Execute=session.Execute;
            Assert.IsTrue(tools.RestoreChanges(new[] {change},0).Ok);
            Assert.AreEqual("A\r\nKeep\r\nB2",module.Code);Assert.IsFalse(change.Restored);
            Assert.IsFalse(tools.RestoreChanges(new[] {change},0).Ok);
            Assert.IsTrue(tools.RestoreChanges(new[] {change},null).Ok);Assert.IsTrue(change.Restored);
            Assert.AreEqual("A\r\nKeep\r\nB",module.Code);
            var first=new CodeChange("P","M","initial","","middle","",1);
            var second=new CodeChange("P","M","middle","","last","",1);
            module.DeleteLines(1,module.CountOfLines); module.InsertLines(1,"last");
            Assert.IsTrue(tools.RestoreChanges(new[] {first,second},null).Ok);
            Assert.AreEqual("initial",module.Code);Assert.IsTrue(first.Restored);Assert.IsTrue(second.Restored);
            tools.ValidateScope=()=> {throw new InvalidOperationException("stale scope");};
            Assert.AreEqual("stale scope",tools.RestoreChanges(new[] {first},null).Error);
            Assert.IsNotNull(Json.DeserializeObject(tools.LiveContextJson()));
        }

        [TestMethod]
        public void ToolResponseParserKeepsSuccessFailureNullAndInvalidJsonContracts()
        {
            var tools = new LlmVbeTools(null,null,new LlmSettings());
            var missing = tools.ReadToolResponse("null");
            Assert.IsNotNull(missing); Assert.IsFalse(missing.Ok); Assert.IsNull(missing.Error);
            var success = tools.ReadToolResponse(Json.Serialize(Response.Success(new {Value="fixture"})));
            Assert.IsTrue(success.Ok); Assert.AreEqual("fixture",Dict(success.Data)["Value"]);
            var failure = tools.ReadToolResponse(Json.Serialize(Response.Failure("declined")));
            Assert.IsFalse(failure.Ok); Assert.AreEqual("declined",failure.Error);
            Assert.ThrowsException<ArgumentException>(()=>tools.ReadToolResponse("{invalid}"));
        }

    }
}
