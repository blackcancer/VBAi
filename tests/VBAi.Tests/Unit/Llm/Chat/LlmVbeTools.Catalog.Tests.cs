using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using VBAi;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ToolCatalogTests
    {
        private static string Name(object definition) => Convert.ToString(((Dictionary<string, object>)((Dictionary<string, object>)new JavaScriptSerializer().DeserializeObject(new JavaScriptSerializer().Serialize(definition)))["function"])["name"]);
        [TestMethod]
        public async Task FamiliesDiscoverEveryToolAndDiscussionCannotExposeOrExecuteWrites()
        {
            var json = new JavaScriptSerializer();
            var tools = new LlmVbeTools(null, null, new LlmSettings { VbeEditApproval = "Automatic" });
            tools.Execute = request => throw new InvalidOperationException("Host must not be reached");
            var core = tools.CatalogForProvider(); Assert.IsTrue(core.Length < 15);
            Console.WriteLine("Catalogue schemas: core=" + core.Length + " / " + json.Serialize(core).Length + " characters; full=" + LlmVbeTools.Definitions.Length + " / " + json.Serialize(LlmVbeTools.Definitions).Length + " characters. Serialized size only, not model token usage.");
            Assert.IsFalse(core.Any(d => Name(d) == "replace_lines"));
            var all = await tools.InvokeAsync("discover_tools", "{\"Family\":\"all\"}");
            Assert.IsTrue(json.Deserialize<Response>(all).Ok, all);
            var discovered = (Dictionary<string, object>)((Dictionary<string, object>)json.DeserializeObject(all))["Data"];
            var discoveredNames = ((object[])discovered["Tools"]).Select(Name).ToArray();
            CollectionAssert.AreEquivalent(LlmVbeTools.Definitions.Select(Name).Where(n => !LlmVbeTools.IsCatalogTool(n)).ToArray(), discoveredNames);
            Assert.IsTrue(tools.CatalogForProvider().Length <= 64);
            Assert.IsTrue(tools.CatalogForProvider().Any(d => Name(d) == "invoke_tool"));
            await tools.InvokeAsync("discover_tools", "{\"Family\":\"forms\"}");
            Assert.AreEqual("forms", LlmVbeTools.ToolFamily(Name(tools.CatalogForProvider()[8])));
            tools.Mode = ChatMode.Discussion;
            Assert.IsFalse(tools.CatalogForProvider().Any(d => Name(d) == "replace_lines"));
            Assert.IsFalse(tools.CatalogForProvider(true).Any(d => Name(d) == "replace_lines"));
            string blocked = await tools.InvokeAsync("invoke_tool", json.Serialize(new { ToolName = "replace_lines", ArgumentsJson = "{}" }));
            Assert.IsFalse(json.Deserialize<Response>(blocked).Ok); Assert.IsFalse(blocked.Contains("Host must not"));
            string recursive = await tools.InvokeAsync("invoke_tool", json.Serialize(new { ToolName = "invoke_tool", ArgumentsJson = "{}" }));
            StringAssert.Contains(recursive, "Recursive");
            tools.BoundProject = "Allowed";
            string crossProject = await tools.InvokeAsync("invoke_tool", json.Serialize(new { ToolName = "read_module", ArgumentsJson = json.Serialize(new { Project = "Other", Module = "M" }) }));
            Assert.IsFalse(json.Deserialize<Response>(crossProject).Ok); Assert.IsFalse(crossProject.Contains("Host must not"));
        }
    }
}
namespace VBAi.Tests.Unit
{
    [Microsoft.VisualStudio.TestTools.UnitTesting.TestClass,Microsoft.VisualStudio.TestTools.UnitTesting.TestCategory("Unit")]
    public sealed class CatalogBoundaryTests
    {
        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public async System.Threading.Tasks.Task CatalogClassifiesEveryFamilyAndRejectsEachInvalidGatewayShape()
        {
            var json=new System.Web.Script.Serialization.JavaScriptSerializer();
            var tools=new LlmVbeTools(null,null,new LlmSettings());
            foreach(var name in new[]{"git_status","form_tree","designer_state","debug_state","watch_state","breakpoint_set","run_any","step_any","compile_project","immediate_execute","read_immediate","inspect_local_scalars","invoke_command","list_commands","monaco_read","code_read","module_read","procedure_read","rename_any","replace_lines","project_symbols","status"}) {
                var expected=name.StartsWith("git_")?"git":name.Contains("form")||name.Contains("designer")?"forms":name.Contains("debug")||name.Contains("watch")||name.Contains("breakpoint")||name.StartsWith("run_")||name.StartsWith("step_")||name=="compile_project"||name=="immediate_execute"||name=="read_immediate"||name=="inspect_local_scalars"||name=="invoke_command"||name=="list_commands"?"debug":name.StartsWith("monaco_")||name.Contains("code")||name.Contains("module")||name.Contains("procedure")||name.Contains("rename")||name=="replace_lines"||name=="project_symbols"?"code":"environment";
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(expected,LlmVbeTools.ToolFamily(name),name);
            }
            foreach(var args in new[]{"null","[]","{}","{\"Family\":1}","{\"Other\":\"code\"}","{\"Family\":\"code\",\"extra\":1}","{\"Family\":\"unknown\"}","broken"})
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsFalse(json.Deserialize<Response>(await tools.InvokeCatalogAsync("discover_tools",args)).Ok,args);
            foreach(var args in new[]{"{}","{\"ToolName\":1,\"ArgumentsJson\":\"{}\"}","{\"Other\":\"status\",\"ArgumentsJson\":\"{}\"}","{\"ToolName\":\"status\",\"Other\":\"{}\"}","{\"ToolName\":\"status\",\"ArgumentsJson\":1}","{\"ToolName\":\"discover_tools\",\"ArgumentsJson\":\"{}\"}"})
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsFalse(json.Deserialize<Response>(await tools.InvokeCatalogAsync("invoke_tool",args)).Ok,args);
            foreach(var mode in new[]{ChatMode.Agent,ChatMode.Discussion,ChatMode.Plan}) {
                tools.Mode=mode;
                foreach(var family in new[]{"code","forms","debug","git","environment","all"}) {
                    Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsTrue(json.Deserialize<Response>(await tools.InvokeCatalogAsync("discover_tools",json.Serialize(new{Family=family}))).Ok);
                    Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsTrue(tools.CatalogForProvider().Length<=64);
                    Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsTrue(tools.CatalogForProvider(true).Length<15);
                }
                tools.ResetCatalog();Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsTrue(tools.CatalogForProvider().Length<15);
            }
        }

        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public async System.Threading.Tasks.Task ReadImmediateGatewayKeepsModePolicyAndPrivacyGuards()
        {
            var json = new System.Web.Script.Serialization.JavaScriptSerializer();
            var settings = new LlmSettings { VbeEditApproval = "ReadOnly" };
            var tools = new LlmVbeTools(null, null, settings) { BoundProject = "A" };
            int reads = 0;
            tools.Execute = request =>
            {
                if (request.Command == "debug_state")
                {
                    var state = new System.Dynamic.ExpandoObject();
                    ((System.Collections.Generic.IDictionary<string, object>)state)["Mode"] = 1;
                    return Response.Success(state);
                }
                throw new System.InvalidOperationException("Unexpected synchronous command.");
            };
            tools.ReadImmediateNative = request =>
            {
                reads++;
                return System.Threading.Tasks.Task.FromResult<object>(new { Text = "private output" });
            };
            const string arguments = "{\"Project\":\"A\",\"ExpectedMode\":1}";
            string gateway = json.Serialize(new { ToolName = "read_immediate", ArgumentsJson = arguments });
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsFalse(json.Deserialize<Response>(tools.Invoke("read_immediate", arguments)).Ok);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsFalse(json.Deserialize<Response>(await tools.InvokeCatalogAsync("invoke_tool", gateway)).Ok);
            settings.VbeEditApproval = "Automatic";
            tools.Mode = ChatMode.Plan;
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsFalse(json.Deserialize<Response>(await tools.InvokeCatalogAsync("invoke_tool", gateway)).Ok);
            tools.Mode = ChatMode.Agent;
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsFalse(json.Deserialize<Response>(await tools.InvokeCatalogAsync("invoke_tool", gateway)).Ok);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(0, reads);
            tools.SetReadAccess(new string[0], true);
            var granted = json.Deserialize<Response>(await tools.InvokeCatalogAsync("invoke_tool", gateway));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsTrue(granted.Ok, granted.Error);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(1, reads);
        }

        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public async System.Threading.Tasks.Task LocalScalarGatewayKeepsModePolicyAndPrivacyGuards()
        {
            var json = new JavaScriptSerializer();
            var settings = new LlmSettings { VbeEditApproval = "ReadOnly" };
            var tools = new LlmVbeTools(null, null, settings) { BoundProject = "A" };
            int calls = 0;
            tools.Execute = request => {
                var state = new System.Dynamic.ExpandoObject();
                ((IDictionary<string, object>)state)["Mode"] = 1;
                return Response.Success(state);
            };
            tools.InspectLocalScalarsNative = request => {
                calls++;
                return Task.FromResult<object>(new { Partial = true });
            };
            const string arguments = "{\"Project\":\"A\",\"Module\":\"M\",\"Procedure\":\"Run\",\"ExpectedSha256\":\"aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa\",\"ExpectedMode\":1}";
            string gateway = json.Serialize(new { ToolName = "inspect_local_scalars", ArgumentsJson = arguments });
            Assert.IsFalse(json.Deserialize<Response>(tools.Invoke("inspect_local_scalars", arguments)).Ok);
            Assert.IsFalse(json.Deserialize<Response>(await tools.InvokeCatalogAsync("invoke_tool", gateway)).Ok);
            settings.VbeEditApproval = "Automatic";
            tools.Mode = ChatMode.Plan;
            Assert.IsFalse(json.Deserialize<Response>(await tools.InvokeCatalogAsync("invoke_tool", gateway)).Ok);
            tools.Mode = ChatMode.Agent;
            Assert.IsFalse(json.Deserialize<Response>(await tools.InvokeCatalogAsync("invoke_tool", gateway)).Ok);
            Assert.AreEqual(0, calls);
            tools.SetReadAccess(new string[0], true);
            var allowed = json.Deserialize<Response>(await tools.InvokeCatalogAsync("invoke_tool", gateway));
            Assert.IsTrue(allowed.Ok, allowed.Error);
            Assert.AreEqual(1, calls);
        }
    }
}
