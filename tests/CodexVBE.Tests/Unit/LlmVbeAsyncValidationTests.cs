using System;
using System.Collections.Generic;
using System.Web.Script.Serialization;
using System.Threading.Tasks;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class LlmVbeAsyncValidationTests
    {
        private static readonly JavaScriptSerializer Json = new JavaScriptSerializer();

        [TestMethod]
        public async Task NativeDebuggerToolsRejectMalformedArgumentsBeforeUiAccess()
        {
            var settings = new LlmSettings { VbeEditApproval = "Automatic" };
            var tools = new LlmVbeTools(null, null, settings);
            await Failure(tools, "debug_item", "{}", "Pane, Action and PathSegments");
            await Failure(tools, "debug_item",
                "{\"Pane\":\"locals\",\"Action\":\"expand\",\"PathSegments\":[\"x\"],\"Extra\":1}",
                "Unexpected debug_item argument");
            await Failure(tools, "immediate_execute", "{}", "Project, ExpectedMode and Text");
            await Failure(tools, "immediate_execute",
                "{\"Project\":\"P\",\"ExpectedMode\":0,\"Text\":\"Debug.Print 1\"}",
                "Project, ExpectedMode and Text");
            await Failure(tools, "respond_debug_dialog", "{}", "Exact Diagnostic and Button");
            await Failure(tools, "debug_dialog", "{\"Unexpected\":1}", "debug_dialog has no arguments");
            await Failure(tools, "compile_project", "{}", "Project and ExpectedMode=2");
            await Failure(tools, "compile_project",
                "{\"Project\":\"P\",\"ExpectedMode\":1}", "Project and ExpectedMode=2");
        }

        [TestMethod]
        public async Task NativeEvaluationAndMutationRequireAutomaticPolicy()
        {
            var settings = new LlmSettings { VbeEditApproval = "ReadOnly" };
            var tools = new LlmVbeTools(null, null, settings);
            await Failure(tools, "quick_watch", "{}", "Automatic VBE edit policy");
            await Failure(tools, "edit_watch", "{}", "Automatic VBE edit policy");
            await Failure(tools, "immediate_execute", "{}", "Automatic VBE edit policy");
            await Failure(tools, "respond_debug_dialog", "{}", "Automatic VBE edit policy");
        }

        [TestMethod]
        public async Task ScopeAndProjectBindingApplyToAsyncInvocationBeforeHostAccess()
        {
            var tools = new LlmVbeTools(null, null,
                new LlmSettings { VbeEditApproval = "Automatic" }) { BoundProject = "WorkbookA" };
            await Failure(tools, "compile_project",
                "{\"Project\":\"WorkbookB\",\"ExpectedMode\":2}", "autre projet");
            tools.ValidateScope = () => { throw new InvalidOperationException("stale conversation scope"); };
            await Failure(tools, "debug_item", "{}", "stale conversation scope");
        }

        private static async Task Failure(LlmVbeTools tools, string name, string arguments, string fragment)
        {
            string serialized = await tools.InvokeAsync(name, arguments);
            var response = Json.Deserialize<Response>(serialized);
            Assert.IsFalse(response.Ok, name + " accepted invalid arguments");
            StringAssert.Contains(response.Error, fragment);
        }
    }
}
