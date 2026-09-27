using System;
using System.Collections.Generic;
using System.Threading;
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

        [TestMethod]
        public async Task AsyncDispatchReturnsReadOnlyStatusFromInMemorySession()
        {
            var tools = new LlmVbeTools(new VbeSession(new object()), null, new LlmSettings());
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
                await Failure(tools, "compile_project", "{\"Project\":\"P\",\"ExpectedMode\":2}",
                    "VBE UI context is unavailable");
            }
            finally { SynchronizationContext.SetSynchronizationContext(prior); }
        }

        [TestMethod]
        public async Task ImmediateExecuteRejectsModeChangedBeforeNativeExecution()
        {
            var host = new VbeSessionTests.FakeVbe();
            host.VBProjects.Add(new VbeSessionTests.FakeProject { Name = "P", FileName = @"C:\Temp\P.xlsm", Mode = 2 });
            var tools = new LlmVbeTools(new VbeSession(host), null,
                new LlmSettings { VbeEditApproval = "Automatic" });
            await Failure(tools, "immediate_execute",
                "{\"Project\":\"P\",\"ExpectedMode\":1,\"Text\":\"Debug.Print 1\"}",
                "Project mode changed");
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
