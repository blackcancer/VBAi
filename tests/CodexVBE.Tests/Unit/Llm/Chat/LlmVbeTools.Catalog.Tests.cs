using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
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
