using System;
using System.Collections.Generic;
using System.Web.Script.Serialization;
using System.Threading.Tasks;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class LlmMonacoValidationTests
    {
        private static LlmVbeTools Tools(string policy = "Automatic") => new LlmVbeTools(null, null, new LlmSettings { VbeEditApproval = policy })
        { MonacoModule = (p, m) => throw new InvalidOperationException("HOST WAS ACCESSED") };

        [TestMethod]
        public async Task EveryMonacoContractValidatesRequiredTypesWhitespaceAndDispatchesAsynchronously()
        {
            var json = new JavaScriptSerializer();
            int count = 0;
            foreach (var definition in LlmVbeTools.Definitions)
            {
                var root = (Dictionary<string, object>)json.DeserializeObject(json.Serialize(definition));
                var function = (Dictionary<string, object>)root["function"];
                string name = (string)function["name"];
                if (!name.StartsWith("monaco_", StringComparison.Ordinal)) continue;
                count++;
                var parameters = (Dictionary<string, object>)function["parameters"];
                var properties = (Dictionary<string, object>)parameters["properties"];
                var valid = new Dictionary<string, object>();
                foreach (var field in properties)
                {
                    var schema = (Dictionary<string, object>)field.Value;
                    valid[field.Key] = (string)schema["type"] == "integer" ? (object)1 : field.Key == "Text" ? "" : "P";
                }
                var tools = Tools();
                StringAssert.Contains(await tools.InvokeAsync(name, json.Serialize(valid)), "HOST WAS ACCESSED", name + ":valid dispatch");
                StringAssert.Contains(tools.Invoke(name, json.Serialize(valid)), "InvokeAsync", name + ":sync refused");
                foreach (string required in (object[])parameters["required"])
                {
                    var missing = new Dictionary<string, object>(valid); missing.Remove(required);
                    StringAssert.Contains(await tools.InvokeAsync(name, json.Serialize(missing)), required + " is required", name);
                }
                foreach (var field in properties)
                {
                    var invalid = new Dictionary<string, object>(valid);
                    bool integer = (string)((Dictionary<string, object>)field.Value)["type"] == "integer";
                    invalid[field.Key] = integer ? (object)"1" : 1;
                    StringAssert.Contains(await tools.InvokeAsync(name, json.Serialize(invalid)), "Invalid Monaco argument", name + ":type " + field.Key);
                    if (integer || field.Key == "Text") continue;
                    invalid[field.Key] = "  ";
                    StringAssert.Contains(await tools.InvokeAsync(name, json.Serialize(invalid)), "Invalid Monaco argument", name + ":whitespace " + field.Key);
                }
                valid["Extra"] = true;
                StringAssert.Contains(await tools.InvokeAsync(name, json.Serialize(valid)), "Unexpected argument", name);
            }
            Assert.AreEqual(5, count);
        }

        [TestMethod]
        public async Task RefusesMalformedArgumentsAndCrossProjectReadsBeforeHostAccess()
        {
            var tools = Tools();
            StringAssert.Contains(await tools.InvokeAsync("monaco_edit", "{\"Project\":\"P\",\"Module\":\"M\",\"ExpectedVersion\":\"1\",\"Text\":\"\"}"), "Invalid Monaco argument");
            StringAssert.Contains(await tools.InvokeAsync("monaco_read", "{\"Project\":\"P\",\"Module\":\"M\",\"Extra\":true}"), "Unexpected argument");
            tools.BoundProject = "Other";
            StringAssert.Contains(new JavaScriptSerializer().Deserialize<Response>(await tools.InvokeAsync("monaco_read", "{\"Project\":\"P\",\"Module\":\"M\"}")).Error, UiText.Get("Read access to another project is not authorized for this conversation."));
            tools.SetReadAccess(new[] { "P" }, false);
            StringAssert.Contains(await tools.InvokeAsync("monaco_read", "{\"Project\":\"P\",\"Module\":\"M\"}"), "HOST WAS ACCESSED");
            StringAssert.Contains(await tools.InvokeAsync("monaco_edit", "{\"Project\":\"P\",\"Module\":\"M\",\"ExpectedVersion\":1,\"Text\":\"\"}"), "autre projet");
            StringAssert.Contains(tools.Invoke("monaco_open", "{}"), "InvokeAsync");
        }

        [TestMethod]
        public async Task EditingHonorsModeScopeAndReadOnlyPolicy()
        {
            const string args = "{\"Project\":\"P\",\"Module\":\"M\",\"ExpectedVersion\":1,\"Text\":\"\"}";
            var tools = Tools("ReadOnly");
            StringAssert.Contains(await tools.InvokeAsync("monaco_edit", args), "Read-only mode");
            tools = Tools(); tools.Mode = ChatMode.Discussion;
            StringAssert.Contains(await tools.InvokeAsync("monaco_edit", args), "monaco_edit");
            tools = Tools(); tools.ValidateScope = () => throw new InvalidOperationException("expired scope");
            StringAssert.Contains(await tools.InvokeAsync("monaco_read", "{\"Project\":\"P\",\"Module\":\"M\"}"), "expired scope");
        }
    }
}
