using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class LlmVbeToolContractTests
    {
        private static readonly JavaScriptSerializer Json = new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 };

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
            var settings = new LlmSettings { VbeEditApproval = "Automatic" };
            var tools = new LlmVbeTools(null, null, settings) { Mode = ChatMode.Plan };
            IsFailure(tools.Invoke("create_module", "{}"), "create_module");
            tools.Mode = ChatMode.Agent;
            settings.VbeEditApproval = "ReadOnly";
            IsFailure(tools.Invoke("create_module", "{\"Project\":\"P\",\"Module\":\"M\",\"ExpectedMode\":2}"),
                "VBE");
            settings.VbeEditApproval = "invalid";
            IsFailure(tools.Invoke("create_module", "{\"Project\":\"P\",\"Module\":\"M\",\"ExpectedMode\":2}"),
                "VBE");
            tools.ValidateScope = () => { throw new InvalidOperationException("scope changed"); };
            IsFailure(tools.Invoke("list_projects", "{}"), "scope changed");
        }

        [TestMethod]
        public void BoundProjectAndFilePathAreCheckedBeforeHostAccess()
        {
            var tools = new LlmVbeTools(null, null, new LlmSettings()) { BoundProject = "WorkbookA" };
            IsFailure(tools.Invoke("create_module", "{\"Project\":\"WorkbookB\",\"Module\":\"M\",\"ExpectedMode\":2}"),
                "autre projet");
            IsFailure(tools.Invoke("inspect_code_file", "{\"Path\":\"C:\\\\Temp\\\\code.bas\"}"),
                "explicitement");
            IsFailure(tools.Invoke("set_form_node_property", "{\"Project\":\"WorkbookA\",\"Form\":\"F\",\"ControlPath\":\"X\",\"ExpectedTreeVersion\":\"v\",\"Property\":\"Caption\",\"Value\":[]}"),
                "Value must be");
        }

        private static IDictionary<string, object> Dict(object value)
        {
            return (IDictionary<string, object>)value;
        }

        private static void IsFailure(string serialized, string fragment)
        {
            var response = Json.Deserialize<Response>(serialized);
            Assert.IsFalse(response.Ok);
            StringAssert.Contains(response.Error, fragment);
        }
    }
}
