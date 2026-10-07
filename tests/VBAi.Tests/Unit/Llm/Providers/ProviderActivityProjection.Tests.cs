using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Web.Script.Serialization;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class ProviderActivityProjectionTests
    {
        [TestMethod]
        public void ValidatedGatewayMetadataUsesContextAndKeepsCodeAndSecretsOutOfTheTimeline()
        {
            using (var culture = new LocalizationScope("fr-FR"))
            {
                var json = new JavaScriptSerializer();
                var args = json.Serialize(new { ToolName = "read_module", ArgumentsJson = json.Serialize(new { Project = "Book", Module = "Module1", Code = "secret VBA", ApiKey = "secret token" }) });
                var activity = ProviderActivityProjection.Tool("call", "invoke_tool", args, "completed");
                Assert.AreEqual(UiText.Get("Reading VBA code") + " · Book · Module1", activity.Title);
                StringAssert.Contains(activity.Detail, "invoke_tool → read_module");
                Assert.IsFalse(activity.Detail.Contains("secret"));
                foreach (var invalid in new[] { "null", "malformed", "[]", json.Serialize(new { ToolName = "read_module", ArgumentsJson = "[]" }), json.Serialize(new { ToolName = "read_module", ArgumentsJson = "{}", Extra = true }), json.Serialize(new { ToolName = "invoke_tool", ArgumentsJson = "{}" }) })
                    StringAssert.Contains(ProviderActivityProjection.Tool("call", "invoke_tool", invalid, "failed").Detail, "invoke_tool");
            }
        }

        [DataTestMethod]
        [DataRow("{\"Ok\":true}", "completed")]
        [DataRow("{\"Ok\":false,\"Error\":\"Refused\"}", "failed")]
        [DataRow("{\"Ok\":true,\"Data\":{\"Uncertain\":true,\"Reason\":\"Check state\"}}", "interrupted")]
        [DataRow("{\"Ok\":false,\"Data\":{\"Uncertain\":true}}", "interrupted")]
        [DataRow("{\"Ok\":\"true\"}", "failed")]
        [DataRow("null", "failed")]
        [DataRow("malformed", "failed")]
        public void ToolStatusUsesActualReceiptsAndUncertaintyIsNeverClaimedAsSuccess(string receipt, string expected)
        {
            Assert.AreEqual(expected, ProviderActivityProjection.ToolOutcome(receipt));
        }

        [TestMethod]
        public void ScopesAreDetachedAndDiagnosticsReadOnlyPublicFailureFields()
        {
            var original = new CodexAgentActivity { Id = "block", Kind = "reasoning", Detail = "public", Status = "completed", DurationMs = 123, Append = true };
            var first = ProviderActivityProjection.WithScope(original, "request1");
            var second = ProviderActivityProjection.WithScope(original, "request2");
            Assert.AreEqual("block", original.Id);
            Assert.AreNotEqual(first.Id, second.Id);
            Assert.AreEqual(original.Detail, first.Detail); Assert.AreEqual(original.DurationMs, first.DurationMs); Assert.IsTrue(first.Append);
            Assert.IsNull(ProviderActivityProjection.WithScope(null, "request"));
            Assert.IsNull(ProviderActivityProjection.WithScope(new CodexAgentActivity(), "request"));
            Assert.AreEqual("Refused", ProviderActivityProjection.ToolDiagnostic("{\"Ok\":false,\"Error\":\"Refused\",\"Data\":{\"Code\":\"secret\"}}"));
            Assert.AreEqual("Check state", ProviderActivityProjection.ToolDiagnostic("{\"Ok\":true,\"Data\":{\"Uncertain\":true,\"Reason\":\"Check state\"}}"));
            Assert.IsNull(ProviderActivityProjection.ToolDiagnostic("{\"Ok\":true,\"Data\":{\"Reason\":\"secret\",\"Code\":\"secret\"}}"));
        }
    }
}
