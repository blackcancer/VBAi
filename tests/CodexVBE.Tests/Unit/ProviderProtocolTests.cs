using System;
using System.Collections.Generic;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using CodexVBE;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class ProviderProtocolTests
    {
        private static readonly JavaScriptSerializer Json = new JavaScriptSerializer();

        private static IDictionary<string, object> Obj(object value)
        {
            return ClaudeProtocol.Object(value);
        }

        [TestMethod]
        public void ClaudeRequestPreservesToolResultOrderAndOpaqueContinuation()
        {
            var original = new object[] { new { type = "thinking", thinking = "opaque", signature = "sig" } };
            var history = new List<object> {
                new { role = "system", content = "First instruction" },
                new { role = "system", content = "Second instruction" },
                new { role = "assistant", content = "", _claude_content = original },
                new { role = "tool", tool_call_id = "one", content = "result 1" },
                new { role = "tool", tool_call_id = "two", content = "result 2" }
            };
            var request = Obj(ClaudeProtocol.Request("claude-test", history, new object[0]));
            Assert.AreEqual("First instruction\n\nSecond instruction", request["system"]);
            var messages = (object[])request["messages"];
            var assistant = Obj(messages[0]);
            Assert.AreEqual("sig", Obj(((object[])assistant["content"])[0])["signature"]);
            var results = (object[])Obj(messages[1])["content"];
            Assert.AreEqual("one", Obj(results[0])["tool_use_id"]);
            Assert.AreEqual("two", Obj(results[1])["tool_use_id"]);
        }

        [TestMethod]
        public void ClaudeResponseRejectsTokenLimitAndRetainsToolInput()
        {
            var content = new object[] {
                new { type = "text", text = "Terminé" },
                new { type = "tool_use", id = "call-1", name = "read_module", input = new { module = "M1" } }
            };
            var response = ClaudeProtocol.Response(Obj(new { content, stop_reason = "tool_use" }));
            Assert.AreEqual("Terminé", response["content"]);
            Assert.AreEqual("call-1", Obj(((object[])response["_claude_content"])[1])["id"]);
            var call = Obj(((object[])response["tool_calls"])[0]);
            var function = Obj(call["function"]);
            Assert.AreEqual("read_module", function["name"]);
            Assert.AreEqual("M1", Obj(Json.DeserializeObject((string)function["arguments"]))["module"]);

            var error = Assert.ThrowsException<InvalidOperationException>(
                () => ClaudeProtocol.Response(Obj(new { content, stop_reason = "max_tokens" })));
            StringAssert.Contains(error.Message, "limite");
        }

        [TestMethod]
        public void BedrockResponseRejectsFilteredOrIncompleteOutput()
        {
            var output = new { message = new { content = new object[] { new { text = "ok" } } } };
            foreach (var stopReason in new[] { "max_tokens", "content_filtered", "unknown" })
            {
                var error = Assert.ThrowsException<InvalidOperationException>(
                    () => BedrockProtocol.Response(Obj(new { output, stopReason })));
                StringAssert.Contains(error.Message, stopReason);
            }
        }

        [TestMethod]
        public void BedrockRoundTripKeepsStructuredToolInputAndResult()
        {
            var input = new { module = "M1", line = 3 };
            var response = BedrockProtocol.Response(Obj(new {
                stopReason = "tool_use",
                output = new { message = new { content = new object[] {
                    new { text = "Reading" },
                    new { toolUse = new { toolUseId = "id-1", name = "read_module", input } }
                } } }
            }));
            var history = new List<object> { response, new { role = "tool", tool_call_id = "id-1", content = "Module text" } };
            var request = Obj(BedrockProtocol.Request(history, new object[0]));
            var messages = (object[])request["messages"];
            var assistantBlocks = (object[])Obj(messages[0])["content"];
            Assert.AreEqual("id-1", Obj(Obj(assistantBlocks[1])["toolUse"])["toolUseId"]);
            var result = Obj(((object[])Obj(messages[1])["content"])[0]);
            Assert.AreEqual("id-1", Obj(result["toolResult"])["toolUseId"]);
            var call = Obj(((object[])response["tool_calls"])[0]);
            var function = Obj(call["function"]);
            Assert.AreEqual("M1", Obj(Json.DeserializeObject((string)function["arguments"]))["module"]);
        }
    }
}
