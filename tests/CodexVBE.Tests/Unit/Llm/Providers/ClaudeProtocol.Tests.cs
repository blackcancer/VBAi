namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Web.Script.Serialization;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using CodexVBE;

    public sealed partial class ProviderProtocolTests
    {
        [TestMethod]
        public void ClaudeRequestPreservesToolResultOrderAndOpaqueContinuation()
        {
            var original = new object[]
            {
                new
                {
                    type = "thinking",
                    thinking = "opaque",
                    signature = "sig"
                }
            };
            var history = new List<object>
            {
                new
                {
                    role = "system",
                    content = "First instruction"
                },
                new
                {
                    role = "system",
                    content = "Second instruction"
                },
                new
                {
                    role = "assistant",
                    content = "",
                    _claude_content = original
                },
                new
                {
                    role = "tool",
                    tool_call_id = "one",
                    content = "result 1"
                },
                new
                {
                    role = "tool",
                    tool_call_id = "two",
                    content = "result 2"
                }
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
            var content = new object[]
            {
                new
                {
                    type = "text",
                    text = "Terminé"
                },
                new
                {
                    type = "tool_use",
                    id = "call-1",
                    name = "read_module",
                    input = new
                    {
                        module = "M1"
                    }
                }
            };
            var response = ClaudeProtocol.Response(Obj(new { content, stop_reason = "tool_use" }));
            Assert.AreEqual("Terminé", response["content"]);
            Assert.AreEqual("call-1", Obj(((object[])response["_claude_content"])[1])["id"]);
            var call = Obj(((object[])response["tool_calls"])[0]);
            var function = Obj(call["function"]);
            Assert.AreEqual("read_module", function["name"]);
            Assert.AreEqual("M1", Obj(Json.DeserializeObject((string)function["arguments"]))["module"]);
            var error = Assert.ThrowsException<InvalidOperationException>(() => ClaudeProtocol.Response(Obj(new { content, stop_reason = "max_tokens" })));
            StringAssert.Contains(error.Message, "limite");
        }
    }
}
