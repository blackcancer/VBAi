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
namespace CodexVBE.Tests.Unit
{
    public sealed partial class ProviderProtocolTests
    {
        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public void ClaudeHelpersHandleNullMissingAndNonArrayValues()
        {
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsNull(ClaudeProtocol.Object(null));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsNull(ClaudeProtocol.Object(7));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsNull(ClaudeProtocol.Text(null, "missing"));
            var values = new System.Collections.Generic.Dictionary<string, object> { ["null"] = null, ["scalar"] = 7, ["array"] = new object[] { 1 } };
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreSame(values, ClaudeProtocol.Object(values));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsNull(ClaudeProtocol.Text(values, "missing"));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("", ClaudeProtocol.Text(values, "null"));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("7", ClaudeProtocol.Text(values, "scalar"));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(0, ClaudeProtocol.Array(null, "missing").Length);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(0, ClaudeProtocol.Array(values, "missing").Length);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(0, ClaudeProtocol.Array(values, "scalar").Length);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(0, ClaudeProtocol.Array(values, "null").Length);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreSame(values["array"], ClaudeProtocol.Array(values, "array"));
        }

        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public void ClaudeRequestBuildsNativeToolsSkipsEmptyAndSeparatesResultGroups()
        {
            var history = new System.Collections.Generic.List<object> {
                new { role = "user", content = "prompt" },
                new { role = "assistant", content = "" },
                new { role = "assistant", tool_calls = new object[] {
                    new { id = "a", function = new { name = "read", arguments = "{\"line\":3}" } },
                    new { id = "b", function = new { name = "empty" } } } },
                new { role = "tool", tool_call_id = "a" },
                new { role = "tool", tool_call_id = "b", content = "ok" },
                new { role = "user", content = "again" },
                new { role = "tool", tool_call_id = "c", content = "next" }
            };
            var request = Obj(ClaudeProtocol.Request("model", history, new object[] {
                new { function = new { name = "read", description = "reads", parameters = new { type = "object" } } }
            }));
            var messages = (object[])request["messages"];
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(5, messages.Length);
            var calls = (object[])Obj(messages[1])["content"];
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(3, Obj(Obj(calls[0])["input"])["line"]);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(0, Obj(Obj(calls[1])["input"]).Count);
            var results = (object[])Obj(messages[2])["content"];
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(2, results.Length);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("", Obj(results[0])["content"]);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("next", Obj(((object[])Obj(messages[4])["content"])[0])["content"]);
            var tool = Obj(((object[])request["tools"])[0]);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("read", tool["name"]);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("object", Obj(tool["input_schema"])["type"]);
            var response = ClaudeProtocol.Response(Obj(new { content = new object[] { new { type = "thinking" }, new { type = "text", text = "a" }, new { type = "text", text = "b" } } }));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("a\nb", response["content"]);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsFalse(response.ContainsKey("tool_calls"));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("", ClaudeProtocol.Response(Obj(new { }))["content"]);
        }
    }
}
