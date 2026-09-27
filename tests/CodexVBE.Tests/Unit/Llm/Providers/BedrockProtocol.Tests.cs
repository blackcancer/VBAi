namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Web.Script.Serialization;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using CodexVBE;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class ProviderProtocolTests
    {
        [TestMethod]
        public void BedrockResponseRejectsFilteredOrIncompleteOutput()
        {
            var output = new
            {
                message = new
                {
                    content = new object[]
                    {
                        new
                        {
                            text = "ok"
                        }
                    }
                }
            };
            foreach (var stopReason in new[]
            {
                "max_tokens",
                "content_filtered",
                "unknown"
            }

            )
            {
                var error = Assert.ThrowsException<InvalidOperationException>(() => BedrockProtocol.Response(Obj(new { output, stopReason })));
                StringAssert.Contains(error.Message, stopReason);
            }
        }

        [TestMethod]
        public void BedrockRoundTripKeepsStructuredToolInputAndResult()
        {
            var input = new
            {
                module = "M1",
                line = 3
            };
            var response = BedrockProtocol.Response(Obj(new { stopReason = "tool_use", output = new { message = new { content = new object[] { new { text = "Reading" }, new { toolUse = new { toolUseId = "id-1", name = "read_module", input } } } } } }));
            var history = new List<object>
            {
                response,
                new
                {
                    role = "tool",
                    tool_call_id = "id-1",
                    content = "Module text"
                }
            };
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
