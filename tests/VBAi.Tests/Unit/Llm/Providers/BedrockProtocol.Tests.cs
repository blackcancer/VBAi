namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Web.Script.Serialization;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using VBAi;

    /// <summary>Vérifie la conversion des requêtes et réponses du protocole Bedrock Converse.</summary>
    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class ProviderProtocolTests
    {
        /// <summary>Refuse les sorties filtrées, incomplètes ou arrêtées pour une raison inconnue.</summary>
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

        /// <summary>Préserve l’entrée structurée d’un outil et son résultat lors d’un aller-retour Converse.</summary>
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
namespace VBAi.Tests.Unit
{
    /// <summary>Complète les scénarios de construction des requêtes Bedrock.</summary>
    public sealed partial class ProviderProtocolTests
    {
        /// <summary>Construit les outils, instructions système et groupes distincts de résultats.</summary>
        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public void BedrockRequestBuildsToolsSystemAndSeparatedResultGroups()
        {
            var request = Obj(BedrockProtocol.Request(new System.Collections.Generic.List<object> {
                new { role = "system", content = "instructions" },
                new { role = "user", content = "prompt" },
                new { role = "assistant", content = "" },
                new { role = "assistant", tool_calls = new object[] {
                    new { id = "a", function = new { name = "read", arguments = "{\"line\":3}" } },
                    new { id = "b", function = new { name = "empty" } } } },
                new { role = "tool", tool_call_id = "a" },
                new { role = "tool", tool_call_id = "b", content = "ok" },
                new { role = "user", content = "again" },
                new { role = "tool", tool_call_id = "c", content = "next" }
            }, new object[] { new { function = new { name = "read", description = "reads", parameters = new { type = "object" } } } }));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("instructions", Obj(((object[])request["system"])[0])["text"]);
            var messages = (object[])request["messages"];
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(5, messages.Length);
            var calls = (object[])Obj(messages[1])["content"];
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(3, Obj(Obj(Obj(calls[0])["toolUse"])["input"])["line"]);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(0, Obj(Obj(Obj(calls[1])["toolUse"])["input"]).Count);
            var results = (object[])Obj(messages[2])["content"];
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(2, results.Length);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("", Obj(((object[])Obj(Obj(results[0])["toolResult"])["content"])[0])["text"]);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(1, ((object[])Obj(messages[4])["content"]).Length);
            var spec = Obj(Obj(((object[])Obj(request["toolConfig"])["tools"])[0])["toolSpec"]);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("read", spec["name"]);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("object", Obj(Obj(spec["inputSchema"])["json"])["type"]);
            foreach (var stop in new[] { "end_turn", "stop_sequence" }) {
                var response = BedrockProtocol.Response(Obj(new { stopReason = stop, output = new { message = new { content = new object[] { new { ignored = true }, new { text = "one" }, new { text = "two" } } } } }));
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("one\ntwo", response["content"]);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsFalse(response.ContainsKey("tool_calls"));
            }
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.ThrowsException<System.InvalidOperationException>(() => BedrockProtocol.Response(Obj(new {})));
        }
    }
}
