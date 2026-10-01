using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Infrastructure.Diagnostics;

namespace VBAi.Tests.Integration
{
    /// <summary>Opt-in live local-model checks through the production HTTP client with synthetic content only.</summary>
    [TestClass, TestCategory("Ollama"), DoNotParallelize]
    public sealed class OllamaQualificationTests
    {
        public TestContext TestContext { get; set; }

        /// <summary>Stops a live stream after its first fragment and verifies a fresh request can still complete.</summary>
        [TestMethod]
        public async Task LocalModelCancellationDoesNotPoisonTheNextConversation()
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_OLLAMA_TESTS") != "1")
                Assert.Inconclusive("Set VBAi_RUN_OLLAMA_TESTS=1 to exercise local streaming cancellation.");
            using (var wire = OllamaSyntheticWireCapture.ForHeadlessFixture("cancellation-recovery"))
            {
                var provider = LlmProvider.All.Single(p => p.IsOllama);
                var profile = OllamaQualificationProfile.Resolve();
                string model = profile.Model;
                var settings = new LlmSettings();
                profile.ApplyTo(settings);
                TestContext.WriteLine("Ollama qualification profile: " + profile.Describe());
                bool cancelled = false;
                Exception failure = null;
                using (var client = new LlmChatClient(provider, settings, model))
                {
                    client.TextDelta = text => { if (!cancelled && !string.IsNullOrEmpty(text)) { cancelled = true; client.Dispose(); } };
                    try
                    {
                        await wire.CompleteAsync("cancellation-stream", client, new List<object> { new { role = "user", content = "Write a long numbered list of 1000 everyday objects, starting immediately with item 1." } }, new object[0]);
                    }
                    catch (Exception error) { failure = error; }
                }
                Assert.IsTrue(cancelled, "Cancellation must follow a real provider fragment.");
                Assert.IsNotNull(failure, "A cancelled stream must not be published as a completed response.");
                Assert.IsTrue(failure is OperationCanceledException || failure is ObjectDisposedException || failure is System.IO.IOException,
                    "Unexpected cancellation failure: " + failure.GetType().FullName);
                using (var recovery = new LlmChatClient(provider, settings, model))
                {
                    recovery.TextDelta = text => { };
                    var answer = await wire.CompleteAsync("recovery", recovery, new List<object> { new { role = "user", content = "Reply with exactly READY." } }, new object[0]);
                    Assert.IsFalse(string.IsNullOrWhiteSpace(Convert.ToString(answer["content"])));
                }
                TestContext.WriteLine("Model={0}; cancellation={1}; subsequent streaming request completed.", model, failure.GetType().Name);
            }
        }

        /// <summary>Streams text, assembles a harmless tool call and consumes its synthetic result.</summary>
        [TestMethod]
        public async Task LocalModelStreamsAndCompletesSyntheticToolRoundTrip()
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_OLLAMA_TESTS") != "1")
                Assert.Inconclusive("Set VBAi_RUN_OLLAMA_TESTS=1 to exercise the local Ollama server with synthetic content.");
            using (var wire = OllamaSyntheticWireCapture.ForHeadlessFixture("synthetic-tool-roundtrip"))
            {
                var provider = LlmProvider.All.Single(p => p.IsOllama);
                var profile = OllamaQualificationProfile.Resolve();
                string model = profile.Model;
                var settings = new LlmSettings();
                profile.ApplyTo(settings);
                TestContext.WriteLine("Ollama qualification profile: " + profile.Describe());
                wire.SetPhase("catalogue");
                var models = await LlmChatClient.ListModelsAsync(provider, settings);
                Assert.IsTrue(models.Any(m => m.Id == model), "The qualification model must already be installed.");
                var json = new JavaScriptSerializer();
                var watch = Stopwatch.StartNew();
                using (var client = new LlmChatClient(provider, settings, model))
                {
                    int deltas = 0;
                    client.TextDelta = text => { if (!string.IsNullOrEmpty(text)) deltas++; };
                    var greeting = await wire.CompleteAsync("greeting", client, new List<object> {
                        new { role = "user", content = "Reply with exactly the word READY." }
                    }, new object[0]);
                    Assert.IsTrue(deltas > 0, "No streamed text was delivered.");
                    Assert.IsFalse(string.IsNullOrWhiteSpace(Convert.ToString(greeting["content"])));
                    var messages = new List<object> {
                        new { role = "system", content = "You are a synthetic test assistant. Call qualification_echo exactly once with marker VB_AI_42. After receiving the result, report it without another tool call." },
                        new { role = "user", content = "Call qualification_echo with marker VB_AI_42 now." }
                    };
                    object[] tools = { new { type = "function", function = new {
                        name = "qualification_echo", description = "Returns a synthetic marker. No filesystem, application, or network side effects.",
                        parameters = new { type = "object", properties = new { marker = new { type = "string" } }, required = new[] { "marker" } }
                    } } };
                    var response = await wire.CompleteAsync("synthetic-tool-request", client, messages, tools);
                    Assert.IsTrue(response.ContainsKey("tool_calls"), "The selected model did not call the supplied synthetic tool.");
                    var calls = (object[])response["tool_calls"];
                    Assert.AreEqual(1, calls.Length);
                    var call = (IDictionary<string, object>)calls[0];
                    var function = (IDictionary<string, object>)call["function"];
                    Assert.AreEqual("qualification_echo", function["name"]);
                    wire.RecordArguments("raw-before-deserialization", function["arguments"], null);
                    var arguments = json.DeserializeObject((string)function["arguments"]) as IDictionary<string, object>;
                    wire.RecordArguments("parsed-before-assertion", function["arguments"], arguments);
                    Assert.IsNotNull(arguments);
                    Assert.AreEqual("VB_AI_42", arguments["marker"]);
                    messages.Add(response);
                    messages.Add(new { role = "tool", tool_call_id = (string)call["id"], content = "Synthetic result VB_AI_42" });
                    var completed = await wire.CompleteAsync("synthetic-tool-result", client, messages, new object[0]);
                    StringAssert.Contains(Convert.ToString(completed["content"]), "VB_AI_42");
                    TestContext.WriteLine("Model={0}; streaming fragments={1}; elapsed ms={2}; production assembly={3}; MVID={4}; no native tool dispatch exercised.",
                        model, deltas, watch.ElapsedMilliseconds, typeof(LlmChatClient).Assembly.Location, typeof(LlmChatClient).Module.ModuleVersionId);
                }
            }
        }
    }
}
