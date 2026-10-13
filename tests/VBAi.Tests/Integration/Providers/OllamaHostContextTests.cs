using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using VBAi.Tests.Infrastructure.Diagnostics;

namespace VBAi.Tests.Integration
{
    /// <summary>Separates model compliance with VBAi's real context from native host or rendering behavior.</summary>
    [TestClass, TestCategory("Ollama"), DoNotParallelize]
    public sealed class OllamaHostContextTests
    {
        /// <summary>Reports the exact model profile and synthetic responses in the original test receipt.</summary>
        public TestContext TestContext { get; set; }

        /// <summary>Allows bounded inspection before one exact module read and an unprompted marker using the production context.</summary>
        /// <param name="host">Host label for evidence; no application is launched.</param>
        /// <param name="project">Invented disposable identifier used only in HTTP messages.</param>
        [DataTestMethod]
        [DataRow("Word", @"C:\VBAiSynthetic\Disposable.docm")]
        [DataRow("Excel", @"C:\VBAiSynthetic\Disposable.xlsm")]
        [DataRow("Access", @"C:\VBAiSynthetic\Disposable.accdb")]
        [DataRow("Publisher", @"C:\VBAiSynthetic\Disposable.pub")]
        [DataRow("Outlook", "Projet 1")]
        [DataRow("PowerPoint", @"C:\VBAiSynthetic\Disposable.pptm")]
        public async Task LocalModelReadsSyntheticHostMarkerAfterInterruptedConversation(string host, string project)
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_OLLAMA_TESTS") != "1")
                Assert.Inconclusive("Set VBAi_RUN_OLLAMA_TESTS=1 to run this synthetic HTTP context matrix.");

            var profile = OllamaQualificationProfile.Resolve();
            var settings = new LlmSettings();
            profile.ApplyTo(settings);
            var catalog = new LlmVbeTools(null, null, settings) { Mode = ChatMode.Discussion, BoundProject = project };
            catalog.Execute = _ => throw new InvalidOperationException("This fixture must never dispatch native tools.");
            object[] definitions = catalog.CatalogForProvider();
            var json = new JavaScriptSerializer();
            string marker = "VB_AI_" + Guid.NewGuid().ToString("N");
            string code = "Option Explicit\r\nPublic Const ObservedMarker As String = \"" + marker + "\"\r\n";
            var messages = new List<object>
            {
                new { role = "system", content = LlmVbeContext.DeveloperInstructions },
                new { role = "user", content = Request(OllamaOfficeChatScenario.StreamingRequest(), host, project) },
                new { role = "assistant", content = UiText.Get("The response did not complete. Actions may already have been applied; read the live code again before continuing.") },
                new { role = "user", content = Request(OllamaOfficeChatScenario.RecoveryRequest(), host, project) },
                new { role = "assistant", content = "UI_READY_42" },
                new { role = "user", content = Request(OllamaOfficeChatScenario.MarkerRequest(project, "Q028Marker", marker), host, project) }
            };
            TestContext.WriteLine("Synthetic HTTP only; no native host, file read or tool dispatch. Host={0}; {1}; product MVID={2}; core schemas={3}",
                host, profile.Describe(), typeof(LlmChatClient).Module.ModuleVersionId, definitions.Length);
            using (var wire = OllamaSyntheticWireCapture.ForHeadlessFixture("synthetic-host-context"))
            using (var client = await profile.CreateClientAsync(settings))
            {
                client.TextDelta = _ => { };
                var inspected = new HashSet<string>(StringComparer.Ordinal);
                bool moduleRead = false;
                // Four distinct preflight inspections, one read and one final answer.
                // Every request has the production deadline; no failed action is replayed.
                for (int round = 0; round < 6; round++)
                {
                    var response = await wire.CompleteAsync(host + " context round " + (round + 1), client, messages, catalog.CatalogForProvider());
                    TestContext.WriteLine("Synthetic response round {0}: {1}", round + 1, json.Serialize(response));
                    if (!moduleRead)
                        Assert.IsFalse(Convert.ToString(response["content"]).Contains(marker), "The marker is unavailable before the tool result.");
                    var calls = response.TryGetValue("tool_calls", out var rawCalls) ? rawCalls as object[] : null;
                    if (moduleRead)
                    {
                        Assert.IsTrue(calls == null || calls.Length == 0, "The result must be consumed without another tool call.");
                        StringAssert.Contains(Convert.ToString(response["content"]), marker);
                        return;
                    }
                    Assert.IsNotNull(calls, "The model must inspect or read the module before answering.");
                    Assert.AreEqual(1, calls.Length, "Exactly one serial inspection per response; no parallel or repeated dispatch.");
                    var call = calls[0] as IDictionary<string, object>;
                    Assert.IsNotNull(call);
                    var function = call["function"] as IDictionary<string, object>;
                    Assert.IsNotNull(function);
                    string name = Convert.ToString(function["name"]);
                    string argumentsJson = Convert.ToString(function["arguments"]);
                    if (name == "invoke_tool")
                    {
                        var gateway = json.DeserializeObject(argumentsJson) as IDictionary<string, object>;
                        Assert.IsNotNull(gateway); Assert.AreEqual(2, gateway.Count);
                        Assert.IsInstanceOfType(gateway["ToolName"], typeof(string));
                        Assert.IsInstanceOfType(gateway["ArgumentsJson"], typeof(string));
                        name = (string)gateway["ToolName"];
                        Assert.IsFalse(LlmVbeTools.IsCatalogTool(name), "Recursive gateway calls are forbidden.");
                        argumentsJson = (string)gateway["ArgumentsJson"];
                    }
                    Assert.IsTrue(new[] { "status", "list_projects", "list_modules", "discover_tools", "read_module" }.Contains(name),
                        "Only the declared read-only synthetic inspections are allowed.");
                    Assert.IsTrue(inspected.Add(name), "An inspection must not be repeated.");
                    var arguments = json.DeserializeObject(argumentsJson) as IDictionary<string, object>;
                    Assert.IsNotNull(arguments);
                    string result;
                    switch (name)
                    {
                        case "status":
                            Assert.AreEqual(0, arguments.Count);
                            result = json.Serialize(Response.Success(new { Version = "0.1.0", Connected = true,
                                AssemblyPath = "Synthetic", AssemblyModuleVersionId = typeof(LlmChatClient).Module.ModuleVersionId.ToString("D"),
                                HostProcessId = 0, ProcessBitness = 64 }));
                            break;
                        case "list_projects":
                            Assert.AreEqual(0, arguments.Count);
                            result = json.Serialize(Response.Success(new[] { new { Name = project, FileName = project,
                                HostPath = project, HostPathError = (string)null, Mode = 2,
                                FileNameErrorHResult = (int?)null, FileNameErrorType = (string)null } }));
                            break;
                        case "list_modules":
                            Assert.AreEqual(1, arguments.Count);
                            Assert.AreEqual(project, arguments["Project"]);
                            result = json.Serialize(Response.Success(new[] { new { Name = "Q028Marker", Type = 1, Lines = 2 } }));
                            break;
                        case "discover_tools":
                            Assert.AreEqual(1, arguments.Count);
                            Assert.IsTrue(new[] { "code", "environment", "all" }.Contains(Convert.ToString(arguments["Family"])));
                            // This production branch only builds schemas; Execute remains forbidden.
                            result = await catalog.InvokeCatalogAsync(name, argumentsJson);
                            break;
                        default:
                            Assert.AreEqual(2, arguments.Count);
                            Assert.AreEqual(project, arguments["Project"]);
                            Assert.AreEqual("Q028Marker", arguments["Module"]);
                            moduleRead = true;
                            result = json.Serialize(Response.Success(new { Project = project, Module = "Q028Marker", Code = code, Sha256 = EditorDocument.Hash(code) }));
                            break;
                    }
                    messages.Add(response);
                    messages.Add(new { role = "tool", tool_call_id = Convert.ToString(call["id"]), content = result });
                }
                Assert.Fail("The bounded inspection/read/answer sequence did not complete.");
            }
        }

        /// <summary>Mirrors the context wrapper applied to an HTTP discussion request.</summary>
        private static string Request(string text, string host, string project) => "<vbe-encoding-context>\n" + LlmVbeContext.EncodingInstructions +
            "\n</vbe-encoding-context>\n\n" + UiText.Get("Mode for this request: ") + ChatMode.Discussion +
            UiText.Get(". Analysis only; no edits or macro execution.\n") + LlmVbeContext.WithProjectScope(text, host, project);
    }
}
