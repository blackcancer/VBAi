using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Web.Script.Serialization;

namespace VBAi.Tests.Unit
{
    [TestClass]
    public sealed class CodexAgentActivityTests
    {
        [TestMethod]
        public void ToolCaptionsUseLocalizedActionsAndOnlyDeclaredPublicTargets()
        {
            using (var scope = new VBAi.Tests.Infrastructure.LocalizationScope("fr-FR"))
            {
                var item = new Dictionary<string, object>
                {
                    ["id"] = "read", ["type"] = "dynamicToolCall", ["tool"] = "read_module",
                    ["arguments"] = new Dictionary<string, object>
                    {
                        ["Project"] = "Book", ["Module"] = "Module1", ["ApiKey"] = "SECRET", ["Text"] = "PRIVATE CODE"
                    }
                };
                var activity = CodexAgentActivity.FromItem(item, false);
                Assert.AreEqual("Lecture du code VBA · Book · Module1", activity.Title);
                StringAssert.StartsWith(activity.Detail, "read_module");
                Assert.IsFalse(activity.Detail.Contains("SECRET") || activity.Title.Contains("PRIVATE"));
                item["tool"] = "invoke_monaco";
                Assert.AreEqual("Travail dans l’éditeur · Book · Module1", CodexAgentActivity.FromItem(item, true).Title);
                item["tool"] = "unknown_private_tool";
                Assert.AreEqual(UiText.Get("Using a tool") + " · Book · Module1", CodexAgentActivity.FromItem(item, true).Title);
                var arguments = (Dictionary<string, object>)item["arguments"];
                arguments["Module"] = "Module1\nInjected heading";
                arguments["Path"] = new Dictionary<string, object> { ["secret"] = "SECRET" };
                Assert.IsFalse(CodexAgentActivity.FromItem(item, true).Title.Contains("Injected") || CodexAgentActivity.FromItem(item, true).Detail.Contains("SECRET"));
            }
        }

        [TestMethod]
        public void NativeActionsRetainTheirTitlesTargetsDurationAndActualOutcome()
        {
            foreach (var kind in new[] { "commandExecution", "fileChange", "dynamicToolCall", "mcpToolCall", "webSearch", "imageView", "collabToolCall" })
            {
                var item = new Dictionary<string, object>
                {
                    ["id"] = "action",
                    ["type"] = kind,
                    ["command"] = "dotnet test",
                    ["cwd"] = "C:/test",
                    ["aggregatedOutput"] = "passed",
                    ["server"] = "service",
                    ["tool"] = "read_module",
                    ["query"] = "documentation",
                    ["path"] = "image.png",
                    ["receiverThreadId"] = "worker",
                    ["durationMs"] = 1200,
                    ["arguments"] = new Dictionary<string, object> { ["Project"] = "Book", ["Module"] = "Module1", ["ApiKey"] = "SECRET" },
                    ["changes"] = new object[] { new Dictionary<string, object> { ["path"] = "module.bas" }, null }
                };
                var started = CodexAgentActivity.FromItem(item, false);
                Assert.AreEqual("inProgress", started.Status); Assert.AreEqual(1200L, started.DurationMs); Assert.IsFalse(started.Append);
                Assert.IsFalse(started.Detail.Contains("SECRET"));
                if (kind == "commandExecution") Assert.AreEqual(UiText.Get("Running a command") + " · dotnet test", started.Title);
                if (kind == "webSearch") Assert.AreEqual(UiText.Get("Searching the web") + " · documentation", started.Title);
                if (kind == "imageView") Assert.AreEqual(UiText.Get("Viewing an image") + " · image.png", started.Title);
                var completed = CodexAgentActivity.FromItem(item, true); Assert.AreEqual("completed", completed.Status);
                item["success"] = false; Assert.AreEqual("failed", CodexAgentActivity.FromItem(item, true).Status);
                item.Remove("success"); item["status"] = "declined"; Assert.AreEqual("declined", CodexAgentActivity.FromItem(item, true).Status);
                if (kind == "commandExecution") { item["exitCode"] = 1; Assert.AreEqual("failed", CodexAgentActivity.FromItem(item, true).Status); }
            }
            Assert.IsNull(CodexAgentActivity.FromItem(null, true));
            Assert.IsNull(CodexAgentActivity.FromItem(new Dictionary<string, object> { ["type"] = "reasoning", ["content"] = "PRIVATE" }, true));
            Assert.AreEqual(16384, CodexAgentActivity.Limit(new string('x', 20000)).Length);
            Assert.AreEqual("", CodexAgentActivity.Limit(null));
        }

        [TestMethod]
        public void DetailedStepsRoundTripInSessionAndLegacyEntriesRemainReadable()
        {
            var json = new JavaScriptSerializer();
            var entry = new ChatEntry
            {
                Speaker = "Outil",
                Text = "target",
                Activity = new CodexAgentActivity
                {
                    Id = "tool",
                    Kind = "dynamicToolCall",
                    Title = "read_module",
                    Detail = "Book · Module1",
                    Status = "completed",
                    DurationMs = 42
                }
            };
            var restored = json.Deserialize<ChatEntry>(json.Serialize(entry));
            Assert.AreEqual("Book · Module1", restored.Activity.Detail); Assert.AreEqual("completed", restored.Activity.Status); Assert.AreEqual(42L, restored.Activity.DurationMs);
            Assert.IsNull(json.Deserialize<ChatEntry>("{\"Speaker\":\"Outil\",\"Text\":\"Legacy\"}").Activity);
        }
    }
}

namespace VBAi.Tests.Unit
{
    [Microsoft.VisualStudio.TestTools.UnitTesting.TestClass, Microsoft.VisualStudio.TestTools.UnitTesting.TestCategory("Unit")]
    public sealed class CodexAgentActivityBoundaryTests
    {
        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public void FileChangeWithoutAnArrayRetainsItsIdentityAndActualOutcome()
        {
            foreach (bool missing in new[] { true, false })
            {
                var item = new System.Collections.Generic.Dictionary<string, object> { ["type"] = "fileChange", ["id"] = "files", ["status"] = "declined" };
                if (!missing) item["changes"] = "unstructured";
                var activity = CodexAgentActivity.FromItem(item, true);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("files", activity.Id);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(UiText.Get("Files"), activity.Title);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("", activity.Detail);
                Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("declined", activity.Status);
            }
        }
    }
}