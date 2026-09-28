using System.Collections.Generic;
using System.Web.Script.Serialization;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    public sealed class CodexAgentActivityTests
    {
        [TestMethod]
        public void NativeActionsRetainTheirTitlesTargetsDurationAndActualOutcome()
        {
            foreach (var kind in new[] { "commandExecution", "fileChange", "dynamicToolCall", "mcpToolCall", "webSearch", "imageView", "collabToolCall" })
            {
                var item = new Dictionary<string, object> { ["id"] = "action", ["type"] = kind, ["command"] = "dotnet test", ["cwd"] = "C:/test", ["aggregatedOutput"] = "passed",
                    ["server"] = "service", ["tool"] = "read_module", ["query"] = "documentation", ["path"] = "image.png", ["receiverThreadId"] = "worker",
                    ["durationMs"] = 1200, ["arguments"] = new Dictionary<string, object> { ["Project"] = "Book", ["Module"] = "Module1", ["ApiKey"] = "SECRET" },
                    ["changes"] = new object[] { new Dictionary<string, object> { ["path"] = "module.bas" }, null } };
                var started = CodexAgentActivity.FromItem(item, false);
                Assert.AreEqual("inProgress", started.Status); Assert.AreEqual(1200L, started.DurationMs); Assert.IsFalse(started.Append);
                Assert.IsFalse(started.Detail.Contains("SECRET"));
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
            var entry = new ChatEntry { Speaker = "Outil", Text = "target", Activity = new CodexAgentActivity {
                Id = "tool", Kind = "dynamicToolCall", Title = "read_module", Detail = "Book · Module1", Status = "completed", DurationMs = 42 } };
            var restored = json.Deserialize<ChatEntry>(json.Serialize(entry));
            Assert.AreEqual("Book · Module1", restored.Activity.Detail); Assert.AreEqual("completed", restored.Activity.Status); Assert.AreEqual(42L, restored.Activity.DurationMs);
            Assert.IsNull(json.Deserialize<ChatEntry>("{\"Speaker\":\"Outil\",\"Text\":\"Legacy\"}").Activity);
        }
    }
}
