using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace VBAi.Tests.Unit
{
    public sealed partial class LlmVbeToolsBoundaryTests
    {
        [TestMethod]
        public async Task GroupedOptionToolPreservesScalarValuesAndUsesOneNativeDispatch()
        {
            var fixture = new ToolFixture(); int calls = 0;
            fixture.Tools.Native.SetVbeOption = request =>
            {
                calls++; Assert.IsNull(request.Pane); Assert.IsNull(request.Value);
                Assert.AreEqual(3, request.OptionChanges.Length);
                Assert.AreEqual(false, request.OptionChanges[0].Value);
                Assert.AreEqual(4, request.OptionChanges[1].Value);
                Assert.AreEqual("NativeIndex:1", request.OptionChanges[2].Value);
                Assert.AreEqual("Normal Text", request.OptionChanges[2].Query);
                return new { CommitRequested = true, DialogClosed = true };
            };
            Success(await fixture.Tools.InvokeAsync("set_vbe_option", Json.Serialize(new
            {
                ExpectedOptionsVersion = new string('a', 64),
                OptionChanges = new object[] {
                    new { Pane = "Editor", Property = "Auto Syntax Check", Value = false },
                    new { Pane = "Editor", Property = "Tab Width", Value = 4 },
                    new { Pane = "Editor Format", Property = "Foreground", Value = "NativeIndex:1", Query = "Normal Text" } }
            })), "grouped options");
            Assert.AreEqual(1, calls);
        }

        [TestMethod]
        public async Task MalformedOptionGroupsAndMixedLegacyFieldsNeverReachTheNativeBoundary()
        {
            var fixture = new ToolFixture(); int calls = 0;
            fixture.Tools.Native.SetVbeOption = request => { calls++; return new { }; };
            var valid = new Dictionary<string, object> { ["Pane"] = "Editor", ["Property"] = "Auto Syntax Check", ["Value"] = false };
            var badEntries = new List<object> { null, "bad", new { Pane = "Editor", Property = "X" },
                new { Pane = "", Property = "X", Value = false }, new { Pane = "Editor", Property = "X\nY", Value = false },
                new { Pane = "Editor", Property = "X", Value = new { Nested = true } },
                new { Pane = "Editor", Property = "X", Value = false, Query = 4 },
                new { Pane = "Editor", Property = "X", Value = false, Query = "" },
                new { Pane = "Editor", Property = "X", Value = false, Extra = true } };
            foreach (object batch in new object[] { null, "bad", new object[0], Enumerable.Repeat((object)valid, 65).ToArray() }
                .Concat(badEntries.Select(entry => (object)new[] { entry })))
                Failed(await fixture.Tools.InvokeAsync("set_vbe_option", Json.Serialize(new { ExpectedOptionsVersion = "revision", OptionChanges = batch })), "malformed batch");
            foreach (string field in new[] { "Pane", "Property", "Value", "Query" })
            {
                var args = new Dictionary<string, object> { ["ExpectedOptionsVersion"] = "revision", ["OptionChanges"] = new[] { valid }, [field] = field == "Value" ? (object)false : "legacy" };
                Failed(await fixture.Tools.InvokeAsync("set_vbe_option", Json.Serialize(args)), "mixed " + field);
            }
            foreach (string field in new[] { "Pane", "Property", "Value" })
            {
                var args = new Dictionary<string, object>(valid) { ["ExpectedOptionsVersion"] = "revision" }; args.Remove(field);
                Failed(await fixture.Tools.InvokeAsync("set_vbe_option", Json.Serialize(args)), "missing " + field);
            }
            Assert.AreEqual(0, calls);
        }

        [TestMethod]
        public async Task GroupedOptionsKeepReadOnlyAndPlanningGuards()
        {
            var fixture = new ToolFixture(); int calls = 0;
            fixture.Tools.Native.SetVbeOption = request => { calls++; return new { }; };
            string args = Json.Serialize(new { ExpectedOptionsVersion = "revision", OptionChanges = new[] { new { Pane = "Editor", Property = "Auto Syntax Check", Value = false } } });
            fixture.Settings.VbeEditApproval = "ReadOnly";
            Failed(await fixture.Tools.InvokeAsync("set_vbe_option", args), "read only");
            fixture.Settings.VbeEditApproval = "Automatic"; fixture.Tools.Mode = ChatMode.Plan;
            Failed(await fixture.Tools.InvokeAsync("set_vbe_option", args), "planning");
            Assert.AreEqual(0, calls);
        }
    }
}
