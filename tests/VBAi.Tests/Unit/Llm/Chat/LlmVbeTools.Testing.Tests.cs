using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class LlmVbaTestingBoundaryTests
    {
        private readonly JavaScriptSerializer json = new JavaScriptSerializer();
        private const string SupportText = "Option Explicit\r\n' Reviewed support source\r\n";
        private static readonly string[] Reads = { "discover_vba_tests", "preview_vba_test_support", "vba_test_run_status", "navigate_vba_test", "vba_test_coverage", "show_vba_test_explorer" };
        private static readonly string[] Writes = { "install_vba_test_support", "run_vba_tests", "stop_vba_tests" };

        [TestMethod]
        public async Task TestingFamilyIsReachableAndDiscussionExposesOnlyInspections()
        {
            var tools = Create();
            var all = Reads.Concat(Writes).ToArray();
            foreach (string name in all) Assert.AreEqual("testing", LlmVbeTools.ToolFamily(name));
            var discovered = json.Deserialize<Response>(await tools.InvokeAsync("discover_tools", "{\"Family\":\"testing\"}"));
            Assert.IsTrue(discovered.Ok, discovered.Error);
            var data = Dictionary(discovered.Data);
            CollectionAssert.AreEquivalent(all, ((object[])data["Tools"]).Select(DefinitionName).ToArray());
            CollectionAssert.IsSubsetOf(all, tools.CatalogForProvider().Select(DefinitionName).ToArray());
            tools.Mode = ChatMode.Discussion;
            discovered = json.Deserialize<Response>(await tools.InvokeAsync("discover_tools", "{\"Family\":\"testing\"}"));
            CollectionAssert.AreEquivalent(Reads, ((object[])Dictionary(discovered.Data)["Tools"]).Select(DefinitionName).ToArray());
            Assert.IsFalse(tools.CatalogForProvider().Select(DefinitionName).Any(Writes.Contains));
            tools.Execute = request => throw new InvalidOperationException("Must not dispatch a write");
            foreach (string name in Writes)
            {
                var response = await Gateway(tools, name, Arguments(name));
                Assert.IsFalse(response.Ok, name);
                Assert.IsFalse(response.Error.Contains("Must not dispatch"), name);
            }
        }

        [TestMethod]
        public void TestingSchemaSupportsWholeBatchesAndRequiresModeRevisionAndReviewedSource()
        {
            foreach (string name in Reads.Concat(Writes))
            {
                var definition = Dictionary(LlmVbeTools.Definitions.Select(x => Dictionary(Dictionary(x)["function"])).Single(x => (string)x["name"] == name));
                var parameters = Dictionary(definition["parameters"]);
                CollectionAssert.Contains((object[])parameters["required"], "Project");
                Assert.AreEqual(false, parameters["additionalProperties"]);
                if (name == "run_vba_tests" || name == "install_vba_test_support")
                {
                    CollectionAssert.Contains((object[])parameters["required"], "ExpectedMode");
                    CollectionAssert.Contains((object[])parameters["required"], "ExpectedProjectVersion");
                }
                if (name == "install_vba_test_support") CollectionAssert.Contains((object[])parameters["required"], "Text");
                if (name == "run_vba_tests")
                {
                    var items = Dictionary(Dictionary(parameters["properties"])["Items"]);
                    Assert.AreEqual(10000, items["maxItems"]);
                    Assert.AreEqual(1, items["minItems"]);
                    Assert.AreEqual(true, items["uniqueItems"]);
                }
            }
        }

        [TestMethod]
        public async Task GatewayCannotBypassProjectReadGrantsOrExecutionSharedContext()
        {
            var tools = Create();
            int commands = 0;
            tools.Execute = request =>
            {
                if (request.Command == "list_projects") return Response.Success(new object[0]);
                commands++;
                return Response.Success(new { Private = "authorized" });
            };
            foreach (string name in Reads)
            {
                var arguments = Arguments(name); arguments["Project"] = "B";
                Assert.IsFalse((await Gateway(tools, name, arguments)).Ok, name);
            }
            Assert.AreEqual(0, commands);
            tools.SetReadAccess(new[] { "B" }, false);
            foreach (string name in Reads)
            {
                var arguments = Arguments(name); arguments["Project"] = "B";
                Assert.IsTrue((await Gateway(tools, name, arguments)).Ok, name);
            }
            Assert.AreEqual(Reads.Length, commands);
            foreach (string name in Writes)
            {
                var arguments = Arguments(name); arguments["Project"] = "B";
                Assert.IsFalse((await Gateway(tools, name, arguments)).Ok, name);
            }
            Assert.IsFalse((await Gateway(tools, "run_vba_tests", Arguments("run_vba_tests"))).Ok);
            Assert.IsFalse((await Gateway(tools, "stop_vba_tests", Arguments("stop_vba_tests"))).Ok);
            Assert.AreEqual(Reads.Length, commands);
            tools.SetReadAccess(new string[0], true);
            Assert.IsTrue((await Gateway(tools, "run_vba_tests", Arguments("run_vba_tests"))).Ok);
            Assert.IsTrue((await Gateway(tools, "stop_vba_tests", Arguments("stop_vba_tests"))).Ok);
            Assert.AreEqual(Reads.Length + 2, commands);
        }

        [TestMethod]
        public void ReadOnlyAndUnknownPoliciesBlockMutationsWhileInspectionsRemainAvailable()
        {
            foreach (string policy in new[] { "ReadOnly", "Unknown" })
            {
                var settings = new LlmSettings { VbeEditApproval = policy };
                var tools = Create(settings); tools.SetReadAccess(new string[0], true);
                int calls = 0;
                tools.Execute = request => { calls++; return Response.Success(new { Available = false }); };
                foreach (string name in Writes) Assert.IsFalse(Invoke(tools, name).Ok, name);
                Assert.AreEqual(0, calls);
                foreach (string name in Reads) Assert.IsTrue(Invoke(tools, name).Ok, name);
                Assert.AreEqual(Reads.Length, calls);
            }
        }

        [TestMethod]
        public void EmptyDuplicateOversizedAndInvalidRunSelectionsNeverReachHost()
        {
            var tools = Create(); tools.SetReadAccess(new string[0], true);
            int calls = 0; tools.Execute = request => { calls++; return Response.Success("unexpected"); };
            foreach (object items in new object[] { new string[0], new[] { "" }, new[] { "x", "X" }, new[] { "new\nline" }, new[] { new string('x', 65) }, new object[] { 5 }, Enumerable.Range(0, 10001).Select(x => x.ToString()).ToArray() })
            {
                var arguments = Arguments("run_vba_tests"); arguments["Items"] = items;
                Assert.IsFalse(json.Deserialize<Response>(tools.Invoke("run_vba_tests", json.Serialize(arguments))).Ok);
            }
            foreach (int mode in new[] { 0, 1, 3 })
            {
                var arguments = Arguments("run_vba_tests"); arguments["ExpectedMode"] = mode;
                Assert.IsFalse(json.Deserialize<Response>(tools.Invoke("run_vba_tests", json.Serialize(arguments))).Ok);
            }
            Assert.AreEqual(0, calls);
            var accepted = Arguments("run_vba_tests"); accepted["Items"] = Enumerable.Range(0, 100).Select(x => "test-" + x).ToArray();
            Assert.IsTrue(json.Deserialize<Response>(tools.Invoke("run_vba_tests", json.Serialize(accepted))).Ok);
            Assert.AreEqual(1, calls);
        }

        [TestMethod]
        public void NavigationRequiresOneIdAndStatusRejectsUnknownRenderingMode()
        {
            var tools = Create(); int calls = 0;
            tools.Execute = request => { calls++; return Response.Success("unexpected"); };
            foreach (var ids in new[] { new string[0], new[] { "a", "b" } })
            {
                var arguments = Arguments("navigate_vba_test"); arguments["Items"] = ids;
                Assert.IsFalse(json.Deserialize<Response>(tools.Invoke("navigate_vba_test", json.Serialize(arguments))).Ok);
            }
            var invalid = Arguments("vba_test_run_status"); invalid["Action"] = "execute";
            Assert.IsFalse(json.Deserialize<Response>(tools.Invoke("vba_test_run_status", json.Serialize(invalid))).Ok);
            Assert.AreEqual(0, calls);
        }

        [TestMethod]
        public void InvalidReportAndCoveragePagingNeverReachDispatchAndAcceptedPagingIsPreserved()
        {
            var tools = Create(); int calls = 0;
            tools.Execute = request => { calls++; Assert.AreEqual(15900, request.Offset); Assert.AreEqual(100, request.Limit); return Response.Success("page"); };
            foreach (string name in new[] { "vba_test_run_status", "vba_test_coverage" })
            {
                foreach (var pair in new[] { new[] { -1, 100 }, new[] { 0, -1 }, new[] { 0, 101 } })
                {
                    var invalid = Arguments(name); invalid["Offset"] = pair[0]; invalid["Limit"] = pair[1];
                    Assert.IsFalse(json.Deserialize<Response>(tools.Invoke(name, json.Serialize(invalid))).Ok);
                }
            }
            Assert.AreEqual(0, calls);
            foreach (string name in new[] { "vba_test_run_status", "vba_test_coverage" })
            {
                var valid = Arguments(name); valid["Offset"] = 15900; valid["Limit"] = 100;
                Assert.IsTrue(json.Deserialize<Response>(tools.Invoke(name, json.Serialize(valid))).Ok);
            }
            Assert.AreEqual(2, calls);
        }

        [TestMethod]
        public void SupportInstallRejectsChangedPreviewBeforeApprovalAndDispatch()
        {
            var settings = new LlmSettings { VbeEditApproval = "AskEachTime" };
            var tools = Create(settings); int installs = 0, approvals = 0;
            string source = SupportText, revision = "revision";
            tools.Execute = request =>
            {
                if (request.Command == "preview_vba_test_support") return Response.Success(new { Text = source, ExpectedProjectVersion = revision });
                installs++; return Response.Success(new { Installed = true });
            };
            tools.ShowApproval = (dialog, owner) => { approvals++; return DialogResult.Yes; };
            source = SupportText + "' unreviewed";
            Assert.IsFalse(Invoke(tools, "install_vba_test_support").Ok);
            source = SupportText; revision = "changed";
            Assert.IsFalse(Invoke(tools, "install_vba_test_support").Ok);
            Assert.AreEqual(0, approvals); Assert.AreEqual(0, installs);
            revision = "revision";
            Assert.IsTrue(Invoke(tools, "install_vba_test_support").Ok);
            Assert.AreEqual(1, approvals); Assert.AreEqual(1, installs);
        }

        [TestMethod]
        public void ApprovalReentrancyRechecksModePolicyBindingAndSharedConsent()
        {
            foreach (string change in new[] { "mode", "policy", "binding", "unbound", "privacy" })
            {
                var settings = new LlmSettings { VbeEditApproval = "AskEachTime" };
                var tools = Create(settings); tools.SetReadAccess(new string[0], true);
                int calls = 0; tools.Execute = request => { calls++; return Response.Success("must not invoke"); };
                tools.ShowApproval = (dialog, owner) =>
                {
                    if (change == "mode") tools.Mode = ChatMode.Plan;
                    if (change == "policy") settings.VbeEditApproval = "ReadOnly";
                    if (change == "binding") tools.BoundProject = "B";
                    if (change == "unbound") tools.BoundProject = null;
                    if (change == "privacy") tools.SetReadAccess(new string[0], false);
                    return DialogResult.Yes;
                };
                Assert.IsFalse(Invoke(tools, "run_vba_tests").Ok, change);
                Assert.AreEqual(0, calls, change);
            }
        }

        [TestMethod]
        public void ResultsAreWithheldIfProjectBindingChangesInsideNativeCall()
        {
            var tools = Create();
            tools.Execute = request => { tools.BoundProject = null; return Response.Success(new { Secret = "private results" }); };
            var read = Invoke(tools, "vba_test_run_status");
            Assert.IsFalse(read.Ok); Assert.IsFalse(json.Serialize(read).Contains("private results"));
            tools.BoundProject = "A"; tools.SetReadAccess(new string[0], true);
            var run = Invoke(tools, "run_vba_tests");
            Assert.IsTrue(run.Ok); Assert.IsTrue((bool)Dictionary(run.Data)["OutcomeUnknown"]);
            Assert.IsFalse(json.Serialize(run).Contains("private results"));
        }

        [TestMethod]
        public void DeferredExecutionGuardRetainsApprovalAndRejectsLaterAuthorityChanges()
        {
            foreach (string change in new[] { "mode", "policy", "binding", "unbound", "privacy", "scope" })
            {
                var settings = new LlmSettings { VbeEditApproval = "AskEachTime" };
                var session = new VbeSession(new VbeSessionTests.FakeVbe());
                Action original = () => { };
                session.TestExecutionGuard = original;
                var tools = new LlmVbeTools(session, null, settings) { BoundProject = "A" };
                tools.SetReadAccess(new string[0], true);
                tools.ShowApproval = (dialog, owner) => DialogResult.Yes;
                Action guard = null;
                tools.Execute = request =>
                {
                    guard = session.TestExecutionGuard;
                    Assert.IsNotNull(guard);
                    Assert.AreNotSame(original, guard);
                    guard();
                    return Response.Success(new { Query = "run-id" });
                };
                Assert.IsTrue(Invoke(tools, "run_vba_tests").Ok, change);
                Assert.AreSame(original, session.TestExecutionGuard, change);
                guard(); // The approved guard remains usable after the session field was restored.
                if (change == "mode") tools.Mode = ChatMode.Plan;
                if (change == "policy") settings.VbeEditApproval = "ReadOnly";
                if (change == "binding") tools.BoundProject = "B";
                if (change == "unbound") tools.BoundProject = null;
                if (change == "privacy") tools.SetReadAccess(new string[0], false);
                if (change == "scope") tools.ValidateScope = () => throw new InvalidOperationException("Conversation closed");
                Assert.ThrowsException<InvalidOperationException>(() => guard(), change);
                Assert.AreSame(original, session.TestExecutionGuard, change);
            }
        }

        [TestMethod]
        public void DeferredExecutionGuardIsRestoredAfterNativeDispatchFailure()
        {
            var session = new VbeSession(new VbeSessionTests.FakeVbe());
            Action original = () => { };
            session.TestExecutionGuard = original;
            var tools = new LlmVbeTools(session, null, new LlmSettings { VbeEditApproval = "Automatic" }) { BoundProject = "A" };
            tools.SetReadAccess(new string[0], true);
            tools.Execute = request =>
            {
                Assert.AreNotSame(original, session.TestExecutionGuard);
                throw new InvalidOperationException("Native scheduling failed");
            };
            var response = Invoke(tools, "run_vba_tests");
            Assert.IsFalse(response.Ok);
            StringAssert.Contains(response.Error, "Native scheduling failed");
            Assert.AreSame(original, session.TestExecutionGuard);
        }

        [DataTestMethod]
        [DataRow("ReadOnly")]
        [DataRow("AskEachTime")]
        [DataRow("Unknown")]
        public void AutomaticRunCannotAcquireApprovalAfterItsPolicyChanges(string changedPolicy)
        {
            var session = new VbeSession(new VbeSessionTests.FakeVbe());
            var settings = new LlmSettings { VbeEditApproval = "Automatic" };
            var tools = new LlmVbeTools(session, null, settings) { BoundProject = "A" };
            tools.SetReadAccess(new string[0], true);
            Action guard = null;
            tools.Execute = request => { guard = session.TestExecutionGuard; return Response.Success(new { Query = "run-id" }); };
            Assert.IsTrue(Invoke(tools, "run_vba_tests").Ok);
            Assert.IsNotNull(guard);
            settings.VbeEditApproval = changedPolicy;
            Assert.ThrowsException<InvalidOperationException>(() => guard());
            Assert.IsNull(session.TestExecutionGuard);
        }

        [TestMethod]
        public void MissingRequiredAndUnexpectedFieldsAreRejectedBeforeHostDispatch()
        {
            var tools = Create(); tools.SetReadAccess(new string[0], true);
            int calls = 0; tools.Execute = request => { calls++; return Response.Success("unexpected"); };
            foreach (string name in Reads.Concat(Writes))
            {
                foreach (string required in Arguments(name).Keys.ToArray())
                {
                    var arguments = Arguments(name); arguments.Remove(required);
                    Assert.IsFalse(json.Deserialize<Response>(tools.Invoke(name, json.Serialize(arguments))).Ok, name + "." + required);
                }
                var extra = Arguments(name); extra["OverridePrivacy"] = true;
                Assert.IsFalse(json.Deserialize<Response>(tools.Invoke(name, json.Serialize(extra))).Ok, name);
            }
            Assert.AreEqual(0, calls);
        }

        [TestMethod]
        public void TestingPreparationDefensivelyRejectsInvalidSelectionsModesActionsAndPreviewFields()
        {
            var tools = Create();
            var prepare = typeof(LlmVbeTools).GetMethod("PrepareTestingRequest", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            foreach (var request in new[] {
                new Request { Command = "run_vba_tests", Items = null },
                new Request { Command = "run_vba_tests", Items = new string[0] },
                new Request { Command = "run_vba_tests", Items = new[] { " " } },
                new Request { Command = "run_vba_tests", Items = new[] { "a", "A" } },
                new Request { Command = "navigate_vba_test", Items = new[] { "a", "b" } },
                new Request { Command = "install_vba_test_support", ExpectedMode = 1 },
                new Request { Command = "run_vba_tests", Items = new[] { "a" }, ExpectedMode = 2, Action = "execute" },
                new Request { Command = "vba_test_run_status", Action = "execute" },
                new Request { Command = "install_vba_test_support", ExpectedMode = 2, Text = " " },
                new Request { Command = "install_vba_test_support", ExpectedMode = 2, Text = new string('x', 1024 * 1024 + 1) }
            })
            {
                var error = Assert.ThrowsException<System.Reflection.TargetInvocationException>(() => prepare.Invoke(tools, new object[] { request }));
                Assert.IsInstanceOfType(error.InnerException, typeof(ArgumentException), request.Command);
            }
            var install = new Request { Command = "install_vba_test_support", ExpectedMode = 2, Text = SupportText, ExpectedProjectVersion = "revision", Project = "A" };
            foreach (object preview in new object[] { null, "invalid", new { Text = SupportText }, new { ExpectedProjectVersion = "revision" } })
            {
                tools.Execute = _ => Response.Success(preview);
                var error = Assert.ThrowsException<System.Reflection.TargetInvocationException>(() => prepare.Invoke(tools, new object[] { install }));
                Assert.IsInstanceOfType(error.InnerException, typeof(InvalidOperationException));
            }
            tools.Execute = _ => Response.Failure("Preview unavailable");
            var failure = Assert.ThrowsException<System.Reflection.TargetInvocationException>(() => prepare.Invoke(tools, new object[] { install }));
            Assert.AreEqual("Preview unavailable", failure.InnerException.Message);
            foreach (string action in new[] { null, "", "compact", "human" })
                prepare.Invoke(tools, new object[] { new Request { Command = "vba_test_run_status", Action = action } });
            prepare.Invoke(tools, new object[] { new Request { Command = "run_vba_tests", ExpectedMode = 2, Items = new[] { "a" }, Action = "coverage" } });
        }

        private LlmVbeTools Create(LlmSettings settings = null) => new LlmVbeTools(null, null, settings ?? new LlmSettings { VbeEditApproval = "Automatic" }) { BoundProject = "A" };
        private Response Invoke(LlmVbeTools tools, string name) => json.Deserialize<Response>(tools.Invoke(name, json.Serialize(Arguments(name))));
        private async Task<Response> Gateway(LlmVbeTools tools, string name, Dictionary<string, object> arguments) => json.Deserialize<Response>(await tools.InvokeAsync("invoke_tool", json.Serialize(new { ToolName = name, ArgumentsJson = json.Serialize(arguments) })));
        private Dictionary<string, object> Dictionary(object value) => (Dictionary<string, object>)json.DeserializeObject(json.Serialize(value));
        private string DefinitionName(object value) => (string)Dictionary(Dictionary(value)["function"])["name"];
        private static Dictionary<string, object> Arguments(string name)
        {
            var args = new Dictionary<string, object> { ["Project"] = "A" };
            if (name == "install_vba_test_support" || name == "run_vba_tests" || name == "navigate_vba_test") args["ExpectedProjectVersion"] = "revision";
            if (name == "install_vba_test_support" || name == "run_vba_tests") args["ExpectedMode"] = 2;
            if (name == "install_vba_test_support") args["Text"] = SupportText;
            if (name == "run_vba_tests" || name == "navigate_vba_test") args["Items"] = new[] { "test-id" };
            if (name == "vba_test_run_status" || name == "stop_vba_tests") args["Query"] = "run-id";
            return args;
        }
    }
}
