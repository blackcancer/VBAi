using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class LlmVbeToolsProjectGeneralTests
    {
        private static readonly JavaScriptSerializer Json = new JavaScriptSerializer();
        private const string Version = "project-version";
        private static string Arguments(bool write = false, object value = null, string property = "HelpContextID", string project = "P")
        {
            var fields = new Dictionary<string, object> { ["Project"] = project, ["ExpectedProjectVersion"] = Version,
                ["ExpectedMode"] = 2, ["ControlCaption"] = "Propriétés de P..." };
            if (write) { fields["Property"] = property; fields["Value"] = value ?? 321; fields["ExpectedOptionsVersion"] = "native-options"; }
            return Json.Serialize(fields);
        }
        private static LlmVbeTools Tools(LlmSettings settings = null)
        {
            var tools = new LlmVbeTools(null, null, settings ?? new LlmSettings { VbeEditApproval = "Automatic" }) { BoundProject = "P" };
            tools.Execute = request => Response.Success(new object[0]);
            return tools;
        }
        private sealed class SaveQueueContext : SynchronizationContext
        {
            private readonly Queue<Action> pending = new Queue<Action>();
            private readonly int ownerThread = Thread.CurrentThread.ManagedThreadId;
            public override void Post(SendOrPostCallback callback, object state) => pending.Enqueue(() => callback(state));
            internal void Drain()
            {
                Assert.AreEqual(ownerThread, Thread.CurrentThread.ManagedThreadId);
                int dispatched = 0;
                while (pending.Count > 0)
                {
                    Assert.IsTrue(dispatched++ < 100, "Injected General continuations exceeded the bounded queue.");
                    pending.Dequeue()();
                }
            }
        }
        private static Task<string> Invoke(LlmVbeTools tools, bool write, string arguments, bool catalog = false)
        {
            var previous = SynchronizationContext.Current;
            var context = new SaveQueueContext();
            SynchronizationContext.SetSynchronizationContext(context);
            try
            {
                var pending = tools.InvokeAsync(catalog ? "invoke_tool" : write ? "set_project_general" : "read_project_general",
                    catalog ? Json.Serialize(new { ToolName = write ? "set_project_general" : "read_project_general", ArgumentsJson = arguments }) : arguments);
                context.Drain();
                Assert.IsTrue(pending.IsCompleted, "The injected General boundary must settle after bounded dispatch.");
                Assert.AreSame(context, SynchronizationContext.Current);
                return Task.FromResult(pending.GetAwaiter().GetResult());
            }
            finally { SynchronizationContext.SetSynchronizationContext(previous); }
        }
        private static IDictionary<string, object> Data(string response) => (IDictionary<string, object>)((IDictionary<string, object>)Json.DeserializeObject(response))["Data"];

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void GeneralCommandsRequireAsyncBeforeAnyHostEntry(bool write)
        {
            var tools = Tools(); int entries = 0;
            tools.ProjectGeneralNative = (request, change) => { entries++; return Task.FromResult<object>(new { Available = true }); };
            var response = Json.Deserialize<Response>(tools.Invoke(write ? "set_project_general" : "read_project_general", Arguments(write)));
            Assert.IsFalse(response.Ok); StringAssert.Contains(response.Error, "InvokeAsync"); Assert.AreEqual(0, entries);
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public async Task ReadGeneralPreservesUnicodeAndUsesSameCatalogPipeline(bool catalog)
        {
            var tools = Tools(new LlmSettings { VbeEditApproval = "ReadOnly" }); tools.Mode = ChatMode.Discussion;
            string path = @"C:\資料\㩃-é-😀.chm"; Request captured = null; int entries = 0;
            tools.ProjectGeneralNative = (request, write) => {
                captured = request; entries++; Assert.IsFalse(write); Assert.AreEqual("read_project_general", request.Command);
                request.RevalidateProjectPropertyAuthorization(true); request.RevalidateProjectPropertyAuthorization(false);
                return Task.FromResult<object>(new { Available = true, HelpFile = path, HelpContextText = "321", OptionsVersion = "native-options", DialogClosed = true });
            };
            string result = await Invoke(tools, false, Arguments(), catalog);
            Assert.IsTrue(Json.Deserialize<Response>(result).Ok, result); Assert.AreEqual(path, Data(result)["HelpFile"]);
            Assert.AreEqual(1, entries); Assert.IsNull(captured.RevalidateProjectPropertyAuthorization);
        }

        [DataTestMethod, DataRow("Automatic", false), DataRow("AskEachTime", false), DataRow("AskEachTime", true)]
        public async Task GeneralWriteRetainsSingleOriginalApprovalThroughAwait(string policy, bool catalog)
        {
            var settings = new LlmSettings { VbeEditApproval = policy }; var tools = Tools(settings);
            int approvals = 0, entries = 0, writes = 0; Request captured = null;
            int ownerThread = Thread.CurrentThread.ManagedThreadId; SynchronizationContext approvalContext = null;
            tools.ShowApproval = (dialog, window) => {
                Assert.AreEqual(ownerThread, Thread.CurrentThread.ManagedThreadId);
                Assert.IsInstanceOfType(SynchronizationContext.Current, typeof(SaveQueueContext));
                approvalContext = SynchronizationContext.Current; approvals++; return DialogResult.Yes;
            };
            tools.ProjectGeneralNative = async (request, write) => {
                captured = request; entries++; Assert.IsTrue(write);
                Assert.AreEqual("HelpContextID", request.Property); Assert.AreEqual(321, request.Value);
                var context = SynchronizationContext.Current;
                Assert.IsInstanceOfType(context, typeof(SaveQueueContext));
                Assert.AreEqual(ownerThread, Thread.CurrentThread.ManagedThreadId);
                if (policy == "AskEachTime") Assert.AreSame(approvalContext, context);
                await Task.Yield();
                Assert.AreEqual(ownerThread, Thread.CurrentThread.ManagedThreadId);
                Assert.AreSame(context, SynchronizationContext.Current);
                request.RevalidateProjectPropertyAuthorization(true); request.RevalidateProjectPropertyAuthorization(false); writes++;
                return new { Available = true, MutationInvoked = true, DialogClosed = true };
            };
            string result = await Invoke(tools, true, Arguments(true), catalog);
            Assert.IsTrue(Json.Deserialize<Response>(result).Ok, result);
            Assert.AreEqual(policy == "AskEachTime" ? 1 : 0, approvals); Assert.AreEqual(1, entries); Assert.AreEqual(1, writes);
            Assert.IsNull(captured.RevalidateProjectPropertyAuthorization);
        }

        [DataTestMethod]
        [DataRow("policy", false)][DataRow("binding", false)][DataRow("mode", false)][DataRow("scope", false)]
        [DataRow("cached", false)][DataRow("policy", true)][DataRow("binding", true)]
        public async Task GeneralRevocationWhilePendingBlocksFinalMutationAndClearsRuntimeAuthorization(string changed, bool catalog)
        {
            var settings = new LlmSettings { VbeEditApproval = "Automatic" }; var tools = Tools(settings);
            int writes = 0; bool scope = true, cached = true; Request captured = null;
            tools.ValidateScope = () => { if (!scope) throw new InvalidOperationException("Scope revoked"); };
            tools.ValidateCachedScope = () => { if (!cached) throw new InvalidOperationException("Cached scope revoked"); };
            tools.ProjectGeneralNative = async (request, write) => {
                captured = request; await Task.Yield();
                if (changed == "policy") settings.VbeEditApproval = "ReadOnly";
                if (changed == "binding") tools.BoundProject = "Other";
                if (changed == "mode") tools.Mode = ChatMode.Plan;
                if (changed == "scope") scope = false;
                if (changed == "cached") cached = false;
                request.RevalidateProjectPropertyAuthorization(true); request.RevalidateProjectPropertyAuthorization(false); writes++;
                return new { Available = true };
            };
            Assert.IsFalse(Json.Deserialize<Response>(await Invoke(tools, true, Arguments(true), catalog)).Ok);
            Assert.AreEqual(0, writes); Assert.IsNotNull(captured); Assert.IsNull(captured.RevalidateProjectPropertyAuthorization);
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public async Task RevokedReadScopeSuppressesNativeMetadataAfterCompletion(bool catalog)
        {
            var tools = Tools(); bool allowed = true; Request captured = null;
            tools.ValidateScope = () => { if (!allowed) throw new InvalidOperationException("Scope revoked"); };
            tools.ProjectGeneralNative = async (request, write) => {
                captured = request; await Task.Yield(); allowed = false;
                return new { Available = true, HelpFile = "private-path-must-not-leak", DialogClosed = true };
            };
            string result = await Invoke(tools, false, Arguments(), catalog);
            Assert.IsFalse(Json.Deserialize<Response>(result).Ok); Assert.IsFalse(result.Contains("private-path-must-not-leak"));
            Assert.IsNull(captured.RevalidateProjectPropertyAuthorization);
        }

        [DataTestMethod, DataRow(true), DataRow(false)]
        public async Task RevokedWriteResultRetainsMutationOrUnsettledDialogUncertaintyWithoutDisclosingMetadata(bool mutationInvoked)
        {
            var settings = new LlmSettings { VbeEditApproval = "Automatic" }; var tools = Tools(settings);
            tools.ProjectGeneralNative = (request, write) => {
                settings.VbeEditApproval = "ReadOnly";
                return Task.FromResult<object>(new { MutationInvoked = mutationInvoked, Uncertain = true, HelpFile = "private-path-must-not-leak" });
            };
            string result = await Invoke(tools, true, Arguments(true));
            Assert.IsTrue(Json.Deserialize<Response>(result).Ok); Assert.AreEqual(true, Data(result)["Uncertain"]);
            Assert.AreEqual(mutationInvoked, Data(result)["MutationInvoked"]); Assert.AreEqual(false, Data(result)["RetryAllowed"]);
            Assert.IsFalse(result.Contains("private-path-must-not-leak"));
        }

        [TestMethod]
        public async Task CachedReadAuthorizationDoesNotPerformHostAliasReadsAndDetectsRevokedGrants()
        {
            var tools = Tools(); tools.SetReadAccess(new[] { "Other" }, false); int aliasReads = 0, scopeReads = 0;
            tools.Execute = request => { aliasReads++; return Response.Success(new object[0]); };
            tools.ValidateScope = () => scopeReads++;
            tools.ProjectGeneralNative = (request, write) => {
                int before = scopeReads; request.RevalidateProjectPropertyAuthorization(false);
                Assert.AreEqual(before, scopeReads); Assert.AreEqual(0, aliasReads);
                tools.SetReadAccess(new string[0], false);
                Assert.ThrowsException<InvalidOperationException>(() => request.RevalidateProjectPropertyAuthorization(false));
                tools.SetReadAccess(new[] { "Other" }, false);
                return Task.FromResult<object>(new { Available = false });
            };
            await Invoke(tools, false, Arguments(project: "Other"));
            Assert.AreEqual(0, aliasReads);
        }

        [DataTestMethod, DataRow(true), DataRow(false)]
        public async Task HelpFileWritePassesExactUnicodeAndNeverUsesSynchronousComSetter(bool catalog)
        {
            var tools = Tools(); string path = @"C:\資料\㩃-é-😀.chm"; int writes = 0;
            tools.Execute = request => throw new InvalidOperationException("Legacy COM setter must not be reached");
            tools.ProjectGeneralNative = (request, write) => { Assert.IsTrue(write); Assert.AreEqual(path, request.Value); writes++; return Task.FromResult<object>(new { Available = true }); };
            string result = await Invoke(tools, true, Arguments(true, path, "HelpFile"), catalog);
            Assert.IsTrue(Json.Deserialize<Response>(result).Ok, result); Assert.AreEqual(1, writes);
        }

        [DataTestMethod]
        [DataRow("HelpContextID", "321")][DataRow("HelpContextID", true)][DataRow("HelpContextID", 1.5)]
        [DataRow("HelpContextID", -1)][DataRow("HelpContextID", 2147483648L)]
        [DataRow("Name", "NewName")][DataRow("HelpFile", "relative.chm")][DataRow("HelpFile", "C:relative.chm")]
        [DataRow("HelpFile", "C:\\file.hlp")][DataRow("HelpFile", 1)]
        public async Task InvalidGeneralFieldOrValueRefusesBeforeNativeEntry(string property, object value)
        {
            var tools = Tools(); int entries = 0;
            tools.ProjectGeneralNative = (request, write) => { entries++; return Task.FromResult<object>(null); };
            Assert.IsFalse(Json.Deserialize<Response>(await Invoke(tools, true, Arguments(true, value, property))).Ok);
            Assert.AreEqual(0, entries);
        }

        [DataTestMethod, DataRow("ReadOnly"), DataRow("unknown")]
        public async Task DisallowedWritePolicyRefusesBeforeNativeEntry(string policy)
        {
            var tools = Tools(new LlmSettings { VbeEditApproval = policy }); int entries = 0;
            tools.ProjectGeneralNative = (request, write) => { entries++; return Task.FromResult<object>(null); };
            Assert.IsFalse(Json.Deserialize<Response>(await Invoke(tools, true, Arguments(true))).Ok); Assert.AreEqual(0, entries);
        }

        [TestMethod]
        public async Task NativeFailureClearsCallbackAndNeverRetries()
        {
            var tools = Tools(); Request captured = null; int entries = 0;
            tools.ProjectGeneralNative = (request, write) => { captured = request; entries++; throw new InvalidOperationException("Original failure"); };
            var result = Json.Deserialize<Response>(await Invoke(tools, true, Arguments(true)));
            Assert.IsFalse(result.Ok); StringAssert.Contains(result.Error, "Original failure"); Assert.AreEqual(1, entries);
            Assert.IsNull(captured.RevalidateProjectPropertyAuthorization);
        }

        [DataTestMethod, DataRow("mode"), DataRow("version"), DataRow("options"), DataRow("extra"), DataRow("callback")]
        public async Task InvalidGeneralSchemaOrRevisionSelectorRefusesBeforeNativeEntry(string change)
        {
            var tools = Tools(); int entries = 0;
            var fields = (IDictionary<string, object>)Json.DeserializeObject(Arguments(true));
            if (change == "mode") fields["ExpectedMode"] = 1;
            if (change == "version") fields.Remove("ExpectedProjectVersion");
            if (change == "options") fields.Remove("ExpectedOptionsVersion");
            if (change == "extra") fields["Unexpected"] = true;
            if (change == "callback") fields["RevalidateProjectPropertyAuthorization"] = true;
            tools.ProjectGeneralNative = (request, write) => { entries++; return Task.FromResult<object>(null); };
            Assert.IsFalse(Json.Deserialize<Response>(await Invoke(tools, true, Json.Serialize(fields))).Ok);
            Assert.AreEqual(0, entries);
        }

        [TestMethod]
        public async Task GeneralReadAndWriteCatalogVisibilityRespectsConversationMode()
        {
            var tools = Tools(); tools.Mode = ChatMode.Plan;
            string result = await tools.InvokeAsync("discover_tools", "{\"Family\":\"environment\"}");
            var definitions = (object[])Data(result)["Tools"];
            string names = Json.Serialize(definitions);
            Assert.IsTrue(names.Contains("read_project_general")); Assert.IsFalse(names.Contains("set_project_general"));
            Assert.AreEqual("environment", LlmVbeTools.ToolFamily("read_project_general"));
            Assert.AreEqual("environment", LlmVbeTools.ToolFamily("set_project_general"));
        }
    }
}
