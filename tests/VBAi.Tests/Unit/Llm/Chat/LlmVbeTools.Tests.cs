namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using VBAi;
    public sealed partial class LlmVbeToolsBoundaryTests
    {
        [DataTestMethod]
        [DataRow("policy")]
        [DataRow("binding")]
        [DataRow("mode")]
        public void FinalMetadataAuthorizationUsesOnlyCachedGuardsAfterHostScopeRead(string changed)
        {
            var fixture = new ToolFixture(); fixture.Tools.BoundProject = "P"; int scopeReads = 0, writes = 0, cachedReads = 0, beforeFinal = -1; Request captured = null;
            string actualSelection = "P";
            fixture.Tools.ValidateScope = () => scopeReads++;
            fixture.Tools.ValidateCachedScope = () => { cachedReads++; fixture.Tools.BoundProject = actualSelection; };
            fixture.Tools.Execute = request =>
            {
                if (request.Command != "set_project_property") return VBAi.Tests.Infrastructure.VbeToolBoundaryFixture.Execute(request);
                captured = request;
                request.RevalidateProjectPropertyAuthorization(true);
                beforeFinal = scopeReads;
                if (changed == "policy") fixture.Settings.VbeEditApproval = "ReadOnly";
                if (changed == "binding") actualSelection = "Other"; // The UI selector changes while the previous BoundProject remains cached.
                if (changed == "mode") fixture.Tools.Mode = ChatMode.Plan;
                try { request.RevalidateProjectPropertyAuthorization(false); writes++; }
                finally { Assert.AreEqual(beforeFinal, scopeReads, "The final phase must not dispatch another host scope read."); }
                return Response.Success(new { Changed = true });
            };
            Failed(fixture.Tools.InvokeAsync("set_project_property", Json.Serialize(new { Project = "P", Property = "HelpContextID", Value = 321, ExpectedProjectVersion = "version" })).GetAwaiter().GetResult(), "cached guard revoked");
            Assert.AreEqual(0, writes); Assert.IsNull(captured.RevalidateProjectPropertyAuthorization);
            Assert.AreEqual(1, cachedReads);
            Assert.AreEqual(beforeFinal, scopeReads, "Final authorization cannot hide an extra host read inside an error response.");
        }

        [DataTestMethod]
        [DataRow("policy", false)]
        [DataRow("binding", false)]
        [DataRow("mode", false)]
        [DataRow("scope", false)]
        [DataRow("policy", true)]
        [DataRow("binding", true)]
        public void ProjectMetadataAuthorizationRefusesRevokedDirectAndCatalogWrites(string changed, bool catalog)
        {
            var fixture = new ToolFixture(); fixture.Tools.BoundProject = "P"; int writes = 0; bool scope = true, callbackSeen = false; Request captured = null;
            fixture.Tools.ValidateScope = () => { if (!scope) throw new InvalidOperationException("Scope revoked"); };
            fixture.Tools.Execute = request =>
            {
                if (request.Command != "set_project_property") return VBAi.Tests.Infrastructure.VbeToolBoundaryFixture.Execute(request);
                captured = request; Assert.IsNotNull(request.RevalidateProjectPropertyAuthorization);
                callbackSeen = request.RevalidateProjectPropertyAuthorization != null;
                if (changed == "policy") fixture.Settings.VbeEditApproval = "ReadOnly";
                if (changed == "binding") fixture.Tools.BoundProject = "Other";
                if (changed == "mode") fixture.Tools.Mode = ChatMode.Plan;
                if (changed == "scope") scope = false;
                request.RevalidateProjectPropertyAuthorization(true); request.RevalidateProjectPropertyAuthorization(false); writes++; return Response.Success(new { Changed = true });
            };
            string arguments = Json.Serialize(new { Project = "P", Property = "HelpContextID", Value = 321, ExpectedProjectVersion = "version" });
            if (catalog) arguments = Json.Serialize(new { ToolName = "set_project_property", ArgumentsJson = arguments });
            Failed(fixture.Tools.InvokeAsync(catalog ? "invoke_tool" : "set_project_property", arguments).GetAwaiter().GetResult(), "revoked metadata authorization");
            Assert.AreEqual(0, writes); Assert.IsNotNull(captured); Assert.IsNull(captured.RevalidateProjectPropertyAuthorization);
            Assert.IsTrue(callbackSeen, "The refusal must come from the installed runtime guard.");
        }

        [DataTestMethod]
        [DataRow("Automatic", false)]
        [DataRow("AskEachTime", false)]
        [DataRow("Automatic", true)]
        public void ProjectMetadataAuthorizationReusesApprovalAndClearsAfterOriginalDispatch(string policy, bool failed)
        {
            var fixture = new ToolFixture(); fixture.Tools.BoundProject = "P"; fixture.Settings.VbeEditApproval = policy;
            int approvals = 0, writes = 0; Request captured = null;
            fixture.Tools.ShowApproval = (dialog, owner) => { approvals++; return System.Windows.Forms.DialogResult.Yes; };
            fixture.Tools.Execute = request =>
            {
                if (request.Command != "set_project_property") return VBAi.Tests.Infrastructure.VbeToolBoundaryFixture.Execute(request);
                captured = request; Assert.IsNotNull(request.RevalidateProjectPropertyAuthorization);
                request.RevalidateProjectPropertyAuthorization(true); request.RevalidateProjectPropertyAuthorization(false); writes++;
                if (failed) throw new InvalidOperationException("Original setter failed");
                return Response.Success(new { Changed = true });
            };
            string result = fixture.Tools.InvokeAsync("set_project_property", Json.Serialize(new { Project = "P", Property = "HelpContextID", Value = 321, ExpectedProjectVersion = "version" })).GetAwaiter().GetResult();
            if (failed) Failed(result, "original metadata failure"); else Success(result, "metadata approval");
            Assert.AreEqual(1, writes); Assert.AreEqual(policy == "AskEachTime" ? 1 : 0, approvals); Assert.IsNull(captured.RevalidateProjectPropertyAuthorization);
        }

        [TestMethod]
        public void MetadataAuthorizationCannotBeSerializedOrInstalledByJson()
        {
            var request = new Request { Command = "set_project_property" };
            request.RevalidateProjectPropertyAuthorization = validateScope => Assert.Fail("Serialization must not invoke authorization.");
            Assert.IsFalse(Json.Serialize(request).Contains("RevalidateProjectPropertyAuthorization"));
            Assert.IsNull(Json.Deserialize<Request>("{\"Command\":\"set_project_property\",\"RevalidateProjectPropertyAuthorization\":true}").RevalidateProjectPropertyAuthorization);
            Assert.IsNull(typeof(Request).GetField("RevalidateProjectPropertyAuthorization"));
            Assert.IsNull(typeof(Request).GetProperty("RevalidateProjectPropertyAuthorization"));
        }

        [DataTestMethod]
        [DataRow(false, false)]
        [DataRow(true, false)]
        [DataRow(false, true)]
        [DataRow(true, true)]
        public void ScalarFailurePhaseSurvivesDirectAndCatalogResponsesWithoutRetry(bool catalog, bool readback)
        {
            var fixture = new ToolFixture();
            fixture.Tools.BoundProject = "P";
            var error = new System.Runtime.InteropServices.COMException("synthetic original COM error", unchecked((int)0x9CFD3148));
            VbeScalarProperty.AnnotateFailure(error, readback ? VbeScalarProperty.FailurePhase.RetentionReadback : VbeScalarProperty.FailurePhase.SetterInvocation);
            int writes = 0;
            fixture.Tools.Execute = request =>
            {
                if (request.Command != "set_project_property") return VBAi.Tests.Infrastructure.VbeToolBoundaryFixture.Execute(request);
                writes++;
                throw error;
            };
            string args = Json.Serialize(new { Project = "P", Property = "HelpContextID", Value = 321, ExpectedProjectVersion = "value" });
            string tool = "set_project_property";
            if (catalog) { tool = "invoke_tool"; args = Json.Serialize(new { ToolName = "set_project_property", ArgumentsJson = args }); }
            var response = Dict(Json.DeserializeObject(fixture.Tools.InvokeAsync(tool, args).GetAwaiter().GetResult()));
            Assert.AreEqual(false, response["Ok"]);
            Assert.IsNull(response["Data"]);
            StringAssert.Contains((string)response["Error"], readback ? "RetentionReadback" : "SetterInvocation");
            StringAssert.Contains((string)response["Error"], error.Message);
            StringAssert.Contains((string)response["Error"], "0x9CFD3148");
            Assert.AreEqual(1, writes);
            Assert.AreEqual("synthetic original COM error", error.Message);
        }

        private sealed class SaveQueueContext : System.Threading.SynchronizationContext
        {
            internal readonly System.Collections.Generic.Queue<System.Action> Pending = new System.Collections.Generic.Queue<System.Action>();
            public override void Post(System.Threading.SendOrPostCallback callback, object state) { Pending.Enqueue(() => callback(state)); }
            internal void Drain() { int limit = 100; while (Pending.Count > 0 && limit-- > 0) Pending.Dequeue()(); Assert.IsTrue(limit > 0); }
        }
        private static string InvokeSaveContract(ToolFixture fixture, string arguments)
        {
            var previous = System.Threading.SynchronizationContext.Current;
            var context = new SaveQueueContext();
            System.Threading.SynchronizationContext.SetSynchronizationContext(context);
            try
            {
                var pending = fixture.Tools.InvokeAsync("save_host_document", arguments);
                context.Drain();
                Assert.IsTrue(pending.IsCompleted, "The injected save boundary must complete after dispatch.");
                return pending.GetAwaiter().GetResult();
            }
            finally { System.Threading.SynchronizationContext.SetSynchronizationContext(previous); }
        }
        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void SaveHostDocumentDefersDirectAndCatalogDispatchThenRechecksApproval(bool catalog)
        {
            var previous = System.Threading.SynchronizationContext.Current;
            var context = new SaveQueueContext();
            System.Threading.SynchronizationContext.SetSynchronizationContext(context);
            try
            {
                var fixture = new ToolFixture();
                int saves = 0;
                fixture.Tools.SaveHostDocumentNative = request => { saves++; return System.Threading.Tasks.Task.FromResult<object>(new { Saved = true }); };
                string args = Json.Serialize(new { Project = "P", ExpectedProjectVersion = "version", ExpectedHostPath = @"C:\fixture\Owned.swp" });
                string command = catalog ? "invoke_tool" : "save_host_document";
                if (catalog) args = Json.Serialize(new { ToolName = "save_host_document", ArgumentsJson = args });
                var pending = fixture.Tools.InvokeAsync(command, args);
                Assert.IsFalse(pending.IsCompleted); Assert.AreEqual(0, saves);
                fixture.Settings.VbeEditApproval = "ReadOnly";
                context.Drain();
                Failed(pending.GetAwaiter().GetResult(), "policy changed during deferred save");
                Assert.AreEqual(0, saves);
                fixture.Settings.VbeEditApproval = "Automatic";
                pending = fixture.Tools.InvokeAsync(command, args);
                Assert.IsFalse(pending.IsCompleted); Assert.AreEqual(0, saves);
                context.Drain();
                Success(pending.GetAwaiter().GetResult(), "deferred save"); Assert.AreEqual(1, saves);
            }
            finally { System.Threading.SynchronizationContext.SetSynchronizationContext(previous); }
        }

        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void SaveWaitsForNativeCompletionWithoutRepeatingDispatch(bool catalog)
        {
            var previous = System.Threading.SynchronizationContext.Current;
            var context = new SaveQueueContext();
            System.Threading.SynchronizationContext.SetSynchronizationContext(context);
            try
            {
                var fixture = new ToolFixture();
                int saves = 0;
                var completion = new System.Threading.Tasks.TaskCompletionSource<object>();
                fixture.Tools.SaveHostDocumentNative = request => { saves++; return completion.Task; };
                string args = Json.Serialize(new { Project = "P", ExpectedProjectVersion = "version", ExpectedHostPath = @"C:\fixture\Owned.swp" });
                Failed(fixture.Tools.Invoke("save_host_document", args), "synchronous save must refuse before invocation");
                Assert.AreEqual(0, saves);
                string command = catalog ? "invoke_tool" : "save_host_document";
                if (catalog) args = Json.Serialize(new { ToolName = "save_host_document", ArgumentsJson = args });
                var pending = fixture.Tools.InvokeAsync(command, args);
                context.Drain();
                Assert.AreEqual(1, saves); Assert.IsFalse(pending.IsCompleted);
                completion.SetResult(new { SaveInvoked = true, Verified = true, Uncertain = false });
                context.Drain();
                Success(pending.GetAwaiter().GetResult(), "native completion");
                Assert.AreEqual(1, saves);
            }
            finally { System.Threading.SynchronizationContext.SetSynchronizationContext(previous); }
        }

        [DataTestMethod]
        [DataRow("project")]
        [DataRow("mode")]
        [DataRow("policy")]
        public void SaveRechecksGuardsAfterApprovalDialog(string changed)
        {
            var previous = System.Threading.SynchronizationContext.Current;
            var context = new SaveQueueContext();
            System.Threading.SynchronizationContext.SetSynchronizationContext(context);
            try
            {
                var fixture = new ToolFixture();
                fixture.Settings.VbeEditApproval = "AskEachTime";
                fixture.Tools.BoundProject = "P";
                int saves = 0;
                fixture.Tools.SaveHostDocumentNative = request => { saves++; return System.Threading.Tasks.Task.FromResult<object>(new { Saved = true }); };
                fixture.Tools.ShowApproval = (dialog, owner) =>
                {
                    if (changed == "project") fixture.Tools.BoundProject = "Other";
                    if (changed == "mode") fixture.Tools.Mode = ChatMode.Plan;
                    if (changed == "policy") fixture.Settings.VbeEditApproval = "ReadOnly";
                    return System.Windows.Forms.DialogResult.Yes;
                };
                var pending = fixture.Tools.InvokeAsync("save_host_document", Json.Serialize(new
                {
                    Project = "P",
                    ExpectedProjectVersion = "version",
                    ExpectedHostPath = @"C:\fixture\Owned.swp"
                }));
                context.Drain();
                Failed(pending.GetAwaiter().GetResult(), "changed after approval: " + changed);
                Assert.AreEqual(0, saves);
            }
            finally { System.Threading.SynchronizationContext.SetSynchronizationContext(previous); }
        }

        [DataTestMethod]
        [DataRow("Other")]
        [DataRow("")]
        [DataRow(null)]
        public void SaveDoesNotExposeResultAfterProjectScopeChangesWhilePending(string newBinding)
        {
            var previous = System.Threading.SynchronizationContext.Current;
            var context = new SaveQueueContext();
            System.Threading.SynchronizationContext.SetSynchronizationContext(context);
            try
            {
                var fixture = new ToolFixture();
                fixture.Tools.BoundProject = "P";
                int saves = 0;
                var completion = new System.Threading.Tasks.TaskCompletionSource<object>();
                fixture.Tools.SaveHostDocumentNative = request => { saves++; return completion.Task; };
                var pending = fixture.Tools.InvokeAsync("save_host_document", Json.Serialize(new
                {
                    Project = "P",
                    ExpectedProjectVersion = "version",
                    ExpectedHostPath = @"C:\fixture\Owned.swp"
                }));
                context.Drain(); Assert.AreEqual(1, saves);
                fixture.Tools.BoundProject = newBinding;
                completion.SetResult(new { SaveInvoked = true, Verified = true, HostPath = "private-path" });
                context.Drain();
                string result = pending.GetAwaiter().GetResult();
                Success(result, "scope changed after invocation");
                StringAssert.Contains(result, "\"Uncertain\":true");
                StringAssert.Contains(result, "do not retry");
                Assert.IsFalse(result.Contains("private-path")); Assert.AreEqual(1, saves);
            }
            finally { System.Threading.SynchronizationContext.SetSynchronizationContext(previous); }
        }

        [DataTestMethod]
        [DataRow("readOnly", false)]
        [DataRow("readOnly", true)]
        [DataRow("binding", false)]
        [DataRow("binding", true)]
        [DataRow("askEachTime", false)]
        [DataRow("unknownPolicy", false)]
        [DataRow("scope", false)]
        [DataRow("mode", false)]
        public void DeferredSaveAuthorizationBlocksConfirmationAfterRevocation(string revoked, bool catalog)
        {
            var previous = System.Threading.SynchronizationContext.Current;
            var context = new SaveQueueContext();
            System.Threading.SynchronizationContext.SetSynchronizationContext(context);
            try
            {
                var fixture = new ToolFixture(); fixture.Tools.BoundProject = "P";
                bool scopeValid = true; int saves = 0, confirmations = 0, dialogs = 0;
                fixture.Tools.ValidateScope = () => { if (!scopeValid) throw new InvalidOperationException("Scope revoked"); };
                fixture.Tools.ShowApproval = (dialog, owner) => { dialogs++; return System.Windows.Forms.DialogResult.Yes; };
                Request captured = null;
                var ready = new System.Threading.Tasks.TaskCompletionSource<object>();
                fixture.Tools.SaveHostDocumentNative = async request =>
                {
                    captured = request; Assert.IsNotNull(request.RevalidateSaveAuthorization);
                    request.RevalidateSaveAuthorization(); saves++; // Simulates the original ID3 mutation.
                    await ready.Task;
                    try { request.RevalidateSaveAuthorization(); }
                    catch (InvalidOperationException)
                    {
                        return new
                        {
                            SaveInvoked = true,
                            Verified = false,
                            Uncertain = true,
                            Reason = "Authorization was revoked before native confirmation; no replay."
                        };
                    }
                    confirmations++;
                    return new { SaveInvoked = true, Verified = true, Uncertain = false };
                };
                string arguments = Json.Serialize(new { Project = "P", ExpectedProjectVersion = "version", ExpectedHostPath = @"C:\fixture\Owned.accdb" });
                if (catalog) arguments = Json.Serialize(new { ToolName = "save_host_document", ArgumentsJson = arguments });
                var pending = fixture.Tools.InvokeAsync(catalog ? "invoke_tool" : "save_host_document", arguments);
                context.Drain(); Assert.AreEqual(1, saves); Assert.IsFalse(pending.IsCompleted);
                if (revoked == "readOnly") fixture.Settings.VbeEditApproval = "ReadOnly";
                if (revoked == "askEachTime") fixture.Settings.VbeEditApproval = "AskEachTime";
                if (revoked == "unknownPolicy") fixture.Settings.VbeEditApproval = "Other";
                if (revoked == "binding") fixture.Tools.BoundProject = "Other";
                if (revoked == "scope") scopeValid = false;
                if (revoked == "mode") fixture.Tools.Mode = ChatMode.Plan;
                ready.SetResult(null); context.Drain();
                string result = pending.GetAwaiter().GetResult();
                Success(result, "original save remains uncertain after " + revoked);
                StringAssert.Contains(result, "\"Uncertain\":true");
                Assert.AreEqual(0, confirmations); Assert.AreEqual(1, saves); Assert.AreEqual(0, dialogs);
                Assert.IsNull(captured.RevalidateSaveAuthorization, "The runtime authorization must end with the native await.");
            }
            finally { System.Threading.SynchronizationContext.SetSynchronizationContext(previous); }
        }

        [DataTestMethod]
        [DataRow("Automatic")]
        [DataRow("AskEachTime")]
        public void DeferredSaveConfirmationReusesOriginalApprovalWithoutAnotherDialog(string policy)
        {
            var fixture = new ToolFixture(); fixture.Tools.BoundProject = "P"; fixture.Settings.VbeEditApproval = policy;
            int approvals = 0, checks = 0, confirmations = 0;
            Request captured = null;
            fixture.Tools.ShowApproval = (dialog, owner) => { approvals++; return System.Windows.Forms.DialogResult.Yes; };
            fixture.Tools.SaveHostDocumentNative = request =>
            {
                captured = request; Assert.IsNotNull(request.RevalidateSaveAuthorization);
                request.RevalidateSaveAuthorization(); checks++;
                request.RevalidateSaveAuthorization(); checks++; confirmations++;
                return System.Threading.Tasks.Task.FromResult<object>(new { SaveInvoked = true, Verified = true });
            };
            Success(InvokeSaveContract(fixture, Json.Serialize(new
            {
                Project = "P",
                ExpectedProjectVersion = "version",
                ExpectedHostPath = @"C:\fixture\Owned.accdb"
            })), "authorized native confirmation");
            Assert.AreEqual(policy == "AskEachTime" ? 1 : 0, approvals);
            Assert.AreEqual(2, checks); Assert.AreEqual(1, confirmations); Assert.IsNull(captured.RevalidateSaveAuthorization);
        }

        [TestMethod]
        public void DeferredSaveAuthorizationClearsOnNativeFailure()
        {
            var fixture = new ToolFixture(); fixture.Tools.BoundProject = "P"; Request captured = null; int saves = 0;
            fixture.Tools.SaveHostDocumentNative = request =>
            {
                captured = request; Assert.IsNotNull(request.RevalidateSaveAuthorization);
                request.RevalidateSaveAuthorization(); saves++;
                return System.Threading.Tasks.Task.FromException<object>(new InvalidOperationException("Original native failure"));
            };
            Failed(InvokeSaveContract(fixture, Json.Serialize(new
            {
                Project = "P",
                ExpectedProjectVersion = "version",
                ExpectedHostPath = @"C:\fixture\Owned.accdb"
            })), "native failure");
            Assert.AreEqual(1, saves); Assert.IsNull(captured.RevalidateSaveAuthorization);
        }

        [TestMethod]
        public void SaveRuntimeAuthorizationIsNotPartOfTheJsonProtocol()
        {
            var request = new Request { Command = "save_host_document", Project = "P" };
            request.RevalidateSaveAuthorization = () => Assert.Fail("Serialization cannot invoke a runtime authorization.");
            string serialized = Json.Serialize(request);
            Assert.IsFalse(serialized.Contains("RevalidateSaveAuthorization"));
            var parsed = Json.Deserialize<Request>("{\"Command\":\"save_host_document\",\"RevalidateSaveAuthorization\":true}");
            Assert.IsNull(parsed.RevalidateSaveAuthorization, "JSON cannot install a runtime delegate.");
            Assert.IsNull(typeof(Request).GetField("RevalidateSaveAuthorization"));
            Assert.IsNull(typeof(Request).GetProperty("RevalidateSaveAuthorization"));
        }

        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void SaveRuntimeAuthorizationCannotBeSuppliedAsToolArgument(bool catalog)
        {
            var fixture = new ToolFixture(); fixture.Tools.BoundProject = "P"; int saves = 0;
            fixture.Tools.SaveHostDocumentNative = request => { saves++; return System.Threading.Tasks.Task.FromResult<object>(new { Saved = true }); };
            string arguments = Json.Serialize(new
            {
                Project = "P",
                ExpectedProjectVersion = "version",
                ExpectedHostPath = @"C:\fixture\Owned.accdb",
                RevalidateSaveAuthorization = true
            });
            var previous = System.Threading.SynchronizationContext.Current;
            var context = new SaveQueueContext(); System.Threading.SynchronizationContext.SetSynchronizationContext(context);
            try
            {
                if (catalog) arguments = Json.Serialize(new { ToolName = "save_host_document", ArgumentsJson = arguments });
                var pending = fixture.Tools.InvokeAsync(catalog ? "invoke_tool" : "save_host_document", arguments);
                context.Drain(); Failed(pending.GetAwaiter().GetResult(), "runtime authorization is not a wire argument");
                Assert.AreEqual(0, saves);
            }
            finally { System.Threading.SynchronizationContext.SetSynchronizationContext(previous); }
        }

        [DataTestMethod]
        [DataRow("project")]
        [DataRow("mode")]
        [DataRow("policy")]
        public async System.Threading.Tasks.Task ImmediateRechecksContextAndPolicyAtNativeSubmission(string change)
        {
            var fixture = new ToolFixture();
            var tools = fixture.Tools;
            var state = new VBAi.Tests.Infrastructure.VbeToolMode { Mode = 2 };
            int enters = 0;
            tools.Execute = request => Response.Success(state);
            tools.Native.ExecuteImmediate = (text, submit) =>
            {
                if (change == "project") state.SelectedProject = "Other";
                if (change == "mode") state.Mode = 1;
                if (change == "policy") fixture.Settings.VbeEditApproval = "Ask";
                submit(() => enters++);
                return new { Executed = true };
            };
            Failed(await tools.InvokeAsync("immediate_execute", Json.Serialize(new { Project = "P", ExpectedMode = 2, Text = "? 1" })), change);
            Assert.AreEqual(0, enters);
        }

        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public async System.Threading.Tasks.Task ImmediateGatewayRefusesOtherActiveProjectEvenWithSharedContext()
        {
            var tools = new ToolFixture().Tools;
            tools.BoundProject = @"C:\Temp\A.xlsm";
            tools.SetReadAccess(new string[0], true);
            int executions = 0;
            tools.Native.ExecuteImmediate = (text, submit) => { submit(() => executions++); return new { Executed = true }; };
            var state = new VBAi.Tests.Infrastructure.VbeToolMode
            {
                Mode = 2,
                Project = "SameName",
                SelectedProject = "SameName",
                SelectedProjectPath = @"C:\Temp\B.xlsm",
                ActiveModule = "Module1"
            };
            tools.Execute = request => Response.Success(state);
            string arguments = Json.Serialize(new { Project = tools.BoundProject, ExpectedMode = 2, Text = "Debug.Print 1" });
            Failed(await tools.InvokeAsync("immediate_execute", arguments), "other active project");
            Failed(await tools.InvokeAsync("invoke_tool", Json.Serialize(new { ToolName = "immediate_execute", ArgumentsJson = arguments })), "gateway other active project");
            Assert.AreEqual(0, executions);
            state.SelectedProject = null;
            state.SelectedProjectPath = null;
            state.ActiveModule = null;
            Failed(await tools.InvokeAsync("immediate_execute", arguments), "native pane not in requested project");
            Assert.AreEqual(0, executions);
            state.SelectedProject = state.Project;
            state.SelectedProjectPath = tools.BoundProject;
            state.ActiveModule = "Module1";
            Success(await tools.InvokeAsync("immediate_execute", arguments), "exact active project");
            Assert.AreEqual(1, executions);
        }

        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public async System.Threading.Tasks.Task ProtectionAndNavigationAsyncRoutesPreserveSchedulingAndNativeResults()
        {
            foreach (string command in new[] { "read_project_protection", "set_project_protection" })
            {
                var tools = Create();
                string arguments = Json.Serialize(new { Project = "P", ExpectedProjectVersion = "version", ExpectedOptionsVersion = "options", ExpectedMode = 2, ControlCaption = "Properties", Action = "clear" });
                if (command == "read_project_protection") arguments = Json.Serialize(new { Project = "P", ExpectedProjectVersion = "version", ExpectedMode = 2, ControlCaption = "Properties" });
                string captured = null;
                tools.Native.ReadProjectProtection = request => { captured = request.Caption; return new { Available = true }; };
                tools.Native.SetProjectProtection = request => { captured = request.Caption; return new { CommittedRequested = true }; };
                Success(await tools.InvokeAsync(command, arguments), command); Assert.AreEqual("P", captured);
                tools.Execute = request => Response.Failure("schedule declined"); Failed(await tools.InvokeAsync(command, arguments), "declined");
                tools.Native.EnsureNoProjectPropertiesDialog = () => { throw new InvalidOperationException("existing dialog"); };
                Failed(await tools.InvokeAsync(command, arguments), "preexisting dialog");
            }
            foreach (string command in new[] { "read_navigation_surface", "change_navigation_surface" })
            {
                var tools = Create(); string arguments = command == "read_navigation_surface" ? Json.Serialize(new { Pane = "project", Query = "Module", Offset = 0, Limit = 5 }) : Json.Serialize(new { Pane = "project", Control = "token", Action = "select", ExpectedWindowVersion = "version" });
                Success(await tools.InvokeAsync(command, arguments), command);
                tools.Native.ReadNavigationSurface = request => { throw new InvalidOperationException("provider unavailable"); }; tools.Native.ChangeNavigationSurface = request => { throw new InvalidOperationException("provider unavailable"); };
                Failed(await tools.InvokeAsync(command, arguments), "native failure");
                foreach (string invalid in new[] { "null", "[]", "{}", "{\"Pane\":\"project\",\"Extra\":1}", "{\"Pane\":1}", "{\"Pane\":\"project\",\"Offset\":\"0\"}" }) Failed(await tools.InvokeAsync(command, invalid), invalid);
            }
        }
    }

}

namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Collections.Generic;
    using System.Threading;
    using System.Threading.Tasks;
    using VBAi;

    /// <summary>Vérifie les validations asynchrones des outils de débogage VBE.</summary>
    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class LlmVbeAsyncValidationTests
    {
        /// <summary>Refuse les arguments mal formés des outils natifs avant tout accès à l’interface.</summary>
        /// <returns>Tâche terminée lorsque les arguments invalides sont refusés.</returns>
        [TestMethod]
        public async Task NativeDebuggerToolsRejectMalformedArgumentsBeforeUiAccess()
        {
            var settings = new LlmSettings
            {
                VbeEditApproval = "Automatic"
            };
            var tools = new LlmVbeTools(null, null, settings);
            await Failure(tools, "debug_item", "{}", "Pane, Action and PathSegments");
            await Failure(tools, "debug_item", "{\"Pane\":\"locals\",\"Action\":\"expand\",\"PathSegments\":[\"x\"],\"Extra\":1}", "Unexpected debug_item argument");
            await Failure(tools, "immediate_execute", "{}", "Project, ExpectedMode and Text");
            await Failure(tools, "immediate_execute", "{\"Project\":\"P\",\"ExpectedMode\":0,\"Text\":\"Debug.Print 1\"}", "Project, ExpectedMode and Text");
            await Failure(tools, "read_immediate", "{}", "Project and ExpectedMode");
            await Failure(tools, "read_immediate", "{\"Project\":\"P\",\"ExpectedMode\":0}", "Project and ExpectedMode");
            await Failure(tools, "read_immediate", "{\"Project\":\"P\",\"ExpectedMode\":1,\"Text\":\"extra\"}", "no other arguments");
            const string local = "{\"Project\":\"P\",\"Module\":\"M\",\"Procedure\":\"Run\",\"ExpectedSha256\":\"aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa\",\"ExpectedMode\":1";
            await Failure(tools, "inspect_local_scalars", "{}", "Project, Module, Procedure");
            await Failure(tools, "inspect_local_scalars", local.Replace("\"ExpectedMode\":1", "\"ExpectedMode\":2") + "}", "ExpectedMode=1");
            await Failure(tools, "inspect_local_scalars", local.Replace("\"ExpectedSha256\":\"aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa\",", "") + "}", "ExpectedSha256");
            await Failure(tools, "inspect_local_scalars", local + ",\"Limit\":17}", "Limit 0..16");
            await Failure(tools, "inspect_local_scalars", local + ",\"Offset\":-1}", "Offset must be nonnegative");
            await Failure(tools, "inspect_local_scalars", local + ",\"Unexpected\":1}", "no other arguments");
            await Failure(tools, "respond_debug_dialog", "{}", "Exact Diagnostic and Button");
            await Failure(tools, "debug_dialog", "{\"Unexpected\":1}", "debug_dialog has no arguments");
            await Failure(tools, "compile_project", "{}", "Project and ExpectedMode=2");
            await Failure(tools, "compile_project", "{\"Project\":\"P\",\"ExpectedMode\":1}", "Project and ExpectedMode=2");
        }

        /// <summary>Exige une politique automatique pour les évaluations et mutations de débogage.</summary>
        /// <returns>Tâche terminée lorsque les outils refusent la politique insuffisante.</returns>
        [TestMethod]
        public async Task NativeEvaluationAndMutationRequireAutomaticPolicy()
        {
            var settings = new LlmSettings
            {
                VbeEditApproval = "ReadOnly"
            };
            var tools = new LlmVbeTools(null, null, settings);
            await Failure(tools, "quick_watch", "{}", "Automatic VBE edit policy");
            await Failure(tools, "edit_watch", "{}", "Automatic VBE edit policy");
            await Failure(tools, "immediate_execute", "{}", "Automatic VBE edit policy");
            await Failure(tools, "read_immediate", "{}", "Automatic VBE edit policy");
            await Failure(tools, "inspect_local_scalars", "{}", "Automatic VBE edit policy");
            await Failure(tools, "respond_debug_dialog", "{}", "Automatic VBE edit policy");
        }

        /// <summary>Valide le scope et le projet lié avant l’accès à l’hôte pendant un appel asynchrone.</summary>
        /// <returns>Tâche terminée lorsque les limites de scope sont vérifiées.</returns>
        [TestMethod]
        public async Task ScopeAndProjectBindingApplyToAsyncInvocationBeforeHostAccess()
        {
            var tools = new LlmVbeTools(null, null, new LlmSettings { VbeEditApproval = "Automatic" })
            {
                BoundProject = "WorkbookA"
            };
            await Failure(tools, "compile_project", "{\"Project\":\"WorkbookB\",\"ExpectedMode\":2}", "autre projet");
            tools.ValidateScope = () =>
            {
                throw new InvalidOperationException("stale conversation scope");
            };
            await Failure(tools, "debug_item", "{}", "stale conversation scope");
        }

        /// <summary>Retourne le statut en lecture seule depuis une session VBE en mémoire.</summary>
        /// <returns>Tâche terminée après la vérification du statut.</returns>
        [TestMethod]
        public async Task AsyncDispatchReturnsReadOnlyStatusFromInMemorySession()
        {
            var tools = new LlmVbeTools(new VbeSession(new VbeSessionTests.FakeVbe()), null, new LlmSettings());
            var response = Json.Deserialize<Response>(await tools.InvokeAsync("status", "{}"));
            Assert.IsTrue(response.Ok);
            Assert.IsNotNull(response.Data);
        }

        /// <summary>Refuse les formes de requête natives invalides sans ouvrir de dialogue.</summary>
        /// <returns>Tâche terminée après le contrôle des requêtes.</returns>
        [TestMethod]
        public async Task AsyncNativePreflightRejectsWrongShapesWithoutOpeningDialogs()
        {
            var tools = new LlmVbeTools(null, null, new LlmSettings { VbeEditApproval = "Automatic" });
            await Failure(tools, "read_runtime_forms", "[]", "empty argument object");
            await Failure(tools, "read_runtime_forms", "{\"Project\":\"P\"}", "empty argument object");
            await Failure(tools, "debug_windows", "[]", "Tool arguments must be an object");
            await Failure(tools, "debug_windows", "{\"IncludeCallStack\":1}", "must be a boolean");
            await Failure(tools, "debug_windows", "{\"Unexpected\":true}", "Unexpected argument");
            await Failure(tools, "quick_watch", "[]", "Tool arguments must be an object");
            await Failure(tools, "edit_watch", "[]", "Tool arguments must be an object");
            await Failure(tools, "remove_watch", "[]", "Tool arguments must be an object");
        }

        /// <summary>Exige un contexte UI avant d’inspecter le dialogue de compilation natif.</summary>
        /// <returns>Tâche terminée lorsque l’absence de contexte UI est refusée.</returns>
        [TestMethod]
        public async Task CompileRequiresUiContextBeforeNativeDialogInspection()
        {
            var tools = new LlmVbeTools(null, null, new LlmSettings());
            var prior = SynchronizationContext.Current;
            try
            {
                SynchronizationContext.SetSynchronizationContext(null);
                await Failure(tools, "compile_project", "{\"Project\":\"P\",\"ExpectedMode\":2}", "VBE UI context is unavailable");
            }
            finally
            {
                SynchronizationContext.SetSynchronizationContext(prior);
            }
        }

        /// <summary>Refuse l’exécution immédiate si le mode du projet a changé avant l’appel natif.</summary>
        /// <returns>Tâche terminée lorsque l’appel natif n’est pas exécuté.</returns>
        [TestMethod]
        public async Task ImmediateExecuteRejectsModeChangedBeforeNativeExecution()
        {
            var host = new VbeSessionTests.FakeVbe();
            host.VBProjects.Add(new VbeSessionTests.FakeProject { Name = "P", FileName = @"C:\Temp\P.xlsm", Mode = 2 });
            var tools = new LlmVbeTools(new VbeSession(host), null, new LlmSettings { VbeEditApproval = "Automatic" });
            await Failure(tools, "immediate_execute", "{\"Project\":\"P\",\"ExpectedMode\":1,\"Text\":\"Debug.Print 1\"}", "Project mode changed");
        }

        [TestMethod]
        public async Task ReadImmediateChecksModeAndDispatchesOnCallingThread()
        {
            var tools = new LlmVbeTools(null, null, new LlmSettings { VbeEditApproval = "Automatic" });
            int callerThread = Thread.CurrentThread.ManagedThreadId;
            int nativeThread = -1;
            int nativeCalls = 0;
            tools.Execute = request =>
            {
                if (request.Command == "debug_state")
                {
                    var state = new System.Dynamic.ExpandoObject();
                    ((IDictionary<string, object>)state)["Mode"] = 1;
                    return Response.Success(state);
                }
                throw new InvalidOperationException("Unexpected synchronous command.");
            };
            tools.ReadImmediateNative = request =>
            {
                nativeCalls++;
                nativeThread = Thread.CurrentThread.ManagedThreadId;
                Assert.AreEqual("P", request.Project);
                Assert.AreEqual(1, request.ExpectedMode);
                return Task.FromResult<object>(new { Text = "41", Method = "NativeCopy" });
            };
            const string arguments = "{\"Project\":\"P\",\"ExpectedMode\":1}";
            var read = Json.Deserialize<Response>(await tools.InvokeAsync("read_immediate", arguments));
            Assert.IsTrue(read.Ok, read.Error);
            Assert.AreEqual(1, nativeCalls);
            Assert.AreEqual(callerThread, nativeThread);
            await Failure(tools, "read_immediate", "{\"Project\":\"P\",\"ExpectedMode\":2}", "Project mode changed");
            Assert.AreEqual(1, nativeCalls);
            tools.Mode = ChatMode.Plan;
            await Failure(tools, "read_immediate", arguments, UiText.Get(" does not allow this editing or execution tool: "));
            Assert.AreEqual(1, nativeCalls);
        }

        [TestMethod]
        public async Task InspectLocalScalarsChecksScopeModeAndDispatchesAsync()
        {
            var tools = new LlmVbeTools(null, null, new LlmSettings { VbeEditApproval = "Automatic" });
            int calls = 0;
            tools.Execute = request =>
            {
                Assert.AreEqual("debug_state", request.Command);
                var state = new System.Dynamic.ExpandoObject();
                ((IDictionary<string, object>)state)["Mode"] = 1;
                return Response.Success(state);
            };
            tools.InspectLocalScalarsNative = async request =>
            {
                calls++;
                Assert.AreEqual("P", request.Project);
                Assert.AreEqual("M", request.Module);
                Assert.AreEqual("Run", request.Procedure);
                Assert.AreEqual(1, request.ExpectedMode);
                Assert.AreEqual(2, request.Offset);
                Assert.AreEqual(3, request.Limit);
                await Task.Yield();
                return new { Partial = true, Value = "41" };
            };
            const string args = "{\"Project\":\"P\",\"Module\":\"M\",\"Procedure\":\"Run\",\"ExpectedSha256\":\"aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa\",\"ExpectedMode\":1,\"Offset\":2,\"Limit\":3}";
            var inspected = Json.Deserialize<Response>(await tools.InvokeAsync("inspect_local_scalars", args));
            Assert.IsTrue(inspected.Ok, inspected.Error);
            Assert.AreEqual(1, calls);
            Assert.IsFalse(Json.Deserialize<Response>(tools.Invoke("inspect_local_scalars", args)).Ok);
            tools.Mode = ChatMode.Plan;
            await Failure(tools, "inspect_local_scalars", args, UiText.Get(" does not allow this editing or execution tool: "));
            Assert.AreEqual(1, calls);
        }
    }
}

namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using VBAi;

    /// <summary>Vérifie les schémas publiés et les règles de validation des outils VBE.</summary>
    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class LlmVbeToolContractTests
    {
        [TestMethod]
        public void ObjectBrowserRejectsArgumentsBeforeNativeAccess()
        {
            var tools = new LlmVbeTools(null, null, new LlmSettings());
            IsFailure(tools.InvokeAsync("list_object_browser", "{}").GetAwaiter().GetResult(), "Pane");
            IsFailure(tools.InvokeAsync("list_object_browser", "{\"Pane\":\"members\",\"Limit\":\"2\"}").GetAwaiter().GetResult(), "integers");
            Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.SelectObjectBrowser(new Request { Context = "VBA", Procedure = "Count" }));
            Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.SelectObjectBrowser(new Request { Context = " " }));
            Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.ListObjectBrowser(new Request { Pane = "other" }));
            Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.ListObjectBrowser(new Request { Pane = "members", Offset = -1 }));
            Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.ListObjectBrowser(new Request { Pane = "classes", Limit = 201 }));
            IsFailure(tools.InvokeAsync("select_object_browser", "{}").GetAwaiter().GetResult(), "ObjectName");
            IsFailure(tools.InvokeAsync("select_object_browser", "{\"ObjectName\":123}").GetAwaiter().GetResult(), "strings");
            IsFailure(tools.InvokeAsync("select_object_browser", "{\"ObjectName\":\"\"}").GetAwaiter().GetResult(), "ObjectName");
            IsFailure(tools.InvokeAsync("read_object_browser", "[]").GetAwaiter().GetResult(), "empty argument object");
            IsFailure(tools.InvokeAsync("read_object_browser", "{\"Query\":\"Range\"}").GetAwaiter().GetResult(), "empty argument object");
        }

        [TestMethod]
        public void ListInitializerSchemaExposesMatrixWithoutRequiringOneColumnItems()
        {
            var function = LlmVbeTools.Definitions.Select(definition => Dict(Dict(Json.DeserializeObject(Json.Serialize(definition)))["function"]))
                .Single(item => (string)item["name"] == "set_form_list_initializer");
            var parameters = Dict(function["parameters"]);
            var properties = Dict(parameters["properties"]);
            Assert.IsTrue(properties.ContainsKey("Items"));
            Assert.IsTrue(properties.ContainsKey("Rows"));
            Assert.IsFalse(((object[])parameters["required"]).Contains("Items"));
            var matrix = Dict(properties["Rows"]);
            Assert.AreEqual("array", matrix["type"]);
            Assert.AreEqual(64, matrix["maxItems"]);
            Assert.AreEqual(10, Dict(matrix["items"])["maxItems"]);
        }

        /// <summary>Vérifie l’unicité des outils publiés et la présence des champs obligatoires de leurs schémas.</summary>
        [TestMethod]
        public void PublishedToolSchemasHaveUniqueNamesAndRequiredFieldsExist()
        {
            var names = new HashSet<string>(StringComparer.Ordinal);
            foreach (object definition in LlmVbeTools.Definitions)
            {
                var outer = Dict(Json.DeserializeObject(Json.Serialize(definition)));
                Assert.AreEqual("function", outer["type"]);
                var function = Dict(outer["function"]);
                var name = (string)function["name"];
                Assert.IsTrue(names.Add(name), "Duplicate LLM tool: " + name);
                var parameters = Dict(function["parameters"]);
                Assert.AreEqual(false, parameters["additionalProperties"]);
                var properties = Dict(parameters["properties"]);
                foreach (object required in (object[])parameters["required"])
                    Assert.IsTrue(properties.ContainsKey((string)required), name + ": " + required);
            }

            Assert.IsTrue(names.Contains("create_class"));
            Assert.IsTrue(names.Contains("form_tree"));
            Assert.IsTrue(names.Contains("debug_state"));
            Assert.IsTrue(names.Contains("list_reference_types"));
        }

        /// <summary>Refuse les arguments JSON mal formés ou inattendus avant tout accès à l’hôte.</summary>
        [TestMethod]
        public void InvocationRejectsMalformedAndUnexpectedArgumentsBeforeHostAccess()
        {
            var tools = new LlmVbeTools(null, null, new LlmSettings());
            IsFailure(tools.Invoke("missing_tool", "{}"), "Unknown tool");
            IsFailure(tools.Invoke("list_modules", "{}"), "Project is required");
            IsFailure(tools.Invoke("list_modules", "[]"), "object");
            IsFailure(tools.Invoke("list_modules", "{\"Project\":\"P\",\"Extra\":1}"), "Unexpected argument");
            IsFailure(tools.Invoke("list_modules", "{\"Project\":12}"), "must be a string");
            IsFailure(tools.Invoke("git_status", "{}"), "InvokeAsync");
            IsFailure(tools.Invoke("form_tree", "{\"Project\":\"P\"}"), "Form is required");
        }

        /// <summary>Vérifie que les modes et politiques incompatibles empêchent toute modification hôte.</summary>
        [TestMethod]
        public void EditingModesAndPoliciesRejectChangesBeforeHostAccess()
        {
            var settings = new LlmSettings
            {
                VbeEditApproval = "Automatic"
            };
            var tools = new LlmVbeTools(null, null, settings)
            {
                Mode = ChatMode.Plan
            };
            IsFailure(tools.Invoke("create_module", "{}"), "create_module");
            IsFailure(tools.Invoke("run_form", "{}"), "run_form");
            IsFailure(tools.Invoke("native_code_history", "{}"), "native_code_history");
            IsFailure(tools.Invoke("native_form_history", "{}"), "native_form_history");
            IsFailure(tools.Invoke("cut_code", "{}"), "cut_code");
            IsFailure(tools.Invoke("paste_code", "{}"), "paste_code");
            tools.Mode = ChatMode.Agent;
            settings.VbeEditApproval = "ReadOnly";
            IsFailure(tools.Invoke("create_module", "{\"Project\":\"P\",\"Module\":\"M\",\"ExpectedMode\":2}"), "VBE");
            IsFailure(tools.Invoke("run_form", "{\"Project\":\"P\",\"Form\":\"F\",\"ExpectedMode\":2,\"ExpectedSha256\":\"hash\",\"ExpectedTreeVersion\":\"tree\",\"ControlCaption\":\"Run\"}"), "VBE");
            IsFailure(tools.Invoke("native_code_history", "{\"Project\":\"P\",\"Action\":\"undo\",\"ExpectedMode\":2,\"ExpectedProjectVersion\":\"hash\",\"ControlCaption\":\"Undo\"}"), "VBE");
            IsFailure(tools.Invoke("native_form_history", "{\"Project\":\"P\",\"Form\":\"F\",\"Action\":\"undo\",\"ExpectedTreeVersion\":\"tree\"}"), "VBE");
            settings.VbeEditApproval = "invalid";
            IsFailure(tools.Invoke("create_module", "{\"Project\":\"P\",\"Module\":\"M\",\"ExpectedMode\":2}"), "VBE");
            tools.ValidateScope = () =>
            {
                throw new InvalidOperationException("scope changed");
            };
            IsFailure(tools.Invoke("list_projects", "{}"), "scope changed");
        }

        /// <summary>Valide le projet lié et le chemin de fichier avant les commandes hôte.</summary>
        [TestMethod]
        public void BoundProjectAndFilePathAreCheckedBeforeHostAccess()
        {
            var tools = new LlmVbeTools(null, null, new LlmSettings())
            {
                BoundProject = "WorkbookA"
            };
            IsFailure(tools.Invoke("create_module", "{\"Project\":\"WorkbookB\",\"Module\":\"M\",\"ExpectedMode\":2}"), "autre projet");
            IsFailure(tools.Invoke("inspect_code_file", "{\"Path\":\"C:\\\\Temp\\\\code.bas\"}"), "explicitement");
            IsFailure(tools.Invoke("set_form_node_property", "{\"Project\":\"WorkbookA\",\"Form\":\"F\",\"ControlPath\":\"X\",\"ExpectedTreeVersion\":\"v\",\"Property\":\"Caption\",\"Value\":[]}"), "Value must be");
        }

        /// <summary>Permet la découverte des projets via le même protocole en modes Plan et lecture seule.</summary>
        [TestMethod]
        public void PlanAndReadOnlyModesCanDiscoverLiveProjectsThroughTheSameToolProtocol()
        {
            var host = new VbeSessionTests.FakeVbe();
            host.VBProjects.Add(new VbeSessionTests.FakeProject { Name = "ProjetÉté", FileName = @"C:\Temp\ProjetÉté.xlsm", Mode = 2 });
            var tools = new LlmVbeTools(new VbeSession(host), null, new LlmSettings { VbeEditApproval = "ReadOnly" })
            {
                Mode = ChatMode.Plan,
                BoundProject = "ProjetÉté"
            };
            var response = Json.Deserialize<Response>(tools.Invoke("list_projects", "{}"));
            Assert.IsTrue(response.Ok);
            var projects = (object[])response.Data;
            Assert.AreEqual(1, projects.Length);
            Assert.AreEqual("ProjetÉté", (string)Dict(projects[0])["Name"]);
            var status = Json.Deserialize<Response>(tools.Invoke("status", "{}"));
            Assert.IsTrue(status.Ok);
        }

        /// <summary>Exige le chemin absolu exact fourni par l’utilisateur avant de demander l’approbation de lecture.</summary>
        [TestMethod]
        public void FileReadRequiresTheExactUserProvidedAbsolutePathBeforeShowingApproval()
        {
            var tools = new LlmVbeTools(null, null, new LlmSettings());
            string supplied = @"C:\Temp\VBAi-user-file-does-not-exist.txt";
            IsFailure(tools.Invoke("read_user_file", Json.Serialize(new { Path = supplied })), "n'a pas fourni");
            tools.NoteUserRequest("Please inspect " + supplied);
            IsFailure(tools.Invoke("read_user_file", Json.Serialize(new { Path = @"C:\Temp\other.txt" })), "n'a pas fourni");
            IsFailure(tools.Invoke("read_user_file", Json.Serialize(new { Path = supplied })), "introuvable");
        }

        /// <summary>Publie une modification VBA vérifiée et permet de la restaurer.</summary>
        [TestMethod]
        public void AutomaticCodeEditPublishesVerifiedChangeAndCanRestoreIt()
        {
            var module = new VbeSessionTests.FakeModule("Alpha\r\nBeta");
            var project = new VbeSessionTests.FakeProject
            {
                Name = "VBAProject",
                FileName = @"C:\Temp\Host.xlsm",
                Mode = 2
            };
            project.VBComponents.Items.Add(new VbeSessionTests.FakeComponent { Name = "Module1", Type = 1, CodeModule = module });
            var host = new VbeSessionTests.FakeVbe();
            host.VBProjects.Add(project);
            var session = new VbeSession(host);
            dynamic before = session.Execute(new Request { Command = "read_module", Project = project.Name, Module = "Module1" }).Data;
            var tools = new LlmVbeTools(session, null, new LlmSettings { VbeEditApproval = "Automatic" });
            CodeChange observed = null;
            tools.CodeEdited += change => observed = change;
            var result = Json.Deserialize<Response>(tools.Invoke("replace_lines", Json.Serialize(new { Project = project.Name, Module = "Module1", ExpectedSha256 = (string)before.Sha256, StartLine = 2, Count = 1, Text = "Gamma" })));
            Assert.IsTrue(result.Ok, result.Error);
            Assert.AreEqual("Alpha\r\nGamma", module.Code);
            Assert.IsNotNull(observed);
            Assert.AreEqual("Alpha\r\nBeta", observed.Before);
            Assert.AreEqual("Alpha\r\nGamma", observed.After);
            Assert.IsTrue(tools.RestoreCodeChange(observed).Ok);
            Assert.AreEqual("Alpha\r\nBeta", module.Code);
            Assert.IsTrue(observed.Restored);
            Assert.IsFalse(tools.RestoreCodeChange(observed).Ok);
        }

        /// <summary>Vérifie tous les conflits d’annulation avant d’écrire dans un module quelconque.</summary>
        [TestMethod]
        public void MultiModuleUndoPreflightsEveryConflictBeforeWritingAnyModule()
        {
            var project = new VbeSessionTests.FakeProject
            {
                Name = "VBAProject",
                FileName = @"C:\Temp\Host.xlsm",
                Mode = 2
            };
            var first = new VbeSessionTests.FakeComponent
            {
                Name = "First",
                Type = 1,
                CodeModule = new VbeSessionTests.FakeModule("A2")
            };
            var second = new VbeSessionTests.FakeComponent
            {
                Name = "Second",
                Type = 1,
                CodeModule = new VbeSessionTests.FakeModule("unexpected")
            };
            project.VBComponents.Items.Add(first);
            project.VBComponents.Items.Add(second);
            var host = new VbeSessionTests.FakeVbe();
            host.VBProjects.Add(project);
            var tools = new LlmVbeTools(new VbeSession(host), null, new LlmSettings { VbeEditApproval = "Automatic" });
            var stale = new CodeChange(project.Name, second.Name, "B", "", "B2", "", 1);
            var valid = new CodeChange(project.Name, first.Name, "A", "", "A2", "", 1);
            var rejected = tools.RestoreChanges(new[] { stale, valid }, null);
            Assert.IsFalse(rejected.Ok);
            Assert.AreEqual("A2", first.CodeModule.Code);
            Assert.AreEqual("unexpected", second.CodeModule.Code);
            Assert.IsFalse(valid.Restored);
            second.CodeModule = new VbeSessionTests.FakeModule("B2");
            var restored = tools.RestoreChanges(new[] { stale, valid }, null);
            Assert.IsTrue(restored.Ok, restored.Error);
            Assert.AreEqual("A", first.CodeModule.Code);
            Assert.AreEqual("B", second.CodeModule.Code);
            Assert.IsTrue(valid.Restored);
            Assert.IsTrue(stale.Restored);
        }
    }
}

namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Windows.Forms;
    using VBAi;
    using VBAi.Tests.Infrastructure;

    /// <summary>Vérifie les frontières injectables des outils VBE et leurs matrices de validation.</summary>
    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class LlmVbeToolsBoundaryTests
    {
        /// <summary>Vérifie les types, champs requis, valeurs blanches et dispatch hôte du contrat.</summary>
        [TestMethod]
        [STATestMethod]
        public void ContractMatrixChecksRequiredTypesWhitespaceOptionalFieldsAndHostDispatch()
        {
            Assert.ThrowsException<ArgumentNullException>(() => new LlmVbeTools(null, null, null));
            var tools = Create();
            const string reviewedSupport = "Option Explicit\r\n' Reviewed test-support source\r\n";
            tools.Execute = request => request.Command == "preview_vba_test_support"
                ? Response.Success(new { ExpectedProjectVersion = "value", Text = reviewedSupport })
                : VbeToolBoundaryFixture.Execute(request);
            var saveFixture = new ToolFixture();
            saveFixture.Tools.SaveHostDocumentNative = request => System.Threading.Tasks.Task.FromResult<object>(new { Saved = true });
            tools.NoteUserRequest(null); tools.NoteUserRequest(" ");
            tools.NoteUserRequest(@"The supplied path is C:\Temp\fixture.bas");
            foreach (object definition in LlmVbeTools.Definitions)
            {
                var function = Dict(Dict(Json.DeserializeObject(Json.Serialize(definition)))["function"]);
                string name = (string)function["name"];
                if (name == "read_immediate" || name == "inspect_local_scalars" ||
                    name == "read_project_general" || name == "set_project_general" ||
                    name == "create_solidworks_macro" || name == "publish_solidworks_macro")
                {
                    // Awaited validation/dispatch matrices cover these asynchronous commands separately.
                    string refused = tools.Invoke(name, "{}");
                    Failed(refused, name);
                    StringAssert.Contains(refused, "requires InvokeAsync");
                    continue;
                }
                // Monaco has its own awaited contract matrix and real WebView2 dispatch test.
                if (name.StartsWith("git_") || name.StartsWith("monaco_") || LlmVbeTools.IsCatalogTool(name) || name == "read_user_file" || name == "replace_lines") continue;
                var parameters = Dict(function["parameters"]);
                var required = (object[])parameters["required"];
                var fields = Dict(parameters["properties"]);
                var values = Arguments(name);
                // These tools have semantic preconditions in addition to their JSON types.
                if (name == "install_vba_test_support") values["Text"] = reviewedSupport;
                if (name == "run_vba_tests" || name == "navigate_vba_test") values["Items"] = new[] { new string('a', 64) };
                if (name == "vba_test_run_status") values["Action"] = "compact";
                if (name == "run_vba_tests") values["Action"] = "coverage";
                Func<string, string> invoke = arguments => name == "save_host_document"
                    ? InvokeSaveContract(saveFixture, arguments) : tools.Invoke(name, arguments);
                Success(invoke(Json.Serialize(values)), name);
                foreach (string field in fields.Keys)
                {
                    object original = values[field];
                    values[field] = null;
                    Failed(invoke(Json.Serialize(values)), field);
                    values[field] = new object[] { new object() };
                    Failed(invoke(Json.Serialize(values)), name + ":" + field);
                    values[field] = original;
                }
                foreach (string field in required.Cast<string>())
                {
                    object original = values[field]; values.Remove(field);
                    Failed(invoke(Json.Serialize(values)), name + ":missing:" + field);
                    values[field] = original;
                    if (original is string && field != "Text" && field != "Caption" && field != "Value")
                    {
                        values[field] = " "; Failed(invoke(Json.Serialize(values)), field); values[field] = original;
                    }
                }
            }
            foreach (object value in new object[] { "", true, 1, 2147483648L, 1.25m, 1e50 })
            {
                var values = Arguments("set_form_node_property"); values["Value"] = value;
                Success(tools.Invoke("set_form_node_property", Json.Serialize(values)), "scalar " + value);
            }
            foreach (object value in new object[] { 1, 2147483648L, 1.25m, 1e50 })
            {
                var values = Arguments("add_form_control");
                // Every numeric field is checked through its published schema.
                foreach (string field in values.Keys.ToArray())
                    if (field == "Left" || field == "Top" || field == "Width" || field == "Height") values[field] = value;
                Success(tools.Invoke("add_form_control", Json.Serialize(values)), "numeric " + value);
            }
            tools.BoundProject = "P";
            Failed(tools.Invoke("create_module", "[]"), "bound array");
            Failed(tools.Invoke("create_module", "{}"), "bound missing project");
            Success(tools.Invoke("create_module", Json.Serialize(Arguments("create_module"))), "bound same project");
            tools.Mode = ChatMode.Plan;
            Failed(tools.Invoke("create_module", Json.Serialize(Arguments("create_module"))), "plan edit");
            Success(tools.Invoke("list_projects", "{}"), "plan read");
            tools.Mode = ChatMode.Agent;
            tools.Settings.VbeEditApproval = "AskEachTime";
            tools.ShowApproval = (dialog, owner) => DialogResult.No;
            Failed(tools.Invoke("create_module", Json.Serialize(Arguments("create_module"))), "approval refusal");
            tools.ShowApproval = (dialog, owner) => DialogResult.Yes;
            Success(tools.Invoke("create_module", Json.Serialize(Arguments("create_module"))), "approval yes");
        }

        /// <summary>Vérifie les longueurs limites des éléments et refuse les types erronés, débordements et retours à la ligne.</summary>
        [TestMethod]
        public void ItemsMatrixAcceptsBoundaryLengthsAndRejectsMultilineWrongTypesAndOverflow()
        {
            var tools = Create();
            string name = LlmVbeTools.Definitions.Select(d => Dict(Dict(Json.DeserializeObject(Json.Serialize(d)))["function"]))
                .Where(f => Dict(Dict(f["parameters"])["properties"]).ContainsKey("Items")).Select(f => (string)f["name"]).First();
            var values = Arguments(name);
            foreach (object items in new object[] { new string[0], new[] { "", new string('x', 256) }, Enumerable.Repeat("x", 64).ToArray() })
            { values["Items"] = items; Success(tools.Invoke(name, Json.Serialize(values)), "valid items"); }
            foreach (object items in new object[] { "x", Enumerable.Repeat("x", 65).ToArray(), new object[] { 2 }, new[] { new string('x', 257) }, new[] { "line\nline" } })
            { values["Items"] = items; Failed(tools.Invoke(name, Json.Serialize(values)), "invalid items"); }
        }

        /// <summary>Valide les matrices rectangulaires et refuse chaque type ou taille incompatible avec le schéma.</summary>
        [TestMethod]
        public void RowsMatrixValidatesShapeTypesAndCellBoundaries()
        {
            var tools = Create();
            const string name = "set_form_list_initializer";
            var values = Arguments(name);
            values.Remove("Items");
            foreach (object rows in new object[] { new string[0][], new[] { new[] { "", new string('x', 256) } }, Enumerable.Range(0, 64).Select(_ => Enumerable.Repeat("x", 10).ToArray()).ToArray() })
            { values["Rows"] = rows; Success(tools.Invoke(name, Json.Serialize(values)), "valid rows"); }
            foreach (object rows in new object[] { null, "x", new object[] { "x" }, new object[] { null }, new[] { new string[0] }, new[] { Enumerable.Repeat("x", 11).ToArray() }, Enumerable.Range(0, 65).Select(_ => new[] { "x" }).ToArray(), new[] { new[] { "x" }, new[] { "x", "y" } }, new object[] { new object[] { 2 } }, new[] { new[] { new string('x', 257) } }, new[] { new[] { "line\nline" } } })
            { values["Rows"] = rows; Failed(tools.Invoke(name, Json.Serialize(values)), "invalid rows"); }
        }

        /// <summary>Vérifie le chemin absolu littéral, la confirmation et la limite de taille des lectures de fichier.</summary>
        [TestMethod]
        public void FileReadMatrixRequiresLiteralAbsolutePathConfirmationAndTextSizeLimit()
        {
            using (var scope = new LlmBoundaryScope())
            {
                var tools = Create();
                string file = Path.Combine(scope.Root, "document.txt");
                File.WriteAllText(file, "étè", new System.Text.UTF8Encoding(false));
                tools.NoteUserRequest("Read " + file);
                tools.ConfirmFile = (owner, text, title) => { StringAssert.Contains(text, file); return DialogResult.No; };
                Failed(tools.Invoke("read_user_file", Json.Serialize(new { Path = file })), "declined");
                tools.ConfirmFile = (owner, text, title) => DialogResult.Yes;
                var data = Data(tools.Invoke("read_user_file", Json.Serialize(new { Path = file })));
                Assert.AreEqual("étè", data["Text"]); Assert.AreEqual(false, data["Truncated"]);
                tools.CurrentProviderName = "FixtureProvider";
                tools.ConfirmFile = (owner, text, title) => { StringAssert.Contains(text, "FixtureProvider"); return DialogResult.Yes; };
                File.WriteAllBytes(file, new byte[] { 65, 0, 66 });
                Failed(tools.Invoke("read_user_file", Json.Serialize(new { Path = file })), "binary");
                File.WriteAllText(file, new string('a', 65537), new System.Text.UTF8Encoding(false));
                data = Data(tools.Invoke("read_user_file", Json.Serialize(new { Path = file })));
                Assert.AreEqual(65536, ((string)data["Text"]).Length); Assert.AreEqual(true, data["Truncated"]);
                Assert.AreEqual(65537, Convert.ToInt32(data["ByteLength"]));
                Failed(tools.Invoke("inspect_code_file", Json.Serialize(new { Path = "relative.bas" })), "relative path");
            }
        }

        /// <summary>Exécute les frontières natives injectées et transmet les échecs retournés par l’hôte.</summary>
        /// <returns>Tâche terminée lorsque les chemins natifs et les erreurs hôte sont vérifiés.</returns>
        [TestMethod]
        public async Task AsyncDispatchMatrixExecutesNativeBoundariesAndReturnsHostFailures()
        {
            var tools = Create();
            foreach (string name in new[] { "sign_project", "read_project_signature_dialog", "read_debug_options", "read_vbe_options",
                "quick_watch", "edit_watch", "debug_item", "immediate_execute", "debug_dialog", "respond_debug_dialog",
                "remove_watch", "add_watch", "debug_windows" })
            {
                var values = AsyncArguments(name);
                Success(await tools.InvokeAsync(name, Json.Serialize(values)), name);
                tools.Execute = request => Response.Failure("host declined");
                if (name != "debug_windows" && name != "debug_item" && name != "debug_dialog" && name != "respond_debug_dialog")
                    Failed(await tools.InvokeAsync(name, Json.Serialize(values)), name + ":host");
                tools.Execute = VbeToolBoundaryFixture.Execute;
            }
            Failed(await tools.InvokeAsync("debug_dialog", "[]"), "dialog array");
            tools.Native.ReadSignatureDialog = p => { throw new InvalidOperationException("signature dialog unavailable"); };
            Failed(await tools.InvokeAsync("read_project_signature_dialog", Json.Serialize(Arguments("read_project_signature_dialog"))), "signature native exception");
            tools.Native.ReadDebugOptions = () => { throw new InvalidOperationException("options unavailable"); };
            Failed(await tools.InvokeAsync("read_debug_options", "{}"), "options native exception");
            tools.Native.VerifyWatchRemoved = r => { throw new InvalidOperationException("watch native unavailable"); };
            Failed(await tools.InvokeAsync("remove_watch", Json.Serialize(Arguments("remove_watch"))), "remove native exception");
            tools.Native.ReadDebugDialog = () => { throw new InvalidOperationException("native declined"); };
            Failed(await tools.InvokeAsync("debug_dialog", "{}"), "native exception");
            tools.Native.CompleteAddWatch = r => { throw new InvalidOperationException("native add declined"); };
            Failed(await tools.InvokeAsync("add_watch", Json.Serialize(Arguments("add_watch"))), "native add exception");
            tools.Native.Capture = stack => new { Stack = stack };
            foreach (string args in new[] { "{}", "{\"IncludeCallStack\":true}", "{\"IncludeCallStack\":false}" })
                Success(await tools.InvokeAsync("debug_windows", args), args);
            foreach (string name in new[] { "sign_project", "read_project_signature_dialog", "quick_watch", "edit_watch", "remove_watch", "add_watch" })
                Failed(await tools.InvokeAsync(name, "{}"), name + ":missing");
            foreach (string args in new[] { "[]", "{\"Diagnostic\":\"x\"}", "{\"Diagnostic\":1,\"Button\":\"ok\"}", "{\"Diagnostic\":\"x\",\"Button\":1}" })
                Failed(await tools.InvokeAsync("respond_debug_dialog", args), args);
            foreach (string args in new[] { "[]", "{}", "{\"Action\":\"expand\",\"PathSegments\":[\"x\"]}", "{\"Pane\":\"locals\",\"PathSegments\":[\"x\"]}", "{\"Pane\":\"locals\",\"Action\":\"expand\"}", "{\"Pane\":\"locals\",\"Action\":\"expand\",\"PathSegments\":[\"x\"],\"Context\":\"p\"}" })
            {
                if (args.Contains("Context")) Success(await tools.InvokeAsync("debug_item", args), args);
                else Failed(await tools.InvokeAsync("debug_item", args), args);
            }
            var immediate = AsyncArguments("immediate_execute");
            foreach (string field in immediate.Keys.ToArray())
            {
                object original = immediate[field]; immediate.Remove(field);
                Failed(await tools.InvokeAsync("immediate_execute", Json.Serialize(immediate)), "missing " + field);
                immediate[field] = original is string ? (object)true : "2";
                Failed(await tools.InvokeAsync("immediate_execute", Json.Serialize(immediate)), "type " + field);
                immediate[field] = original;
            }
            immediate["ExpectedMode"] = 1;
            Failed(await tools.InvokeAsync("immediate_execute", Json.Serialize(immediate)), "changed mode");
            immediate["ExpectedMode"] = 3;
            Failed(await tools.InvokeAsync("immediate_execute", Json.Serialize(immediate)), "invalid mode");
            Failed(await tools.InvokeAsync("immediate_execute", "[]"), "immediate array");
            tools.Settings.VbeEditApproval = "ReadOnly";
            Failed(await tools.InvokeAsync("remove_watch", Json.Serialize(Arguments("remove_watch"))), "remove policy");
        }

        /// <summary>Vérifie la persistance de signature sans certificat et les nouvelles tentatives de sauvegarde.</summary>
        /// <returns>Tâche terminée lorsque les différentes réponses de signature sont vérifiées.</returns>
        [TestMethod]
        public async Task SignaturePersistenceMatrixHandlesMissingCertificateAndNeverReplaysSave()
        {
            var tools = Create(); string args = Json.Serialize(Arguments("sign_project"));
            tools.Execute = r => Response.Success(new { Missing = true });
            Failed(await tools.InvokeAsync("sign_project", args), "missing certificate");
            tools.Execute = r => Response.Success("unexpected host shape");
            Failed(await tools.InvokeAsync("sign_project", args), "host shape");
            tools.Execute = VbeToolBoundaryFixture.Execute;
            foreach (bool saved in new[] { false, true })
            {
                tools.PersistSignature = (p, authorize) => { authorize(); return new VBAi.Tests.Infrastructure.VbeToolPersistence { Saved = saved }; };
                Assert.AreEqual(!saved, Data(await tools.InvokeAsync("sign_project", args))["SaveRequired"]);
            }
            tools.PersistSignature = (p, authorize) => { authorize(); return null; };
            Assert.AreEqual(true, Data(await tools.InvokeAsync("sign_project", args))["SaveRequired"]);
            tools.PersistSignature = (p, authorize) => { authorize(); throw new InvalidOperationException("save declined"); };
            tools.Execute = r => r.Command == "project_signature_status" ? Response.Failure("status declined") : VbeToolBoundaryFixture.Execute(r);
            var data = Data(await tools.InvokeAsync("sign_project", args));
            Assert.AreEqual("save declined", data["PersistenceError"]); Assert.AreEqual("status declined", data["HostStatusError"]);
            int attempts = 0;
            tools.PersistSignature = (p, authorize) => { authorize(); if (++attempts == 2) return new VBAi.Tests.Infrastructure.VbeToolPersistence { Saved = true }; throw new InvalidOperationException("0x800AC472 busy"); };
            Assert.AreEqual(true, Data(await tools.InvokeAsync("sign_project", args))["SaveRequired"]); Assert.AreEqual(1, attempts);
            attempts = 0; tools.PersistSignature = (p, authorize) => { authorize(); attempts++; throw new InvalidOperationException("0x800AC472 busy"); };
            data = Data(await tools.InvokeAsync("sign_project", args)); Assert.AreEqual(1, attempts);
            StringAssert.Contains((string)data["PersistenceError"], "busy");
        }

        /// <summary>Vérifie le travail UI différé de compilation et distingue délais, diagnostics et erreurs.</summary>
        [TestMethod]
        [STATestMethod]
        public void CompileMatrixUsesPostedUiWorkAndDistinguishesTimeoutDiagnosisAndErrors()
        {
            var prior = SynchronizationContext.Current;
            try
            {
                var tools = Create();
                SynchronizationContext.SetSynchronizationContext(new ImmediateContext());
                const string args = "{\"Project\":\"P\",\"ExpectedMode\":2}";
                Assert.AreEqual(true, Data(tools.InvokeAsync("compile_project", args).GetAwaiter().GetResult())["Compiled"]);
                tools.Native.AwaitCompileDialog = completed => { Assert.IsTrue(completed.Wait(5000)); return "compile diagnostic"; };
                var data = Data(tools.InvokeAsync("compile_project", args).GetAwaiter().GetResult());
                Assert.AreEqual(false, data["Compiled"]); Assert.AreEqual("NativeDiagnosticCaptured", data["Verification"]);
                tools.Execute = r => Response.Success(new { Executed = false, Available = false, Capability = "Absent", Reason = "not available" });
                data = Data(tools.InvokeAsync("compile_project", args).GetAwaiter().GetResult());
                Assert.AreEqual(false, data["Compiled"]); Assert.AreEqual("NativeCompileAbsent", data["Verification"]);
                tools.Execute = r => Response.Failure("compile failure"); Failed(tools.InvokeAsync("compile_project", args).GetAwaiter().GetResult(), "compile failure");
                tools.Execute = r => { throw new InvalidOperationException("compile exception"); };
                Failed(tools.InvokeAsync("compile_project", args).GetAwaiter().GetResult(), "compile exception");
                var deferred = new DeferredContext();
                SynchronizationContext.SetSynchronizationContext(deferred);
                int lateExecutions = 0;
                tools.Execute = r => { lateExecutions++; return VbeToolBoundaryFixture.Execute(r); };
                tools.Native.AwaitCompileDialog = completed => null;
                Failed(tools.InvokeAsync("compile_project", args).GetAwaiter().GetResult(), "timeout");
                deferred.ReleasePending();
                Assert.AreEqual(0, lateExecutions, "A cancelled queued callback must never compile later.");
                SynchronizationContext.SetSynchronizationContext(new ImmediateContext());
                foreach (string invalid in new[] { "[]", "{}", "{\"ExpectedMode\":2}", "{\"Project\":\"P\"}", "{\"Project\":1,\"ExpectedMode\":2}", "{\"Project\":\" \",\"ExpectedMode\":2}", "{\"Project\":\"P\",\"ExpectedMode\":\"2\"}", "{\"Project\":\"P\",\"ExpectedMode\":1}" })
                    Failed(tools.InvokeAsync("compile_project", invalid).GetAwaiter().GetResult(), invalid);
            }
            finally { SynchronizationContext.SetSynchronizationContext(prior); }
        }
        /// <summary>Refuse les lectures obsolètes et couvre le code inchangé ainsi que les erreurs d’abonné.</summary>
        [TestMethod]
        public void CodeEditReadbackMatrixRejectsStaleReadsAndHandlesUnchangedCodeOrSubscriberErrors()
        {
            var host = new VbeSessionTests.FakeVbe();
            var module = new VbeSessionTests.FakeModule("A\r\nB");
            var project = new VbeSessionTests.FakeProject { Name = "P", FileName = @"C:\Temp\fixture.xlsm", Mode = 2 };
            project.VBComponents.Items.Add(new VbeSessionTests.FakeComponent { Name = "M", Type = 1, CodeModule = module }); host.VBProjects.Add(project);
            var session = new VbeSession(host);
            var tools = new LlmVbeTools(session, null, new LlmSettings { VbeEditApproval = "Automatic" });
            var logs = new List<string>(); tools.WriteLog = logs.Add;
            Func<string, string> arguments = text => Json.Serialize(new { Project = "P", Module = "M", ExpectedSha256 = (string)((dynamic)session.Execute(new Request { Command = "read_module", Project = "P", Module = "M" }).Data).Sha256, StartLine = 1, Count = 2, Text = text });
            string args = arguments("A\r\nC");
            module.DeleteLines(1, module.CountOfLines); module.InsertLines(1, "stale");
            Failed(tools.Invoke("replace_lines", args), "stale code");
            module.DeleteLines(1, module.CountOfLines); module.InsertLines(1, "A\r\nB");
            tools.Execute = r => Response.Failure("read unavailable");
            Failed(tools.Invoke("replace_lines", args), "read unavailable");
            tools.Execute = session.Execute;
            Success(tools.Invoke("replace_lines", arguments("A\r\nB")), "unchanged code");
            Success(tools.Invoke("replace_lines", arguments("A\r\nC")), "edit without subscriber");
            tools.CodeEdited += change => { throw new InvalidOperationException("subscriber failed"); };
            Success(tools.Invoke("replace_lines", arguments("A\r\nD")), "subscriber error is contained");
            int reads = 0;
            tools.Execute = r => r.Command == "read_module" && ++reads == 2 ? Response.Failure("readback failed") : session.Execute(r);
            Success(tools.Invoke("replace_lines", arguments("A\r\nE")), "readback error is contained");
            Assert.AreEqual(2, logs.Count);
            StringAssert.Contains(logs[0], "subscriber failed");
            StringAssert.Contains(logs[1], "readback failed");
            tools.Execute = r => r.Command == "replace_lines" ? Response.Failure("write refused") : session.Execute(r);
            Failed(tools.Invoke("replace_lines", arguments("A\r\nF")), "write refusal");
        }

        /// <summary>Vérifie la restauration avec entrées nulles, blocs partagés et échec d’écriture.</summary>
        [TestMethod]
        public void RestorationMatrixHandlesNullEntriesSharedModuleHunksAndWriteFailure()
        {
            var host = new VbeSessionTests.FakeVbe();
            var module = new VbeSessionTests.FakeModule("A2\r\nKeep\r\nB2");
            var project = new VbeSessionTests.FakeProject { Name = "P", FileName = @"C:\Temp\fixture.xlsm", Mode = 2 };
            project.VBComponents.Items.Add(new VbeSessionTests.FakeComponent { Name = "M", Type = 1, CodeModule = module }); host.VBProjects.Add(project);
            var session = new VbeSession(host);
            var settings = new LlmSettings { VbeEditApproval = "Automatic" };
            var tools = new LlmVbeTools(session, null, settings) { Mode = ChatMode.Plan };
            Assert.IsFalse(tools.RestoreChanges(new CodeChange[] { null }, null).Ok);
            var change = new CodeChange("P", "M", "A\r\nKeep\r\nB", "", "A2\r\nKeep\r\nB2", "", 3);
            tools.Execute = r => r.Command == "replace_lines" ? Response.Failure("write refused") : session.Execute(r);
            var failure = tools.RestoreChanges(new[] { change }, null);
            Assert.IsFalse(failure.Ok); StringAssert.Contains(failure.Error, "write refused"); Assert.IsFalse(change.Restored);
            tools.Execute = session.Execute;
            Assert.IsTrue(tools.RestoreChanges(new[] { change }, 0).Ok);
            Assert.AreEqual("A\r\nKeep\r\nB2", module.Code); Assert.IsFalse(change.Restored);
            Assert.IsFalse(tools.RestoreChanges(new[] { change }, 0).Ok);
            Assert.IsTrue(tools.RestoreChanges(new[] { change }, null).Ok); Assert.IsTrue(change.Restored);
            Assert.AreEqual("A\r\nKeep\r\nB", module.Code);
            var first = new CodeChange("P", "M", "initial", "", "middle", "", 1);
            var second = new CodeChange("P", "M", "middle", "", "last", "", 1);
            module.DeleteLines(1, module.CountOfLines); module.InsertLines(1, "last");
            Assert.IsTrue(tools.RestoreChanges(new[] { first, second }, null).Ok);
            Assert.AreEqual("initial", module.Code); Assert.IsTrue(first.Restored); Assert.IsTrue(second.Restored);
            tools.ValidateScope = () => { throw new InvalidOperationException("stale scope"); };
            Assert.AreEqual("stale scope", tools.RestoreChanges(new[] { first }, null).Error);
            Assert.ThrowsException<InvalidOperationException>(() => tools.LiveContextJson());
        }

        /// <summary>Conserve les contrats de réponse pour réussite, échec, null et JSON invalide.</summary>
        [TestMethod]
        public void ToolResponseParserKeepsSuccessFailureNullAndInvalidJsonContracts()
        {
            var tools = new LlmVbeTools(null, null, new LlmSettings());
            var missing = tools.ReadToolResponse("null");
            Assert.IsNotNull(missing); Assert.IsFalse(missing.Ok); Assert.IsNull(missing.Error);
            var success = tools.ReadToolResponse(Json.Serialize(Response.Success(new { Value = "fixture" })));
            Assert.IsTrue(success.Ok); Assert.AreEqual("fixture", Dict(success.Data)["Value"]);
            var failure = tools.ReadToolResponse(Json.Serialize(Response.Failure("declined")));
            Assert.IsFalse(failure.Ok); Assert.AreEqual("declined", failure.Error);
            Assert.ThrowsException<ArgumentException>(() => tools.ReadToolResponse("{invalid}"));
        }


        [TestMethod]
        public async Task OptionMutationAndProcedureArgumentsExerciseValidatedAsyncBoundaries()
        {
            var tools = new ToolFixture().Tools; int calls = 0;
            tools.Native.SetVbeOption = request => { calls++; Assert.AreEqual("editor", request.Pane); return new { Changed = true }; };
            var option = Arguments("set_vbe_option"); option["Pane"] = "editor";
            Success(await tools.InvokeAsync("set_vbe_option", Json.Serialize(option)), "set option"); Assert.AreEqual(1, calls);
            tools.Native.SetVbeOption = request => throw new InvalidOperationException("native option rejected");
            Failed(await tools.InvokeAsync("set_vbe_option", Json.Serialize(option)), "native option rejected");
            tools.Execute = request => Response.Failure("schedule rejected");
            Failed(await tools.InvokeAsync("set_vbe_option", Json.Serialize(option)), "schedule rejected");
            tools.Execute = Infrastructure.VbeToolBoundaryFixture.Execute;
            var run = Arguments("run_procedure");
            foreach (object arguments in new object[] { null, "not an array", new object[31] })
            { run["Arguments"] = arguments; Failed(tools.Invoke("run_procedure", Json.Serialize(run)), "invalid arguments"); }
            run["Arguments"] = new object[] { "literal", true, 2, 1.5 }; Success(tools.Invoke("run_procedure", Json.Serialize(run)), "scalars");
            run["Arguments"] = new object[] { new { Nested = true } }; Failed(tools.Invoke("run_procedure", Json.Serialize(run)), "non scalar");
            run["Arguments"] = new object[] { 7 };
            foreach (object names in new object[] { null, "not an array", new string[31], new object[] { 1 },
                new[] { new string('x', 256) }, new[] { "first\nEnd" } })
            { run["ArgumentNames"] = names; Failed(tools.Invoke("run_procedure", Json.Serialize(run)), "invalid names"); }
            Request captured = null;
            tools.Execute = request => { captured = request; return Response.Success("delivered"); };
            run["ArgumentNames"] = new[] { "first" };
            Success(tools.Invoke("run_procedure", Json.Serialize(run)), "named scalar");
            CollectionAssert.AreEqual(new[] { "first" }, captured.ArgumentNames);
        }
    }
}

namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Web.Script.Serialization;
    using VBAi;

    public sealed partial class VbeProcedureMutationTests
    {
        /// <summary>Vérifie le routage LLM, les modes, la portée, le diff et la restauration du renommage privé.</summary>
        [TestMethod]
        public void ParameterRenameToolsRespectReadOnlyAndProjectScopeThenPublishRestorableDiff()
        {
            const string original = "Private Sub Run(ByVal value As Long)\r\nDebug.Print value\r\nEnd Sub\r\nPublic Sub Caller()\r\nRun value:=7\r\nEnd Sub";
            var fixture = new Fixture(original);
            var host = new FakeVbe(); host.VBProjects.Add(fixture.Project);
            var session = new VbeSession(host);
            var json = new JavaScriptSerializer();
            var request = fixture.Request(null); request.Query = "value"; request.NewName = "amount";
            request.StartLine = 1; request.StartColumn = original.IndexOf("value", StringComparison.Ordinal) + 1; request.ExpectedMode = 2;
            var arguments = json.Serialize(new
            {
                request.Project,
                request.Module,
                request.Procedure,
                request.ProcKind,
                request.Query,
                request.NewName,
                request.StartLine,
                request.StartColumn,
                request.ExpectedMode,
                request.ExpectedSha256
            });
            var tools = new LlmVbeTools(session, null, new LlmSettings { VbeEditApproval = "Automatic" })
            {
                Mode = ChatMode.Plan,
                BoundProject = fixture.Project.Name
            };
            string previewArguments = json.Serialize(new
            {
                request.Project,
                request.Module,
                request.Procedure,
                request.ProcKind,
                request.Query,
                request.NewName,
                request.StartLine,
                request.StartColumn,
                request.ExpectedSha256
            });
            var preview = json.Deserialize<Response>(tools.Invoke("preview_parameter_rename", previewArguments));
            Assert.IsTrue(preview.Ok, preview.Error); Assert.AreEqual(original, fixture.Module.Code);
            var forbidden = json.Deserialize<Response>(tools.Invoke("apply_parameter_rename", arguments));
            Assert.IsFalse(forbidden.Ok); Assert.AreEqual(original, fixture.Module.Code);
            tools.Mode = ChatMode.Agent; tools.BoundProject = "AnotherProject";
            forbidden = json.Deserialize<Response>(tools.Invoke("apply_parameter_rename", arguments));
            Assert.IsFalse(forbidden.Ok); Assert.AreEqual(original, fixture.Module.Code);
            tools.BoundProject = fixture.Project.Name;
            var readOnly = new LlmVbeTools(session, null, new LlmSettings { VbeEditApproval = "ReadOnly" });
            Assert.IsFalse(json.Deserialize<Response>(readOnly.Invoke("apply_parameter_rename", arguments)).Ok);
            CodeChange observed = null; tools.CodeEdited += change => observed = change;
            var applied = json.Deserialize<Response>(tools.Invoke("apply_parameter_rename", arguments));
            Assert.IsTrue(applied.Ok, applied.Error); Assert.IsNotNull(observed);
            Assert.AreEqual(original, observed.Before); Assert.AreEqual(fixture.Module.Code, observed.After);
            StringAssert.Contains(observed.After, "Run amount:=7");
            Assert.IsTrue(tools.RestoreCodeChange(observed).Ok); Assert.AreEqual(original, fixture.Module.Code);
            tools.Mode = ChatMode.Plan;
            string received = null;
            tools.Execute = input => { received = input.Command; return Response.Success(new { Status = "VerifierUnavailable" }); };
            Assert.IsTrue(json.Deserialize<Response>(tools.Invoke("verify_vba_signature_file", json.Serialize(new { Path = @"C:\Temp\Example.xlsm" }))).Ok);
            Assert.AreEqual("verify_vba_signature_file", received);
        }
    }
}

namespace VBAi.Tests.Unit
{
    public sealed partial class LlmVbeToolsBoundaryTests
    {
        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public void SynchronousCatalogueGatewaysRefuseBeforeAnyHostDispatch()
        {
            var tools = Create(); int requests = 0;
            tools.Execute = request => { requests++; throw new Microsoft.VisualStudio.TestTools.UnitTesting.AssertFailedException("Catalogue requires async dispatch"); };
            foreach (string name in new[] { "discover_tools", "invoke_tool" }) Failed(tools.Invoke(name, "{}"), "InvokeAsync");
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(0, requests);
        }
    }
}

namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Threading.Tasks;
    using VBAi;

    public sealed partial class LlmVbeToolsBoundaryTests
    {
        [DataTestMethod]
        [DataRow("policy")]
        [DataRow("binding")]
        [DataRow("mode")]
        [DataRow("scope")]
        [DataRow("identity")]
        public async Task DeferredSignatureRevalidatesAuthorityBeforePersisting(string changed)
        {
            var fixture = new ToolFixture(); var tools = fixture.Tools;
            tools.BoundProject = "P"; int saves = 0, statusReads = 0; bool identityCurrent = true;
            tools.CaptureSignaturePersistence = p => () => { if (!identityCurrent) throw new InvalidOperationException("identity changed"); };
            tools.PersistSignature = (p, authorize) => { authorize(); saves++; return new VBAi.Tests.Infrastructure.VbeToolPersistence { Saved = true }; };
            tools.Execute = r => { if (r.Command == "project_signature_status") statusReads++; return VBAi.Tests.Infrastructure.VbeToolBoundaryFixture.Execute(r); };
            tools.Native.CompleteProjectSignature = (p, thumbprint, name, unsigned) =>
            {
                if (changed == "policy") fixture.Settings.VbeEditApproval = "ReadOnly";
                if (changed == "binding") tools.BoundProject = "Other";
                if (changed == "mode") tools.Mode = ChatMode.Discussion;
                if (changed == "scope") tools.ValidateScope = () => { throw new InvalidOperationException("scope changed"); };
                if (changed == "identity") identityCurrent = false;
                return new { SignatureAssigned = true };
            };
            var result = Data(await tools.InvokeAsync("sign_project", Json.Serialize(Arguments("sign_project"))));
            Assert.AreEqual(0, saves); Assert.AreEqual(0, statusReads);
            Assert.AreEqual(true, result["SaveRequired"]);
            Assert.IsFalse(string.IsNullOrEmpty((string)result["PersistenceError"]));
            Assert.IsNotNull(result["Signature"]);
        }
    }
}
