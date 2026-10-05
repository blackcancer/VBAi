using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class LlmVbeToolsSolidWorksMacrosTests
    {
        private static readonly JavaScriptSerializer Json = new JavaScriptSerializer();
        private const string Path = @"C:\Qualification\Published.swp";
        private static string Arguments(string command)
        {
            var fields = new Dictionary<string, object> { ["Path"] = Path, ["ExpectedMode"] = 2, ["ExpectedProjectVersion"] = "version" };
            if (command == "publish_solidworks_macro") fields["Project"] = "Draft";
            return Json.Serialize(fields);
        }
        private static LlmVbeTools Tools(LlmSettings settings = null)
        {
            var tools = new LlmVbeTools(null, null, settings ?? new LlmSettings { VbeEditApproval = "Automatic" })
                { BoundProject = "Draft" };
            tools.Execute = _ => Response.Success(new object[0]);
            tools.SetReadAccess(new string[0], true);
            tools.NoteUserRequest(Path);
            return tools;
        }
        private sealed class QueueContext : SynchronizationContext
        {
            private readonly Queue<Action> queue = new Queue<Action>();
            public override void Post(SendOrPostCallback callback, object state) => queue.Enqueue(() => callback(state));
            internal void Drain()
            {
                int count = 0;
                while (queue.Count != 0)
                {
                    Assert.IsTrue(++count < 100, "Native test dispatch exceeded its finite queue.");
                    queue.Dequeue()();
                }
            }
        }
        private static string Invoke(LlmVbeTools tools, string name, bool catalog)
        {
            var old = SynchronizationContext.Current;
            var context = new QueueContext();
            SynchronizationContext.SetSynchronizationContext(context);
            try
            {
                var pending = tools.InvokeAsync(catalog ? "invoke_tool" : name,
                    catalog ? Json.Serialize(new { ToolName = name, ArgumentsJson = Arguments(name) }) : Arguments(name));
                context.Drain();
                Assert.IsTrue(pending.IsCompleted, "The injected native boundary must settle.");
                return pending.GetAwaiter().GetResult();
            }
            finally { SynchronizationContext.SetSynchronizationContext(old); }
        }

        [DataTestMethod, DataRow("create_solidworks_macro"), DataRow("publish_solidworks_macro")]
        public void SynchronousInvocationRefusesBeforeNativeEntry(string command)
        {
            var tools = Tools(); int entries = 0;
            tools.SolidWorksMacroNative = _ => { entries++; return Task.FromResult<object>(null); };
            var result = Json.Deserialize<Response>(tools.Invoke(command, Arguments(command)));
            Assert.IsFalse(result.Ok); StringAssert.Contains(result.Error, "InvokeAsync"); Assert.AreEqual(0, entries);
        }

        [DataTestMethod]
        [DataRow("create_solidworks_macro", false)][DataRow("create_solidworks_macro", true)]
        [DataRow("publish_solidworks_macro", false)][DataRow("publish_solidworks_macro", true)]
        public void NativeRoutesAndCatalogPreserveOneApprovalAndIdentityMapping(string command, bool catalog)
        {
            var tools = Tools(new LlmSettings { VbeEditApproval = "AskEachTime" });
            int approvals = 0, entries = 0; Request captured = null;
            tools.ShowApproval = (_, __) => { approvals++; return DialogResult.Yes; };
            tools.SolidWorksMacroNative = async request => {
                captured = request; entries++;
                Assert.AreEqual(command, request.Command);
                await Task.Yield();
                request.RevalidateMacroAuthorization(true); request.RevalidateMacroAuthorization(false);
                return new { Terminal = true, Verified = true, Uncertain = false, MutationInvoked = true,
                    OriginalProject = "Draft", DestinationProject = "Published", IdentityChanged = true, OriginalPreserved = true };
            };
            string output = Invoke(tools, command, catalog);
            Assert.IsTrue(Json.Deserialize<Response>(output).Ok, output);
            StringAssert.Contains(output, "Published"); StringAssert.Contains(output, "OriginalPreserved");
            Assert.AreEqual(1, approvals); Assert.AreEqual(1, entries); Assert.IsNull(captured.RevalidateMacroAuthorization);
        }

        [DataTestMethod]
        [DataRow("policy", false)][DataRow("binding", false)][DataRow("shared", false)][DataRow("grants", false)]
        [DataRow("mode", true)][DataRow("cached", true)][DataRow("scope", true)]
        public void RevocationWhilePendingRefusesFurtherNativeMutation(string changed, bool catalog)
        {
            var settings = new LlmSettings { VbeEditApproval = "Automatic" };
            var tools = Tools(settings); bool scope = true, cached = true; int writes = 0; Request captured = null;
            tools.ValidateScope = () => { if (!scope) throw new InvalidOperationException("Scope revoked"); };
            tools.ValidateCachedScope = () => { if (!cached) throw new InvalidOperationException("Cached scope revoked"); };
            tools.SolidWorksMacroNative = async request => {
                captured = request; await Task.Yield();
                if (changed == "policy") settings.VbeEditApproval = "ReadOnly";
                if (changed == "binding") tools.BoundProject = "Other";
                if (changed == "shared") tools.SetReadAccess(new string[0], false);
                if (changed == "grants") tools.SetReadAccess(new[] { "Other" }, true);
                if (changed == "mode") tools.Mode = ChatMode.Discussion;
                if (changed == "cached") cached = false;
                if (changed == "scope") scope = false;
                request.RevalidateMacroAuthorization(true); request.RevalidateMacroAuthorization(false);
                writes++; return new { Verified = true };
            };
            var output = Json.Deserialize<Response>(Invoke(tools, "publish_solidworks_macro", catalog));
            Assert.IsFalse(output.Ok); Assert.AreEqual(0, writes); Assert.IsNull(captured.RevalidateMacroAuthorization);
        }

        [DataTestMethod, DataRow("create_solidworks_macro"), DataRow("publish_solidworks_macro")]
        public void SharedContextIsRequiredBeforeCreationOrPublication(string command)
        {
            var tools = Tools(); tools.SetReadAccess(new string[0], false); int entries = 0;
            tools.SolidWorksMacroNative = _ => { entries++; return Task.FromResult<object>(null); };
            Assert.IsFalse(Json.Deserialize<Response>(Invoke(tools, command, false)).Ok); Assert.AreEqual(0, entries);
        }

        [DataTestMethod, DataRow("create_solidworks_macro"), DataRow("publish_solidworks_macro")]
        public void DestinationMustBeExplicitlyProvidedByUserBeforeNativeEntry(string command)
        {
            var tools = new LlmVbeTools(null, null, new LlmSettings { VbeEditApproval = "Automatic" }) { BoundProject = "Draft" };
            tools.SetReadAccess(new string[0], true); tools.Execute = _ => Response.Success(new object[0]);
            int entries = 0;
            tools.SolidWorksMacroNative = _ => { entries++; return Task.FromResult<object>(null); };
            Assert.IsFalse(Json.Deserialize<Response>(Invoke(tools, command, false)).Ok); Assert.AreEqual(0, entries);
        }

        [DataTestMethod, DataRow("create_solidworks_macro"), DataRow("publish_solidworks_macro")]
        public void RevokedAccessAfterNativeMutationRedactsDestinationAndKeepsUncertainty(string command)
        {
            var tools = Tools();
            tools.SolidWorksMacroNative = async request => {
                await Task.Yield(); tools.BoundProject = "Other";
                return new { Verified = true, MutationInvoked = true, CommandEntered = true,
                    OriginalProject = "Draft", DestinationProject = "DoNotDiscloseNativeTarget" };
            };
            string output = Invoke(tools, command, true);
            var result = Json.Deserialize<Response>(output);
            Assert.IsTrue(result.Ok); Assert.IsFalse(output.Contains("DoNotDiscloseNativeTarget"));
            var data = (IDictionary<string, object>)((IDictionary<string, object>)Json.DeserializeObject(output))["Data"];
            Assert.AreEqual(true, data["MutationInvoked"]); Assert.AreEqual(true, data["Uncertain"]);
            Assert.AreEqual(false, data["Verified"]); Assert.AreEqual(false, data["RetryAllowed"]);
        }
    }
}
