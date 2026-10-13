using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using VBAi.Tests.Infrastructure;

namespace VBAi.Tests.Unit
{
    public sealed partial class LlmVbeToolsBoundaryTests
    {
        /// <summary>Exercises actual renderer suspension before native compilation or a guarded refusal.</summary>
        /// <param name="baseContext">Uses the default context rather than no context.</param>
        /// <param name="catalog">Dispatches through the guarded catalogue gateway.</param>
        /// <param name="outcome">Renderer completion, an unsynchronized draft, unavailability or failure.</param>
        [WinFormsTestMethod]
        [DataRow(false, false, "ready"), DataRow(false, true, "ready")]
        [DataRow(true, false, "ready"), DataRow(true, true, "ready")]
        [DataRow(false, false, "draft"), DataRow(false, true, "draft")]
        [DataRow(true, false, "draft"), DataRow(true, true, "draft")]
        [DataRow(false, false, "unavailable"), DataRow(false, true, "unavailable")]
        [DataRow(true, false, "unavailable"), DataRow(true, true, "unavailable")]
        [DataRow(false, false, "failure"), DataRow(false, true, "failure")]
        [DataRow(true, false, "failure"), DataRow(true, true, "failure")]
        public void CompileRetainsOwnerAcrossRendererAwaitAndPreservesDraftGuards(bool baseContext, bool catalog, string outcome)
        {
            using (var editor = new ModernEditorDebugFixture())
            {
                editor.Ready(true);
                int owner = Thread.CurrentThread.ManagedThreadId, rendererCalls = 0, compileCalls = 0;
                string originalSource = editor.Native.Adapter.Read();
                var fixture = new ToolFixture();
                fixture.Tools.MonacoWindow = create =>
                {
                    Assert.IsFalse(create);
                    Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                    return editor.Window;
                };
                editor.Window.ScriptExecution = async (method, values) =>
                {
                    Assert.AreEqual("snapshots", method);
                    Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                    rendererCalls++;
                    await Task.Delay(2);
                    Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                    if (outcome == "draft") editor.Document.Edit("unsynchronized draft");
                    if (outcome == "unavailable") editor.Ready(false);
                    if (outcome == "failure") throw new InvalidOperationException("owned renderer failure");
                    return "null";
                };
                fixture.Tools.Native.EnsureNoCompileDialog = () => Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                fixture.Tools.Native.AwaitCompileDialog = completed =>
                {
                    if (!completed.Wait(5000)) throw new TimeoutException("The original compilation callback did not finish.");
                    return null;
                };
                var execute = fixture.Tools.Execute;
                fixture.Tools.Execute = request =>
                {
                    Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                    if (request.Command != "compile_project") return execute(request);
                    compileCalls++;
                    return Response.Failure("native compile entry reached");
                };
                string arguments = Json.Serialize(new { Project = "P", ExpectedMode = 2 });
                if (catalog) arguments = Json.Serialize(new { ToolName = "compile_project", ArgumentsJson = arguments });
                var previous = SynchronizationContext.Current;
                var ambient = baseContext ? new SynchronizationContext() : null;
                try
                {
                    SynchronizationContext.SetSynchronizationContext(ambient);
                    var pending = fixture.Tools.InvokeAsync(catalog ? "invoke_tool" : "compile_project", arguments);
                    Assert.AreSame(ambient, SynchronizationContext.Current, "Caller context must be restored before any message pump.");
                    var response = Json.Deserialize<Response>(ModernEditorDebugFixture.Wait(pending));
                    Assert.IsFalse(response.Ok);
                    Assert.AreEqual(1, rendererCalls);
                    Assert.AreEqual(outcome == "ready" ? 1 : 0, compileCalls);
                    string expected = outcome == "ready" ? "native compile entry reached"
                        : outcome == "draft" ? "unsynchronized changes"
                        : outcome == "unavailable" ? "Monaco became unavailable" : "owned renderer failure";
                    StringAssert.Contains(response.Error, expected);
                    Assert.AreEqual(originalSource, editor.Native.Adapter.Read(), "No compilation guard may overwrite the original source.");
                    if (outcome == "draft") Assert.AreEqual("unsynchronized draft", editor.Document.Text);
                    Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                }
                finally { SynchronizationContext.SetSynchronizationContext(previous); }
            }
        }

        /// <summary>Checks one deferred save, actual awaited native completion and approval changes before dispatch.</summary>
        /// <param name="baseContext">Uses a context whose default Post would otherwise use a pool thread.</param>
        /// <param name="catalog">Uses the catalogue gateway instead of direct dispatch.</param>
        /// <param name="denyAfterYield">Revokes write approval before the deferred save enters the host.</param>
        [WinFormsTestMethod]
        [DataRow(false, false, false), DataRow(false, true, false)]
        [DataRow(true, false, false), DataRow(true, true, false)]
        [DataRow(false, false, true), DataRow(false, true, true)]
        [DataRow(true, false, true), DataRow(true, true, true)]
        public void DeferredSavePreservesOwnerAndRechecksApproval(bool baseContext, bool catalog, bool denyAfterYield)
        {
            int owner = Thread.CurrentThread.ManagedThreadId, saves = 0;
            var fixture = new ToolFixture();
            var execute = fixture.Tools.Execute;
            fixture.Tools.Execute = request =>
            {
                Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                return execute(request);
            };
            fixture.Tools.SaveHostDocumentNative = async request =>
            {
                Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                Assert.AreEqual("P", request.Project);
                saves++;
                await Task.Delay(2);
                Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                return new { Saved = true };
            };
            string arguments = Json.Serialize(new { Project = "P", ExpectedProjectVersion = "version", ExpectedHostPath = @"C:\fixture\Owned.swp" });
            if (catalog) arguments = Json.Serialize(new { ToolName = "save_host_document", ArgumentsJson = arguments });
            var previous = SynchronizationContext.Current;
            var ambient = baseContext ? new SynchronizationContext() : null;
            try
            {
                SynchronizationContext.SetSynchronizationContext(ambient);
                var pending = fixture.Tools.InvokeAsync(catalog ? "invoke_tool" : "save_host_document", arguments);
                Assert.AreSame(ambient, SynchronizationContext.Current, "Caller context must be restored before pumping.");
                Assert.IsFalse(pending.IsCompleted);
                Assert.AreEqual(0, saves);
                if (denyAfterYield) fixture.Settings.VbeEditApproval = "ReadOnly";
                var response = Json.Deserialize<Response>(ModernEditorDebugFixture.Wait(pending));
                Assert.AreEqual(!denyAfterYield, response.Ok, response.Error);
                Assert.AreEqual(denyAfterYield ? 0 : 1, saves, "An uncertain or refused save must never be replayed.");
                Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
            }
            finally { SynchronizationContext.SetSynchronizationContext(previous); }
        }

        /// <summary>Retains the existing host-free MTA path without creating or relocating native work.</summary>
        /// <param name="tool">Read-only tool executed against the explicit managed boundary.</param>
        [DataTestMethod]
        [DataRow("read_module"), DataRow("list_modules")]
        public async Task ManagedMtaToolBoundaryRetainsItsCaller(string tool)
        {
            await Task.Run(() =>
            {
                Assert.AreEqual(ApartmentState.MTA, Thread.CurrentThread.GetApartmentState());
                int owner = Thread.CurrentThread.ManagedThreadId;
                var ambient = SynchronizationContext.Current;
                var tools = new LlmVbeTools(null, null, new LlmSettings { VbeEditApproval = "Automatic" });
                tools.MonacoWindow = create => { Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId); return null; };
                tools.Execute = request =>
                {
                    Assert.AreEqual(owner, Thread.CurrentThread.ManagedThreadId);
                    Assert.AreEqual(tool, request.Command);
                    return Response.Success(new { ManagedBoundary = true });
                };
                var values = new Dictionary<string, object> { ["Project"] = "P" };
                if (tool == "read_module") values["Module"] = "M";
                var pending = tools.InvokeAsync(tool, Json.Serialize(values));
                Assert.IsTrue(pending.IsCompleted);
                Assert.IsTrue(Json.Deserialize<Response>(pending.GetAwaiter().GetResult()).Ok);
                Assert.AreSame(ambient, SynchronizationContext.Current);
            });
        }
    }
}
