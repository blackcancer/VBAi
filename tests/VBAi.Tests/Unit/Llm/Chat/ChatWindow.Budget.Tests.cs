using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using VBAi;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace VBAi.Tests.Unit
{
    public sealed partial class ChatWindowStateTests
    {
        [STATestMethod, TestCategory("Unit")]
        public void HttpSafetyPauseCountsIdenticalResultsNullAndMalformedResponsesAsStalledRounds()
        {
            foreach (string result in new[] { "null", "malformed", "{\"Ok\":true,\"Data\":\"same\"}" })
            using (var runtime = new RuntimeScope())
            using (var window = ReadyHttpWindow(new ChatSessionState { Scope = "temporary:test", Provider = "Ollama" }))
            {
                int count = result.Contains("same") ? 9 : 8; int calls = 0; var json = new JavaScriptSerializer();
                var replies = Enumerable.Range(0, count).Select(i => json.Serialize(new { choices = new[] { new { message = new { role = "assistant", tool_calls = new[] { new { id = "stall-" + i, type = "function", function = new { name = "status", arguments = "{}" } } } } } } })).ToArray();
                var handler = new ChatResponseHandler(replies); window.HttpHandlerOverride = () => handler;
                ChatWindow.InvokeTool = (tools, name, args) => { calls++; return Task.FromResult(result); };
                Question(window, "bounded repeated work"); CompleteOnSta((Task)Call(window, "SendAsync"));
                var state = Get<ChatSessionState>(window, "currentSession");
                Assert.AreEqual(count, calls); Assert.AreEqual(count, handler.Requests.Count);
                Assert.IsTrue(state.BudgetPaused); Assert.IsFalse(Get<bool>(window, "busy"));
                Assert.AreEqual(count, state.CompletedToolActions.Count); Set(window, "currentSession", null);
            }
        }

        [STATestMethod, TestCategory("Unit")]
        public void EightRoundsWithoutProgressPauseAndResumeKeepsResultsWithoutReplayingActions()
        {
            using (var runtime = new RuntimeScope())
            using (var window = ReadyHttpWindow(new ChatSessionState { Scope = "temporary:test", Provider = "Ollama" }))
            {
                int actions = 0; var json = new JavaScriptSerializer();
                var replies = Enumerable.Range(0, 8).Select(i => json.Serialize(new { choices = new[] { new { message = new { role = "assistant", tool_calls = new[] { new { id = "action-" + i, type = "function", function = new { name = "status", arguments = "{}" } } } } } } })).ToArray();
                var first = new ChatResponseHandler(replies);
                window.HttpHandlerOverride = () => first;
                ChatWindow.InvokeTool = (t, n, a) => { actions++; return Task.FromResult(json.Serialize(Response.Failure("No progress"))); };
                Question(window, "perform work"); CompleteOnSta((Task)Call(window, "SendAsync"));
                var state = Get<ChatSessionState>(window, "currentSession");
                Assert.IsTrue(state.BudgetPaused); Assert.AreEqual(8, actions);
                string pausedTurn = state.PausedTurnId;
                Assert.AreEqual(8, state.CompletedToolActions.Count);
                Assert.IsFalse(Get<bool>(window, "busy"));
                var history = Get<List<object>>(window, "messages");
                Assert.AreEqual(8, history.Count(m => json.Serialize(m).Contains("\"role\":\"tool\"")));
                var restored = json.Deserialize<ChatSessionState>(json.Serialize(state));
                Assert.IsTrue(restored.BudgetPaused); Assert.AreEqual(state.PausedTurnId, restored.PausedTurnId);
                var resume = new ChatResponseHandler(replies[7], "{\"choices\":[{\"message\":{\"role\":\"assistant\",\"content\":\"done\"}}]}");
                window.HttpHandlerOverride = () => resume;
                var mode = state.Mode; state.Mode = ChatMode.Discussion;
                CompleteOnSta((Task)Call(window, "ResumeBudgetAsync")); Assert.IsTrue(state.BudgetPaused); Assert.AreEqual(0, resume.Requests.Count);
                state.Mode = mode;
                CompleteOnSta((Task)Call(window, "ResumeBudgetAsync"));
                Assert.IsFalse(state.BudgetPaused); Assert.AreEqual(8, actions);
                Assert.AreEqual(2, resume.Requests.Count);
                StringAssert.Contains(resume.Requests[0], "action-7");
                StringAssert.Contains(resume.Requests[1], "No action was replayed");
                Assert.IsTrue(Get<List<ChatEntry>>(window, "transcriptEntries").Any(e => e.Text == "done" && e.TurnId == pausedTurn));
                Assert.AreEqual(1, history.Count(m => json.Serialize(m).Contains("\"role\":\"user\"")));
                Set(window, "currentSession", null);
            }
        }
    }
}
namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using System.Windows.Forms;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    public sealed partial class ChatWindowStateTests
    {
        [STATestMethod,TestCategory("Unit")]
        public void BudgetToolLabelsRecordExactSuccessRefusalMalformedAndInterruptedOutcomes()
        {
            using(var runtime=new RuntimeScope())
            using(var window=ReadyHttpWindow(new ChatSessionState())) {
                var state=Get<ChatSessionState>(window,"currentSession");
                foreach(var args in new[]{"null","{}","{\"ToolName\":\"read_module\"}"})
                foreach(var result in new[]{"{\"Ok\":true}","{\"Ok\":false}","null","malformed"}) {
                    state.CompletedToolActions=null;
                    ChatWindow.InvokeTool=(t,n,a)=>Task.FromResult(result);
                    var task=(Task<string>)Call(window,"ExecuteBudgetTool","invoke_tool",args);CompleteOnSta(task);
                    Assert.AreEqual(result,task.Result);
                    StringAssert.Contains(state.CompletedToolActions.Single(),args.Contains("read_module")?"read_module":"invoke_tool");
                    StringAssert.Contains(state.CompletedToolActions.Single(),UiText.Get(result=="{\"Ok\":true}"?" — response received; inspect returned state":" — refused or failed"));
                }
                foreach(bool existing in new[]{false,true}) {
                    state.CompletedToolActions=existing?new List<string>():null;
                    ChatWindow.InvokeTool=(t,n,a)=>throw new InvalidOperationException("native action failed");
                    var task=(Task<string>)Call(window,"ExecuteBudgetTool","status","{}");
                    Assert.ThrowsException<InvalidOperationException>(()=>CompleteOnSta(task));
                    StringAssert.Contains(state.CompletedToolActions.Single(),UiText.Get(" — interrupted; verify live state before retrying"));
                }
                Set(window,"currentSession",null);
            }
        }

        [STATestMethod,TestCategory("Unit")]
        public void BudgetPauseAndButtonsRespectBusyDraftSessionAndAppliedTurnChanges()
        {
            using(var runtime=new RuntimeScope())
            using(var window=ReadyHttpWindow(new ChatSessionState{Mode=ChatMode.Agent})) {
                var state=Get<ChatSessionState>(window,"currentSession");
                Set(window,"activeTurnId","turn");
                var effort=Get<ComboBox>(window,"effortPicker");effort.Items.Add(new LlmEffortOption("high","High"));effort.SelectedIndex=0;
                foreach(bool verification in new[]{false,true}) {
                    Get<CheckBox>(window,"verifyAfterEdit").Checked=verification;
                    Get<List<CodeChange>>(window,"codeChanges").Add(new CodeChange{TurnId="turn",Project="P",Module="M",Before="",After=""});
                    state.CompletedToolActions=null;
                    Call(window,"PauseBudget",LlmProvider.All[2],"local-test");
                    Assert.IsTrue(state.BudgetPaused);Assert.AreEqual("turn",state.PausedTurnId);Assert.AreEqual("high",state.PausedEffort);
                    StringAssert.Contains(Get<List<ChatEntry>>(window,"transcriptEntries").Last().Text,UiText.Get(verification?"Pending: automatic verification and final response.":"Pending: final response."));
                }
                var send=Get<Button>(window,"send");var resume=Get<ToolStripMenuItem>(window,"resumeTurn");
                foreach(bool busy in new[]{true,false}) foreach(bool paused in new[]{true,false}) foreach(string draft in new[]{"","draft"}) {
                    Set(window,"busy",busy);state.BudgetPaused=paused;Question(window,draft);Call(window,"UpdateBudgetControls");
                    Assert.AreEqual(!busy&&paused,resume.Enabled);
                    if(!busy) Assert.AreEqual(UiText.Get(paused&&draft==""?"Resume ▶":"Send ↑"),send.Text);
                }
                Set(window,"currentSession",null);Set(window,"busy",false);Call(window,"UpdateBudgetControls");Assert.IsFalse(resume.Enabled);
                Set(window,"resumeTurn",null);Call(window,"UpdateBudgetControls");Set(window,"resumeTurn",resume);
                Set(window,"send",null);Call(window,"UpdateBudgetControls");Set(window,"send",send);
                var raw=new ChatWindow();Call(raw,"UpdateBudgetControls");raw.Dispose();
            }
        }

        [STATestMethod,TestCategory("Unit")]
        public void SendAndResumeControlsRespectScopeDuringDraftBusyPauseAndStopTransitions()
        {
            foreach (string scopeState in new[] { "available", "closed", "loading" })
            using (var runtime = new RuntimeScope())
            using (var window = LoadedWindow(runtime.Session))
            {
                if (scopeState == "closed")
                {
                    runtime.Host = r => Response.Success(new object[0]);
                    Assert.AreEqual(true, Call(window, "RefreshAvailableScopes", runtime.Session));
                    Assert.AreEqual(-1, Get<ComboBox>(window, "scopePicker").SelectedIndex);
                }
                else
                {
                    Assert.IsNotNull(Get<ComboBox>(window, "scopePicker").SelectedItem);
                }
                Set(window, "loadingScope", scopeState == "loading");
                var state = Get<ChatSessionState>(window, "currentSession");
                var send = Get<Button>(window, "send");
                var resume = Get<ToolStripMenuItem>(window, "resumeTurn");
                bool available = scopeState == "available";
                foreach (bool paused in new[] { false, true })
                foreach (bool stopping in new[] { false, true })
                foreach (bool busy in new[] { false, true })
                foreach (string draft in new[] { "", "Owned synthetic draft", " \r\n" })
                {
                    string scenario = scopeState + "; paused=" + paused + "; stopping=" + stopping +
                        "; busy=" + busy + "; draft=" + draft;
                    state.BudgetPaused = paused;
                    Set(window, "stopRequested", stopping);
                    Question(window, "");
                    Call(window, "SetBusy", busy);
                    Assert.AreEqual(busy ? !stopping : available, send.Enabled,
                        "An empty composer must preserve Stop and scope gating after a busy change: " + scenario);
                    Question(window, draft);
                    bool hasText = !string.IsNullOrWhiteSpace(draft);
                    Assert.AreEqual(busy ? (hasText ? available : !stopping) : available, send.Enabled,
                        "A real draft event must preserve scope gating and cancellation: " + scenario);
                    Assert.AreEqual(available && !busy && paused, resume.Enabled,
                        "Resume must require the selected, loaded scope: " + scenario);
                    Assert.AreEqual(UiText.Get(busy ? (hasText ? "Queue ↑" : "Stop ■") :
                        paused && !hasText ? "Resume ▶" : "Send ↑"), send.Text, scenario);
                }
                Set(window, "loadingScope", false);
                Set(window, "stopRequested", false);
                Call(window, "SetBusy", false);
            }
        }

        [STATestMethod,TestCategory("Unit")]
        public void BudgetResumeRejectsEachChangedProviderModelModeEffortAndClosedScope()
        {
            using(var runtime=new RuntimeScope()) {
                foreach(int scenario in Enumerable.Range(0,11)) using(var window=ReadyHttpWindow(new ChatSessionState {
                    BudgetPaused=true,PausedProvider="Ollama",PausedModel="local-test",PausedMode=ChatMode.Agent,Mode=ChatMode.Agent })) {
                    var state=Get<ChatSessionState>(window,"currentSession");var providers=Get<ComboBox>(window,"providerPicker");var models=Get<ComboBox>(window,"modelPicker");
                    if(scenario==0) Set(window,"busy",true);
                    if(scenario==1) Set(window,"currentSession",null);
                    if(scenario==2) state.BudgetPaused=false;
                    if(scenario==3) providers.SelectedIndex=-1;
                    if(scenario==4) models.SelectedIndex=-1;
                    if(scenario==5) {providers.Items.Add(LlmProvider.All[0]);providers.SelectedIndex=1;}
                    if(scenario==6) state.PausedProvider="different";
                    if(scenario==7) state.PausedModel="different";
                    if(scenario==8) state.PausedMode=ChatMode.Discussion;
                    if(scenario==9) state.PausedEffort="high";
                    if(scenario==10) Set(window,"scopeSession",runtime.Session);
                    var handler=new RuntimeHttpHandler();window.HttpHandlerOverride=()=>handler;
                    CompleteOnSta((Task)Call(window,"ResumeBudgetAsync"));
                    Assert.IsNull(Get<LlmChatClient>(window,"activeHttpClient"));
                    Set(window,"busy",false);Set(window,"currentSession",null);
                }
            }
        }

        [STATestMethod,TestCategory("Unit")]
        public void PendingBudgetResponsesDeduplicateIdsAndIgnoreMalformedToolRecords()
        {
            using(var window=ReadyHttpWindow(new ChatSessionState())) {
                var records=Get<List<object>>(window,"messages");
                records.Add(null);records.Add(3);records.Add(new{});records.Add(new{tool_calls="bad"});
                records.Add(new{tool_calls=new object[]{null,4,new{},new{id="existing"},new{id="new"},new{id="new"}}});
                records.Add(new{role="tool",tool_call_id="existing",content="saved"});
                Call(window,"CompletePendingToolResponses");Call(window,"CompletePendingToolResponses");
                Assert.AreEqual(7,records.Count);
                StringAssert.Contains(new JavaScriptSerializer().Serialize(records.Last()),"Execution is unconfirmed");
                Set(window,"currentSession",null);
            }
        }

        [STATestMethod,TestCategory("Unit")]
        public void HttpBudgetRejectsMalformedToolCallsAndStopsWithoutRepeatingActions()
        {
            using(var runtime=new RuntimeScope()) {
                foreach(var calls in new object[]{new object[]{null},new object[]{new{id="missing-function"}},new object[]{new{function=new{name="status",arguments="{}"}}}}) {
                    using(var window=ReadyHttpWindow(new ChatSessionState())) {
                        var handler=new RuntimeHttpHandler{Body=new JavaScriptSerializer().Serialize(new{choices=new[]{new{message=new{role="assistant",tool_calls=calls}}}})};
                        window.HttpHandlerOverride=()=>handler;
                        var task=(Task<bool>)Call(window,"RunHttpBudgetAsync",LlmProvider.All[2],"local-test");
                        Assert.ThrowsException<InvalidOperationException>(()=>CompleteOnSta(task));
                        Set(window,"currentSession",null);
                    }
                }
                foreach(int stopAt in new[]{0,1,2,3}) using(var window=ReadyHttpWindow(new ChatSessionState())) {
                    var handler=new RuntimeHttpHandler();
                    if(stopAt==0) Set(window,"stopRequested",true);
                    if(stopAt==1) handler.BeforeResponse=()=>Set(window,"stopRequested",true);
                    if(stopAt==2) {handler.Body="{\"choices\":[{\"message\":{\"tool_calls\":[{\"id\":\"action\",\"function\":{\"name\":\"status\",\"arguments\":\"{}\"}}]}}]}"; ChatWindow.InvokeTool=(t,n,a)=>{Set(window,"stopRequested",true);return Task.FromResult("{\"Ok\":true}");};}
                    if(stopAt==3) handler.BeforeResponse=()=>{Set(window,"stopRequested",true);Get<LlmChatClient>(window,"activeHttpClient").ToolHandler("status","{}").GetAwaiter().GetResult();};
                    window.HttpHandlerOverride=()=>handler;
                    var task=(Task<bool>)Call(window,"RunHttpBudgetAsync",LlmProvider.All[2],"local-test");
                    Assert.ThrowsException<OperationCanceledException>(()=>CompleteOnSta(task));
                    Set(window,"currentSession",null);
                }
            }
        }

        [STATestMethod,TestCategory("Unit")]
        public void BudgetResumeRebuildsSystemHistoryPreservesErrorsAndReportsAppliedChanges()
        {
            using(var runtime=new RuntimeScope()) {
                foreach(var first in new object[]{null,3,new{},new{role="user",content="request"},new{role="system",content="old"}}) {
                    using(var window=ReadyHttpWindow(new ChatSessionState{BudgetPaused=true,PausedProvider="Ollama",PausedModel="local-test",PausedMode=ChatMode.Agent,Mode=ChatMode.Agent,PausedTurnId="turn"})) {
                        Get<List<object>>(window,"messages").Add(first);
                        var handler=new RuntimeHttpHandler();window.HttpHandlerOverride=()=>handler;
                        CompleteOnSta((Task)Call(window,"ResumeBudgetAsync"));
                        Assert.IsFalse(Get<ChatSessionState>(window,"currentSession").BudgetPaused);
                        StringAssert.Contains(new JavaScriptSerializer().Serialize(Get<List<object>>(window,"messages")[0]),"system");
                        Set(window,"currentSession",null);
                    }
                }
                foreach(bool fail in new[]{false,true}) using(var window=ReadyHttpWindow(new ChatSessionState{BudgetPaused=true,PausedProvider="Ollama",PausedModel="local-test",PausedMode=ChatMode.Agent,Mode=ChatMode.Agent,PausedTurnId="turn"})) {
                    var changes=Get<List<CodeChange>>(window,"codeChanges");changes.Add(new CodeChange{TurnId="other",Before="",After=""});
                    Get<CheckBox>(window,"verifyAfterEdit").Checked=true;
                    var handler=new RuntimeHttpHandler{BeforeResponse=()=>{changes.Add(new CodeChange{TurnId="turn",Before="",After=""});if(fail) throw new InvalidOperationException("provider failure");}};
                    window.HttpHandlerOverride=()=>handler;
                    CompleteOnSta((Task)Call(window,"ResumeBudgetAsync"));
                    Assert.AreEqual(fail,Get<ChatSessionState>(window,"currentSession").BudgetPaused);
                    Assert.IsTrue(Get<List<ChatEntry>>(window,"transcriptEntries").Any(entry=>entry.Speaker=="Intervention"));
                    if(fail) Assert.IsTrue(Get<List<ChatEntry>>(window,"transcriptEntries").Any(entry=>entry.Text.Contains("provider failure")));
                    Set(window,"currentSession",null);
                }
            }
        }

        [STATestMethod,TestCategory("Unit")]
        public void BudgetHttpUsesInstalledProtocolCallbacksAndEmptyAnswersWithoutNetwork()
        {
            using(var runtime=new RuntimeScope())
            using(var window=ReadyHttpWindow(new ChatSessionState())) {
                var handler=new RuntimeHttpHandler{Body="{\"choices\":[{\"message\":{\"content\":\"\",\"tool_calls\":[]}}]}",BeforeResponse=()=>{
                    var client=Get<LlmChatClient>(window,"activeHttpClient");
                    client.ToolHandler("status","{}").GetAwaiter().GetResult();
                    client.TextDelta("partial");
                }};
                window.HttpHandlerOverride=()=>handler;
                var task=(Task<bool>)Call(window,"RunHttpBudgetAsync",LlmProvider.All[2],"local-test");CompleteOnSta(task);
                Assert.IsTrue(task.Result);
                Set(window,"stopRequested",true);Get<LlmChatClient>(window,"activeHttpClient").TextDelta("ignored");
                var completedClient = Get<LlmChatClient>(window,"activeHttpClient");
                Set(window,"stopRequested",false);window.Dispose();completedClient.TextDelta("ignored disposed");
                Set(window,"currentSession",null);
            }
        }
    }
}
namespace VBAi.Tests.Unit
{
    public sealed partial class ChatWindowStateTests
    {
        [STATestMethod,TestCategory("Unit")]
        public void ResumeAcceptsMatchingEffortAndSkipsUiCleanupAfterOwnedWindowDisposal()
        {
            using(var runtime=new RuntimeScope()) foreach(bool dispose in new[]{false,true})
            using(var window=ReadyHttpWindow(new ChatSessionState{BudgetPaused=true,PausedProvider="Ollama",PausedModel="local-test",PausedMode=ChatMode.Agent,Mode=ChatMode.Agent,PausedTurnId="turn",PausedEffort="high"})) {
                var effort=Get<System.Windows.Forms.ComboBox>(window,"effortPicker");
                effort.Items.Add(new LlmEffortOption("high","High"));effort.SelectedIndex=0;
                var handler=new RuntimeHttpHandler{BeforeResponse=()=>{if(dispose)window.Dispose();}};
                window.HttpHandlerOverride=()=>handler;
                CompleteOnSta((Task)Call(window,"ResumeBudgetAsync"));
                Assert.AreEqual(dispose,window.IsDisposed);
                if(!dispose) Assert.IsFalse(Get<bool>(window,"busy"));
                Set(window,"currentSession",null);
            }
        }
    }
}
namespace VBAi.Tests.Unit
{
    public sealed partial class ChatWindowStateTests
    {
        [STATestMethod,TestCategory("Unit")]
        public void ResumeWaitsForAsynchronousVerificationBeforeFinishingTheTurn()
        {
            using(var runtime=new RuntimeScope())
            using(var window=ReadyHttpWindow(new ChatSessionState{Scope=@"C:\Temp\P.xlsm",BudgetPaused=true,PausedProvider="Ollama",PausedModel="local-test",PausedMode=ChatMode.Agent,Mode=ChatMode.Agent,PausedTurnId="turn"})) {
                AddScope(window,@"C:\Temp\P.xlsm");Get<System.Windows.Forms.ComboBox>(window,"scopePicker").SelectedIndex=0;
                Set(window,"scopeSession",runtime.Session);Get<System.Windows.Forms.CheckBox>(window,"verifyAfterEdit").Checked=true;
                int verified=0;
                ChatWindow.InvokeTool=async(tools,name,arguments)=>{
                    Assert.AreEqual("compile_project",name);await Task.Delay(20);verified++;
                    return "{\"Ok\":true,\"Data\":{\"Compiled\":true}}";
                };
                var handler=new RuntimeHttpHandler{BeforeResponse=()=>Get<List<CodeChange>>(window,"codeChanges").Add(new CodeChange{TurnId="turn",Project="P",Module="M",Before="",After=""})};
                window.HttpHandlerOverride=()=>handler;
                var task=(Task)Call(window,"ResumeBudgetAsync");
                Assert.IsFalse(task.IsCompleted);Assert.IsTrue(Get<bool>(window,"busy"));
                CompleteOnSta(task);
                Assert.AreEqual(1,verified);Assert.IsFalse(Get<bool>(window,"busy"));
                Assert.IsTrue(Get<List<ChatEntry>>(window,"transcriptEntries").Any(entry=>entry.Speaker=="Vérification"&&entry.Text.Contains("no native diagnostics")));
                Set(window,"currentSession",null);
            }
        }
    }
}
namespace VBAi.Tests.Unit
{
    public sealed partial class ChatWindowStateTests
    {
        [STATestMethod,TestCategory("Unit")]
        public void ResumePropagatesOversizedSavedHistoryFailureAndReleasesBusyClient()
        {
            using(var runtime=new RuntimeScope())
            using(var window=ReadyHttpWindow(new ChatSessionState{BudgetPaused=true,PausedProvider="Ollama",PausedModel="local-test",PausedMode=ChatMode.Agent,Mode=ChatMode.Agent,PausedTurnId="turn"})) {
                var messages=Get<List<object>>(window,"messages");
                messages.Add(new{role="system",content="instructions"});
                messages.Add(new{role="user",content=new string('x',10*1024*1024+1)});
                int actions=0;ChatWindow.InvokeTool=(tools,name,arguments)=>{actions++;return Task.FromResult("{\"Ok\":true}");};
                window.HttpHandlerOverride=()=>new RuntimeHttpHandler();
                var task=(Task)Call(window,"ResumeBudgetAsync");
                Assert.ThrowsException<InvalidOperationException>(()=>CompleteOnSta(task));
                Assert.IsTrue(task.IsFaulted);Assert.AreEqual(0,actions);
                Assert.IsFalse(Get<bool>(window,"busy"));Assert.IsNull(Get<LlmChatClient>(window,"activeHttpClient"));
                Assert.AreEqual(2,messages.Count);
                Set(window,"currentSession",null);
            }
        }
    }
}

namespace VBAi.Tests.Unit
{
    public sealed partial class ChatWindowStateTests
    {
        [STATestMethod,TestCategory("Unit")]
        public void ResumePropagatesRecoveryFailureWhenTheWindowClosesDuringAProviderReply()
        {
            using(var runtime=new RuntimeScope())
            using(var window=ReadyHttpWindow(new ChatSessionState{BudgetPaused=true,PausedProvider="Ollama",PausedModel="local-test",PausedMode=ChatMode.Agent,Mode=ChatMode.Agent,PausedTurnId="turn"})) {
                var messages=Get<List<object>>(window,"messages");
                messages.Add(new{role="user",content="continue"});
                var handler=new RuntimeHttpHandler{BeforeResponse=()=>{
                    window.Dispose();
                    messages.Add(new{role="assistant",content=new string('x',10*1024*1024+1)});
                    throw new System.IO.IOException("provider reply interrupted after closure");
                }};
                window.HttpHandlerOverride=()=>handler;
                var task=(Task)Call(window,"ResumeBudgetAsync");
                var error=Assert.ThrowsException<InvalidOperationException>(()=>CompleteOnSta(task));
                StringAssert.Contains(error.StackTrace,"CompletePendingToolResponses");
                Assert.IsTrue(task.IsFaulted);Assert.IsTrue(window.IsDisposed);
                Assert.IsNull(Get<LlmChatClient>(window,"activeHttpClient"));
                Assert.AreEqual(3,messages.Count);
                Set(window,"currentSession",null);
            }
        }
    }
}
