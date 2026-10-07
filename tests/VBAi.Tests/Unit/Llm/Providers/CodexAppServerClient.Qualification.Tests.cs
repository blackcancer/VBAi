namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Threading.Tasks;
    using System.Web.Script.Serialization;
    using VBAi;

    public sealed partial class CodexAppServerClientTests
    {
        [TestMethod]
        public async Task QualificationReceiptsObserveActualDispatchResultAndPreAdmissionRefusalWithoutChangingTools()
        {
            foreach(bool brokenSink in new[]{false,true})
            {
                var now=DateTime.UtcNow;var lines=new List<string>();int invocations=0;
                var trace=new CodexToolQualificationTrace(CodexToolQualificationTraceTests.Manifest(now),()=>now,line=> {if(brokenSink)throw new IOException("diagnostic unavailable");lines.Add(line);});
                var transport=new FakeTransport();var settings=new LlmSettings();var tools=new LlmVbeTools(null,null,settings){BoundProject="SyntheticProject"};
                using(var client=new CodexAppServerClient(new ImmediateContext(),tools,null,settings,null,transport))
                {
                    client.QualificationTrace=trace;
                    client.InvokeTool=(name,args)=>{invocations++;return Task.FromResult(invocations==1 ? "{\"Ok\":true}" : "{\"Ok\":false}");};
                    await client.ListModelsAsync();
                    var attempted=typeof(CodexAppServerClient).GetField("qualificationTraceAttempted",System.Reflection.BindingFlags.Instance|System.Reflection.BindingFlags.NonPublic);
                    Assert.AreEqual(false,attempted.GetValue(client));
                    var turn=client.TurnAsync("synthetic request",null,null);await transport.TurnStarted.Task;
                    Assert.AreEqual(true,attempted.GetValue(client));
                    foreach(string module in new[]{"SyntheticExpected","SyntheticMissing"})
                    {
                        var parameters=CodexToolQualificationTraceTests.Parameters(module);parameters["threadId"]="thread-1";parameters["turnId"]="turn-1";
                        transport.Emit(new {id="same-rpc-id",method="item/tool/call",@params=parameters});
                    }
                    var wrong=CodexToolQualificationTraceTests.Parameters();wrong["threadId"]="other-thread";
                    transport.Emit(new{id="wrong-thread",method="item/tool/call",@params=wrong});
                    Assert.AreEqual(2,invocations);var replies=transport.Sent.Where(r=>!r.ContainsKey("method")&&r.ContainsKey("result")).ToArray();
                    Assert.AreEqual(true,FakeTransport.Object(replies[replies.Length-3]["result"])["success"]);
                    Assert.AreEqual(false,FakeTransport.Object(replies[replies.Length-2]["result"])["success"]);
                    Assert.AreEqual(false,FakeTransport.Object(replies.Last()["result"])["success"]);
                    transport.EmitTurnCompleted("completed",null);await turn;
                }
                await trace.PendingWrites;
                if(!brokenSink)
                {
                    var records=lines.Select(line=>new JavaScriptSerializer().DeserializeObject(line) as Dictionary<string,object>).ToArray();
                    Assert.AreEqual(3,records.Count(r=>Equals(r["Stage"],"Received")));
                    Assert.AreEqual(2,records.Count(r=>Equals(r["Stage"],"Admitted")));
                    Assert.AreEqual(1,records.Count(r=>Equals(r["Stage"],"Returned")&&Equals(r["Ok"],true)));
                    Assert.AreEqual(1,records.Count(r=>Equals(r["Stage"],"Returned")&&Equals(r["Ok"],false)));
                    Assert.AreEqual(1,records.Count(r=>Equals(r["Stage"],"RejectedBeforeAdmission")));
                }
            }
        }
        [TestMethod]
        public async Task OldTurnToolQueuedAfterStopIsRejectedWhenSameThreadResumes()
        {
            var queue=new VBAi.Tests.Infrastructure.LlmQueuedContext();var transport=new FakeTransport();var settings=new LlmSettings();int calls=0;
            using(var client=new CodexAppServerClient(queue,new LlmVbeTools(null,null,settings),null,settings,null,transport))
            {
                client.InvokeTool=(name,args)=>{calls++;return Task.FromResult("{\"Ok\":true}");};
                var stopped=client.TurnAsync("old synthetic request",null,null);await transport.TurnStarted.Task;
                await client.InterruptAsync();queue.Drain();await Assert.ThrowsExceptionAsync<OperationCanceledException>(()=>stopped);
                // The old request arrives while idle; its callback is still waiting on the owner context.
                transport.Emit(new{id="old-request",method="item/tool/call",@params=new{threadId="thread-1",turnId="turn-1",callId="old-call",tool="read_module",arguments=new{}}});
                var started=new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
                transport.Intercept=message=>
                {
                    if(Method(message)!="turn/start")return false;
                    transport.Emit(new{method="turn/started",@params=new{threadId="thread-1",turn=new{id="turn-2"}}});
                    transport.Emit(new{id=message["id"],result=new{turn=new{id="turn-2"}}});started.TrySetResult(true);return true;
                };
                var resumed=client.TurnAsync("new synthetic request",null,null);await started.Task;
                queue.Drain();
                var reply=transport.Sent.Last(r=>!r.ContainsKey("method")&&Equals(r["id"],"old-request"));
                transport.Emit(new{method="turn/completed",@params=new{threadId="thread-1",turn=new{id="turn-2",status="completed"}}});queue.Drain();await resumed;
                Assert.AreEqual(0,calls,"A callback from the stopped turn must never enter InvokeTool in the new turn.");
                Assert.AreEqual(false,FakeTransport.Object(reply["result"])["success"]);
            }
        }
        [TestMethod]
        public async Task OldTurnIdsAndMissingIdsCannotDispatchCompleteOrResetActiveTurn()
        {
            var transport=new FakeTransport();int calls=0;using(var client=Client(transport))
            {
                client.InvokeTool=(name,args)=>{calls++;return Task.FromResult("{\"Ok\":true}");};
                var turn=client.TurnAsync("synthetic request",null,null);await transport.TurnStarted.Task;
                transport.Emit(new{id="old",method="item/tool/call",@params=new{threadId="thread-1",turnId="old-turn",tool="read_module",arguments=new{}}});
                transport.Emit(new{method="turn/completed",@params=new{threadId="thread-1",turn=new{id="old-turn",status="completed"}}});
                Assert.IsFalse(turn.IsCompleted);Assert.AreEqual(0,calls);
                transport.Emit(new{method="turn/started",@params=new{threadId="thread-1",turn=new{id="old-turn"}}});
                transport.Emit(new{id="current",method="item/tool/call",@params=new{threadId="thread-1",turnId="turn-1",tool="read_module",arguments=new{}}});
                transport.Emit(new{id="legacy",method="item/tool/call",@params=new{threadId="thread-1",tool="read_module",arguments=new{}}});
                Assert.AreEqual(1,calls);Assert.AreEqual(false,FakeTransport.Object(transport.Sent.Last(r=>!r.ContainsKey("method")&&Equals(r["id"],"legacy"))["result"])["success"]);transport.EmitTurnCompleted("completed",null);await turn;
            }
        }
        [TestMethod]
        public async Task LateToolCompletionCannotPublishIntoOrFailResumedTurn()
        {
            foreach(bool brokenReply in new[]{false,true})
            {
                var transport=new FakeTransport();var output=new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);
                var reply=new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);var newStarted=new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
                var updates=new List<string>();var activities=new List<string>();
                using(var client=Client(transport))
                {
                    client.ChatUpdate+=(kind,id,text,complete)=>updates.Add(id);client.ActivityUpdate+=activity=>activities.Add(activity.Id);
                    client.InvokeTool=(name,args)=>output.Task;
                    var old=client.TurnAsync("old request",null,null);await transport.TurnStarted.Task;
                    transport.Emit(new{id="pending-old",method="item/tool/call",@params=new{threadId="thread-1",turnId="turn-1",callId="old-call",tool="read_module",arguments=new{}}});
                    await client.InterruptAsync();await Assert.ThrowsExceptionAsync<OperationCanceledException>(()=>old);
                    transport.Intercept=message=>
                    {
                        if(Method(message)=="turn/start")
                        {
                            transport.Emit(new{method="turn/started",@params=new{threadId="thread-1",turn=new{id="turn-2"}}});
                            transport.Emit(new{id=message["id"],result=new{turn=new{id="turn-2"}}});newStarted.TrySetResult(true);return true;
                        }
                        if(Method(message)==null&&Equals(message["id"],"pending-old"))
                        {
                            reply.TrySetResult(true);if(brokenReply)throw new IOException("old request transport failure");
                        }
                        return false;
                    };
                    var current=client.TurnAsync("new request",null,null);await newStarted.Task;updates.Clear();activities.Clear();
                    output.TrySetResult("{\"Ok\":true}");await reply.Task;
                    Assert.IsFalse(current.IsCompleted);Assert.AreEqual(0,updates.Count);Assert.AreEqual(0,activities.Count);
                    if(!brokenReply){var response=transport.Sent.Last(r=>!r.ContainsKey("method")&&Equals(r["id"],"pending-old"));Assert.AreEqual(true,FakeTransport.Object(response["result"])["success"]);}
                    transport.Emit(new{method="turn/completed",@params=new{threadId="thread-1",turn=new{id="turn-2",status="completed"}}});await current;
                }
            }
        }
        private sealed class ToolOwnerQueueContext : System.Threading.SynchronizationContext
        {
            private readonly Queue<Action> work=new Queue<Action>();
            public override void Post(System.Threading.SendOrPostCallback callback,object state){lock(work)work.Enqueue(()=>callback(state));}
            internal void Drain()
            {
                while(true){Action action;lock(work){if(work.Count==0)return;action=work.Dequeue();}
                    var previous=Current;SetSynchronizationContext(this);try{action();}finally{SetSynchronizationContext(previous);}}
            }
        }
        private static async Task PumpToolsUntil(ToolOwnerQueueContext queue,Func<bool> done)
        {
            for(int i=0;i<200;i++){queue.Drain();if(done())return;await Task.Delay(5);}
            Assert.Fail("Synthetic owner callbacks did not settle.");
        }
        [TestMethod]
        public async Task EarlyCallsWaitForRpcIdentityAndOldStartedCompletedCannotBindTheTurn()
        {
            var transport=new FakeTransport();var queue=new ToolOwnerQueueContext();var settings=new LlmSettings();object startId=null;int calls=0;
            transport.Intercept=message=>{if(Method(message)!="turn/start")return false;startId=message["id"];return true;};
            using(var client=new CodexAppServerClient(queue,new LlmVbeTools(null,null,settings),null,settings,null,transport))
            {
                client.InvokeTool=(name,args)=>{calls++;return Task.FromResult("{\"Ok\":true}");};
                var turn=client.TurnAsync("early request",null,null);Assert.IsNotNull(startId);
                transport.Emit(new{method="turn/started",@params=new{threadId="thread-1",turn=new{id="old-turn"}}});
                transport.Emit(new{id="early-old",method="item/tool/call",@params=new{threadId="thread-1",turnId="old-turn",tool="read_module",arguments=new{}}});
                transport.Emit(new{method="turn/completed",@params=new{threadId="thread-1",turn=new{id="old-turn",status="completed"}}});
                transport.Emit(new{id="early-valid",method="item/tool/call",@params=new{threadId="thread-1",turnId="new-turn",tool="read_module",arguments=new{}}});
                transport.Emit(new{id="early-legacy",method="item/tool/call",@params=new{threadId="thread-1",tool="read_module",arguments=new{}}});
                queue.Drain();Assert.AreEqual(0,calls);Assert.IsFalse(turn.IsCompleted);
                transport.Emit(new{id=startId,result=new{turn=new{id="new-turn"}}});
                await PumpToolsUntil(queue,()=>transport.Sent.Count(r=>!r.ContainsKey("method")&&r.ContainsKey("result"))==3);
                Assert.AreEqual(1,calls);Assert.IsFalse(turn.IsCompleted);
                var rejected=transport.Sent.Single(r=>!r.ContainsKey("method")&&Equals(r["id"],"early-old"));Assert.AreEqual(false,FakeTransport.Object(rejected["result"])["success"]);
                transport.Emit(new{method="turn/completed",@params=new{threadId="thread-1",turn=new{id="new-turn",status="completed"}}});queue.Drain();await turn;
            }
        }
        [TestMethod]
        public async Task FailedStartSettlesEarlyToolLatchWithoutDispatchOrDeadlock()
        {
            var transport=new FakeTransport();var queue=new ToolOwnerQueueContext();var settings=new LlmSettings();object startId=null;int calls=0;
            transport.Intercept=message=>{if(Method(message)!="turn/start")return false;startId=message["id"];return true;};
            using(var client=new CodexAppServerClient(queue,new LlmVbeTools(null,null,settings),null,settings,null,transport))
            {
                client.InvokeTool=(name,args)=>{calls++;return Task.FromResult("null");};var turn=client.TurnAsync("early request",null,null);
                transport.Emit(new{id="early-tool",method="item/tool/call",@params=new{threadId="thread-1",turnId="unknown",tool="read_module",arguments=new{}}});queue.Drain();
                transport.Emit(new{id=startId,error=new{message="synthetic start refusal"}});await Assert.ThrowsExceptionAsync<InvalidOperationException>(()=>turn);
                await PumpToolsUntil(queue,()=>transport.Sent.Any(r=>!r.ContainsKey("method")&&Equals(r["id"],"early-tool")));Assert.AreEqual(0,calls);
            }
        }
        [TestMethod]
        public async Task InterruptAndDisposeSettleEarlyIdentityWithoutAdmittingTools()
        {
            foreach(bool dispose in new[]{false,true})
            {
                var transport=new FakeTransport();var queue=new ToolOwnerQueueContext();var settings=new LlmSettings();object startId=null;int calls=0;
                transport.Intercept=message=>{if(Method(message)!="turn/start")return false;startId=message["id"];return true;};
                using(var client=new CodexAppServerClient(queue,new LlmVbeTools(null,null,settings),null,settings,null,transport))
                {
                    client.InvokeTool=(name,args)=>{calls++;return Task.FromResult("null");};var turn=client.TurnAsync("early request",null,null);
                    var latch=(TaskCompletionSource<string>)typeof(CodexAppServerClient).GetField("turnIdentityReady",System.Reflection.BindingFlags.Instance|System.Reflection.BindingFlags.NonPublic).GetValue(client);
                    transport.Emit(new{id="early-tool",method="item/tool/call",@params=new{threadId="thread-1",turnId="new-turn",tool="read_module",arguments=new{}}});queue.Drain();
                    if(dispose){client.Dispose();await Assert.ThrowsExceptionAsync<ObjectDisposedException>(()=>turn);Assert.IsTrue(latch.Task.IsCompleted);queue.Drain();}
                    else
                    {
                        await client.InterruptAsync();Assert.IsTrue(latch.Task.IsCompleted);
                        await PumpToolsUntil(queue,()=>transport.Sent.Any(r=>!r.ContainsKey("method")&&Equals(r["id"],"early-tool")));
                        transport.Emit(new{id=startId,result=new{turn=new{id="new-turn"}}});
                        await PumpToolsUntil(queue,()=>turn.IsCompleted);await Assert.ThrowsExceptionAsync<OperationCanceledException>(()=>turn);
                    }
                    Assert.AreEqual(0,calls);
                }
            }
        }
        [TestMethod]
        public async Task NaturallyCompletedTurnRejectsItsAlreadyQueuedToolCallback()
        {
            foreach(bool toolFirst in new[]{false,true})
            {
                var transport=new FakeTransport();var queue=new ToolOwnerQueueContext();var settings=new LlmSettings();int calls=0;
                using(var client=new CodexAppServerClient(queue,new LlmVbeTools(null,null,settings),null,settings,null,transport))
                {
                    client.InvokeTool=(name,args)=>{calls++;return Task.FromResult("{\"Ok\":true}");};
                    var turn=client.TurnAsync("queued request",null,null);await transport.TurnStarted.Task;
                    if(!toolFirst)transport.EmitTurnCompleted("completed",null);
                    transport.Emit(new{id="after-completion",method="item/tool/call",@params=new{threadId="thread-1",turnId="turn-1",tool="read_module",arguments=new{}}});
                    if(toolFirst)transport.EmitTurnCompleted("completed",null);
                    queue.Drain();await turn;Assert.AreEqual(0,calls);
                    Assert.AreEqual(false,FakeTransport.Object(transport.Sent.Last(r=>!r.ContainsKey("method")&&Equals(r["id"],"after-completion"))["result"])["success"]);
                }
            }
        }
        [TestMethod]
        public async Task SynchronousUiCallbackCannotAdmitToolAfterStopDisposeOrProjectRebind()
        {
            foreach(string invalidation in new[]{"stop","dispose","project"})
            {
                var transport=new FakeTransport();var settings=new LlmSettings();var tools=new LlmVbeTools(null,null,settings){BoundProject="OriginalProject"};int calls=0;
                using(var client=new CodexAppServerClient(new ImmediateContext(),tools,null,settings,null,transport))
                {
                    client.InvokeTool=(name,args)=>{calls++;return Task.FromResult("{\"Ok\":true}");};
                    client.ChatUpdate+=(kind,id,text,complete)=>
                    {
                        if(kind!="tool"||complete)return;
                        if(invalidation=="dispose")client.Dispose();
                        else if(invalidation=="stop")client.InterruptAsync().GetAwaiter().GetResult();
                        else tools.BoundProject="ChangedProject";
                    };
                    var turn=client.TurnAsync("callback request",null,null);await transport.TurnStarted.Task;
                    transport.Emit(new{id="callback-tool",method="item/tool/call",@params=new{threadId="thread-1",turnId="turn-1",tool="read_module",arguments=new{}}});
                    Assert.AreEqual(0,calls,invalidation);
                    if(invalidation=="dispose")await Assert.ThrowsExceptionAsync<ObjectDisposedException>(()=>turn);
                    else if(invalidation=="stop")await Assert.ThrowsExceptionAsync<OperationCanceledException>(()=>turn);
                    else {transport.EmitTurnCompleted("completed",null);await turn;}
                }
            }
        }
        [TestMethod]
        public async Task MissingNotificationTurnIdentityCannotPublishOrCompleteTheActiveTurn()
        {
            var transport=new FakeTransport();int updates=0;
            using(var client=Client(transport))
            {
                client.ChatUpdate+=(kind,id,text,complete)=>updates++;
                var turn=client.TurnAsync("identity request",null,null);await transport.TurnStarted.Task;
                transport.Emit(new{method="item/agentMessage/delta",@params=new{threadId="thread-1",itemId="anonymous",delta="unowned"}});
                transport.Emit(new{method="turn/completed",@params=new{threadId="thread-1",turn=new{status="completed"}}});
                Assert.AreEqual(0,updates);Assert.IsFalse(turn.IsCompleted);
                transport.EmitTurnCompleted("completed",null);await turn;
            }
        }
        [TestMethod]
        public async Task ModelCatalogueDoesNotResumeMissingHistoricalConversationOrReadInstructions()
        {
            var transport=new FakeTransport{PaginateModels=true};int instructions=0,ready=0;
            transport.Intercept=message=>
            {
                if(Method(message)!="thread/resume")return false;
                transport.Emit(new{id=message["id"],error=new{message="no rollout found for thread id synthetic-history"}});return true;
            };
            using(var client=Client(transport,"synthetic-history"))
            {
                client.DeveloperInstructionSource=()=>{instructions++;return "synthetic instructions";};client.ThreadReady+=id=>ready++;
                var models=await client.ListModelsAsync();Assert.AreEqual(2,models.Length);
                Assert.AreEqual(0,instructions);Assert.AreEqual(0,ready);
                Assert.IsFalse(transport.Methods.Any(method=>method=="thread/start"||method=="thread/resume"));
                var error=await Assert.ThrowsExceptionAsync<InvalidOperationException>(()=>client.TurnAsync("request",null,null));
                StringAssert.Contains(error.Message,"no rollout found");Assert.AreEqual("synthetic-history",client.ThreadId);
                Assert.AreEqual(1,transport.Methods.Count(method=>method=="thread/resume"));
                Assert.IsFalse(transport.Methods.Contains("thread/start"));Assert.IsFalse(transport.Methods.Contains("turn/start"));Assert.AreEqual(0,ready);
            }
        }
        [TestMethod]
        public async Task ConcurrentModelAndTurnInitializationWaitForTheSameAuthenticatedTransport()
        {
            foreach(bool fail in new[]{false,true})
            {
                var transport=new FakeTransport{CompleteTurn=true};object initializeId=null;
                transport.Intercept=message=>{if(Method(message)!="initialize")return false;initializeId=message["id"];return true;};
                using(var client=Client(transport))
                {
                    var models=client.ListModelsAsync();var turn=client.TurnAsync("request",null,null);
                    Assert.AreEqual(1,transport.Methods.Count(method=>method=="initialize"));
                    Assert.IsFalse(transport.Methods.Contains("account/read"));Assert.IsFalse(transport.Methods.Contains("model/list"));Assert.IsFalse(transport.Methods.Contains("thread/start"));
                    if(fail)transport.Emit(new{id=initializeId,error=new{message="initialization refused"}});
                    else transport.Emit(new{id=initializeId,result=new{}});
                    if(fail)
                    {
                        var modelError=await Assert.ThrowsExceptionAsync<InvalidOperationException>(()=>models);
                        var turnError=await Assert.ThrowsExceptionAsync<InvalidOperationException>(()=>turn);
                        StringAssert.Contains(modelError.Message,"initialization refused");StringAssert.Contains(turnError.Message,"initialization refused");
                        Assert.IsFalse(transport.Methods.Contains("account/read"));Assert.IsFalse(transport.Methods.Contains("thread/start"));
                    }
                    else
                    {
                        await models;Assert.AreEqual("Final answer",await turn);
                        Assert.AreEqual(1,transport.Methods.Count(method=>method=="account/read"));Assert.AreEqual(1,transport.Methods.Count(method=>method=="thread/start"));
                    }
                }
            }
        }
    }
}