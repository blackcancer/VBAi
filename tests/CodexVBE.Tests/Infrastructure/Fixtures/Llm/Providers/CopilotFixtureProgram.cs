namespace CodexVBE.Tests.Infrastructure
{
    // This executable implements only fixture RPC. It never opens a browser, host or network connection.
    internal static class CopilotFixtureProgram
    {
        internal const string Source = @"
using System;using System.IO;using System.Text;using System.Linq;using System.Collections.Generic;using System.Web.Script.Serialization;
internal static class Fixture {
static readonly JavaScriptSerializer Json=new JavaScriptSerializer();static string Session;static readonly Stream Input=Console.OpenStandardInput(),Output=Console.OpenStandardOutput();
static IDictionary<string,object> Obj(object o){return o as IDictionary<string,object>??new Dictionary<string,object>();}static string Text(IDictionary<string,object> o,string k){object v;return o.TryGetValue(k,out v)?Convert.ToString(v):null;}
static void Raw(string s){byte[] b=Encoding.UTF8.GetBytes(s);Output.Write(b,0,b.Length);Output.Flush();}
static void Frame(object o){string s=Json.Serialize(o);Raw(""Content-Length: ""+Encoding.UTF8.GetByteCount(s)+""\r\n\r\n""+s);}
static IDictionary<string,object> Read(){var h=new StringBuilder();while(!h.ToString().EndsWith(""\r\n\r\n"")){int b=Input.ReadByte();if(b<0)return null;h.Append((char)b);}int n=int.Parse(h.ToString().Substring(15).Trim());byte[] bytes=new byte[n];int i=0;while(i<n){int c=Input.Read(bytes,i,n-i);if(c==0)return null;i+=c;}return Obj(Json.DeserializeObject(Encoding.UTF8.GetString(bytes)));}
static void Event(string type,object data){Frame(new{method=""session.event"",@params=new{sessionId=Session,@event=new{type,data}}});}
static void Finish(string mode){Event(""assistant.message_delta"",new{deltaContent=""""});Event(""assistant.message_delta"",new{deltaContent=""delta été""});Event(""assistant.message"",new{content=""answer été""});Event(""session.idle"",new{});Event(""session.idle"",new{});}
static void Main(string[] args){
 if(args.Contains(""login"")){File.WriteAllText(Environment.GetEnvironmentVariable(""CODEXVBE_TEST_COPILOT_MARKER""),""fixture login only"");return;}
 string mode=Environment.GetEnvironmentVariable(""CODEXVBE_TEST_COPILOT_MODE"")??""normal"";
 if(mode==""header-eof""||mode==""header-long""||mode==""size-zero""||mode==""size-large""||mode==""size-invalid""||mode==""size-negative""||mode==""body-eof""||mode==""bad-json"")Read();
 if(mode==""header-eof"")return;if(mode==""header-long""){Raw(new string('x',4097));return;}if(mode==""size-zero""){Raw(""Other: x\r\nContent-Length: 0\r\n\r\n"");return;}if(mode==""size-large""){Raw(""Content-Length: 10485761\r\n\r\n"");return;}if(mode==""size-invalid""){Raw(""Content-Length: abc\r\n\r\n"");return;}if(mode==""size-negative""){Raw(""Content-Length: -1\r\n\r\n"");return;}if(mode==""body-eof""){Raw(""Content-Length: 9\r\n\r\n{}"");return;}if(mode==""bad-json""){Raw(""Content-Length: 1\r\n\r\nx"");return;}
 while(true){var message=Read();if(message==null)return;string method=Text(message,""method"");object id=message.ContainsKey(""id"")?message[""id""]:null;var p=message.ContainsKey(""params"")?Obj(message[""params""]):new Dictionary<string,object>();Action<object> reply=r=>Frame(new{id,result=r});
 if(method==""ping""){if(mode==""request-wait"")continue;if(mode==""rpc-error""){Frame(new{id,error=new{message=""secret must not leak""}});continue;}Frame(new{id=""nonnumeric"",result=new{}});Frame(new{id=999,result=new{}});Frame(new{method=""notification.unknown""});if(mode==""early-tool"")Frame(new{id=""early-tool"",method=""tool.call"",@params=new{sessionId=(string)null,toolName=""read_module""}});reply(new{protocolVersion=mode==""version-invalid""?""x"":mode==""version-low""?""1"":mode==""version-high""?""4"":mode==""legacy""?""2"":""3""});}
 else if(method==""models.list""){if(mode==""models-null"")reply(null);else reply(new{models=new object[]{null,new{},new{id="" ""},new{id=""model"",name=""Model""},new{id=""fallback""}}});if(mode==""exit-models"")return;}
 else if(method==""session.create""){Session=Text(p,""sessionId"");reply(new{sessionId=mode==""session-mismatch""?""foreign"":Session});}
 else if(method==""session.send""){reply(new{});Frame(new{method=""session.event"",@params=new{sessionId=""foreign"",@event=new{type=""session.idle"",data=new{}}}});Frame(new{method=""session.event"",@params=new{sessionId=Session,@event=new{type=""unknown.event"",data=new{}}}});if(mode==""completion-wait"")continue;if(mode==""session-error""){Event(""session.error"",new{});continue;}if(mode==""empty-answer""){Event(""session.idle"",new{});continue;}if(mode==""queued-tool""){Event(""external_tool.requested"",new{requestId=""tool"",toolName=""read_module""});Finish(mode);continue;}if(mode==""queued-wait""){Event(""assistant.message_delta"",new{deltaContent=""queued""});Event(""external_tool.requested"",new{requestId=""tool"",toolName=""read_module""});continue;}
 Frame(new{id=""unsupported"",method=""future.client.method"",@params=new{sessionId=Session}});Frame(new{id=""foreign-permission"",method=""permission.request"",@params=new{sessionId=""foreign"",permissionRequest=new{kind=""custom-tool"",toolName=""read_module""}}});
 if(mode==""legacy""){Frame(new{id=""permission"",method=""permission.request"",@params=new{sessionId=Session,permissionRequest=new{kind=""custom-tool"",toolName=""read_module""}}});Frame(new{id=""tool"",method=""tool.call"",@params=new{sessionId=Session,toolName=""read_module"",arguments=new{},toolCallId=""call""}});}
 else {Event(""permission.requested"",new{requestId=""permission"",permissionRequest=new{kind=""shell""}});Event(""external_tool.requested"",new{requestId=""tool"",toolName=mode==""unknown-tool""?""unknown"":""read_module""});Event(""external_tool.requested"",new{requestId=""tool"",toolName=""read_module""});}}
 else if(method==""session.permissions.handlePendingPermissionRequest""){if(mode==""permission-error"")Frame(new{id,error=new{}});else reply(new{});}
 else if(method==""session.tools.handlePendingToolCall""){if(mode==""tool-reply-error"")Frame(new{id,error=new{}});else {reply(new{});Finish(mode);}}
 else if(method==null&&Convert.ToString(id)==""tool"")Finish(mode);
 }
}
}";
    }
}
