namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading;
    using System.Web.Script.Serialization;
    using CodexVBE;
    using CodexVBE.Tests.Infrastructure;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class LlmVbeToolsBoundaryTests
    {
        private static readonly JavaScriptSerializer Json = new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 };
        private static IDictionary<string, object> Dict(object value) { return (IDictionary<string, object>)value; }
        private sealed class ToolFixture
        {
            internal readonly LlmSettings Settings = new LlmSettings { VbeEditApproval = "Automatic" };
            internal readonly LlmVbeTools Tools;
            internal ToolFixture()
            {
                Tools = new LlmVbeTools(new VbeSession(new VbeSessionTests.FakeVbe()), null, Settings);
                VbeToolBoundaryFixture.Configure(Tools.Native);
                Tools.Execute = VbeToolBoundaryFixture.Execute;
                Tools.PersistSignature = p => new CodexVBE.Tests.Infrastructure.VbeToolPersistence { Saved = true };
            }
        }
        // Each tool retains its real orchestrator; only its host/native boundaries are replaced.
        private static ToolProxy Create() { return new ToolProxy(); }
        private sealed class ToolProxy
        {
            private readonly ToolFixture fixture = new ToolFixture();
            internal LlmSettings Settings { get { return fixture.Settings; } }
            internal VbeToolNativeBoundary Native { get { return fixture.Tools.Native; } }
            internal Func<Request, Response> Execute { set { fixture.Tools.Execute = value; } }
            internal Func<string, object> PersistSignature { set { fixture.Tools.PersistSignature = value; } }
            internal string CurrentProviderName { set { fixture.Tools.CurrentProviderName = value; } }
            internal string BoundProject { set { fixture.Tools.BoundProject = value; } }
            internal ChatMode Mode { set { fixture.Tools.Mode = value; } }
            internal Func<System.Windows.Forms.IWin32Window,string,string,System.Windows.Forms.DialogResult> ConfirmFile { set { fixture.Tools.ConfirmFile = value; } }
            internal Func<VbeApprovalDialog,System.Windows.Forms.IWin32Window,System.Windows.Forms.DialogResult> ShowApproval { set { fixture.Tools.ShowApproval = value; } }
            internal void NoteUserRequest(string text) { fixture.Tools.NoteUserRequest(text); }
            internal string Invoke(string name,string arguments) { return fixture.Tools.Invoke(name,arguments); }
            internal System.Threading.Tasks.Task<string> InvokeAsync(string name,string arguments) { return fixture.Tools.InvokeAsync(name,arguments); }
        }
        private static Dictionary<string, object> Arguments(string name)
        {
            var definition = LlmVbeTools.Definitions.Select(d => Dict(Dict(Json.DeserializeObject(Json.Serialize(d)))["function"]))
                .Single(d => (string)d["name"] == name);
            var fields = Dict(Dict(definition["parameters"])["properties"]);
            return fields.ToDictionary(f => f.Key, f => f.Key == "Value" ? (object)"value" :
                f.Key == "Items" ? new string[0] : f.Key == "PathSegments" ? new[] { "item" } :
                (string)Dict(f.Value)["type"] == "integer" ? (object)2 :
                (string)Dict(f.Value)["type"] == "number" ? (object)1.5 :
                (string)Dict(f.Value)["type"] == "boolean" ? (object)true :
                f.Key == "Path" ? @"C:\Temp\fixture.bas" : f.Key == "Project" ? "P" : "value");
        }
        private static Dictionary<string, object> AsyncArguments(string name)
        {
            if (name == "debug_item") return new Dictionary<string, object> { ["Pane"]="locals",["Action"]="expand",["PathSegments"]=new[] {"item"} };
            if (name == "immediate_execute") return new Dictionary<string, object> { ["Project"]="P",["ExpectedMode"]=2,["Text"]="Debug.Print 1" };
            if (name == "debug_dialog") return new Dictionary<string, object>();
            if (name == "respond_debug_dialog") return new Dictionary<string, object> { ["Diagnostic"]="fixture",["Button"]="ok" };
            return Arguments(name);
        }
        private static void Success(string json,string context)
        { var response=Json.Deserialize<Response>(json); Assert.IsTrue(response.Ok,context+": "+response.Error); }
        private static void Failed(string json,string context)
        { var response=Json.Deserialize<Response>(json); Assert.IsFalse(response.Ok,context); Assert.IsFalse(string.IsNullOrWhiteSpace(response.Error),context); }
        private static IDictionary<string,object> Data(string json)
        { Success(json,"response"); return Dict(Json.Deserialize<Response>(json).Data); }
        private sealed class ImmediateContext : SynchronizationContext
        { public override void Post(SendOrPostCallback callback, object state) { callback(state); } }
        private sealed class DeferredContext : SynchronizationContext
        { private int posts; public override void Post(SendOrPostCallback callback, object state) { if (Interlocked.Increment(ref posts)>1) ThreadPool.QueueUserWorkItem(_=>callback(state)); } }
    }
}
