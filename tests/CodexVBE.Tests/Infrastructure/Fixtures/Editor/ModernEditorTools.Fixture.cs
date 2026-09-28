using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Web.Script.Serialization;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
namespace CodexVBE.Tests.Unit.Editor
{
    internal sealed class ModernEditorToolFixture : IDisposable
    {
        internal readonly ModernEditorDebugFixture Base = new ModernEditorDebugFixture();
        internal ModernEditorWindow Window => Base.Window;
        internal EditorFixture Module => Base.Storage;
        internal readonly EditorDocument Document;
        internal Dictionary<string, int> Versions => Base.Get<Dictionary<string, int>>("versions");
        internal readonly JavaScriptSerializer Json = new JavaScriptSerializer();
        internal Func<string, object[], string> Override;
        internal int Captures, Applies;
        internal ModernEditorToolFixture()
        {
            Document = ModernEditorDebugFixture.Wait(Window.OpenModule(Module)); Base.Ready(true);
            Base.Rendering = (method, values) =>
            {
                if (method == "snapshots") Captures++; if (method == "apply") Applies++;
                string overridden = Override?.Invoke(method, values); if (overridden != null) return overridden;
                if (method == "read") return Json.Serialize(new { text = Document.Text, version = Versions[Document.Id], selection = new { startLineNumber = 1 } });
                if (method == "apply") return ((int)values[1] + 1).ToString();
                return "null";
            };
        }
        internal object Snapshot(string text, int version) => new { id = Document.Id, text, version };
        internal Dictionary<string, object> Result(object value) => Json.Deserialize<Dictionary<string, object>>(Json.Serialize(value));
        internal object Private(string name, params object[] values)
        {
            try { return typeof(ModernEditorWindow).GetMethod(name, BindingFlags.Instance | BindingFlags.NonPublic).Invoke(Window, values); }
            catch (TargetInvocationException error) { System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(error.InnerException).Throw(); throw; }
        }
        public void Dispose() { Base.Dispose(); }
    }
}