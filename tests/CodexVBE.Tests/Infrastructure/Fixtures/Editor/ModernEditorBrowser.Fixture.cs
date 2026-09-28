using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using System.Threading.Tasks;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit.Editor
{
    internal sealed class OwnedWebRaw : RealProxy
    {
        internal readonly Dictionary<string, object> Values = new Dictionary<string, object>();
        internal readonly Dictionary<string, object> Handlers = new Dictionary<string, object>();
        internal Func<string, object[], object> Operation;
        internal readonly List<string> Errors = new List<string>();
        internal object Proxy => GetTransparentProxy();
        internal static Type Raw(string name) => typeof(CoreWebView2).Assembly.GetType("Microsoft.Web.WebView2.Core.Raw." + name, true);
        internal static void Complete(object handler, params object[] values) { var contract = handler.GetType().GetInterfaces().Single(type => type.FullName.StartsWith("Microsoft.Web.WebView2.Core.Raw.") && type.GetMethod("Invoke") != null); contract.GetMethod("Invoke").Invoke(handler, values); }
        internal OwnedWebRaw(Type contract) : base(contract) { }
        public override IMessage Invoke(IMessage message)
        {
            var call = (IMethodCallMessage)message; var args = (object[])call.Args.Clone(); object result = null;
            try
            {
                string name = call.MethodName;
                if (name.StartsWith("get_")) { if (!Values.TryGetValue(name.Substring(4), out result)) throw new AssertFailedException("Unexpected owned WebView getter " + name); }
                else if (name.StartsWith("set_")) Values[name.Substring(4)] = args[0];
                else if (name.StartsWith("add_")) { Handlers[name.Substring(4)] = args[0]; args[1] = Activator.CreateInstance(call.MethodBase.GetParameters()[1].ParameterType.GetElementType()); }
                else if (name.StartsWith("remove_")) Handlers.Remove(name.Substring(7));
                else { if (Operation == null) throw new AssertFailedException("Unexpected owned WebView call " + name); result = Operation(name, args); }
                return new ReturnMessage(result, args, args.Length, call.LogicalCallContext, call);
            }
            catch (Exception error) { Errors.Add(error.ToString()); return new ReturnMessage(error, call); }
        }
        internal void Fire(string name, object sender, object args)
        {
            var handler = Handlers[name]; var method = handler.GetType().GetInterfaces().Single(type => type.Namespace == "Microsoft.Web.WebView2.Core.Raw" && type.Name.EndsWith("EventHandler")).GetMethod("Invoke");
            try { method.Invoke(handler, new[] { sender, args }); } catch (TargetInvocationException error) { System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(error.InnerException).Throw(); }
        }
        internal static T Wrap<T>(object raw) => (T)typeof(T).GetConstructor(BindingFlags.NonPublic | BindingFlags.Instance, null, new[] { typeof(object) }, null).Invoke(new[] { raw });
    }
    internal sealed class ModernEditorBrowserFixture : IDisposable
    {
        internal readonly ModernEditorToolFixture Editor = new ModernEditorToolFixture();
        internal ModernEditorWindow Window => Editor.Window;
        internal readonly OwnedWebRaw Core = new OwnedWebRaw(OwnedWebRaw.Raw("ICoreWebView2_4"));
        internal readonly OwnedWebRaw Settings = new OwnedWebRaw(OwnedWebRaw.Raw("ICoreWebView2Settings"));
        internal readonly OwnedWebRaw Environment = new OwnedWebRaw(OwnedWebRaw.Raw("ICoreWebView2Environment"));
        internal readonly OwnedWebRaw Controller = new OwnedWebRaw(OwnedWebRaw.Raw("ICoreWebView2Controller"));
        internal readonly string Assets;
        internal Action EnvironmentCreated, CoreEnsured, TranslationAdded;
        internal int Scripts, Executes, NativeCloses;
        internal string Navigation, Diagnostic;
        internal OwnedWebRaw CreatedResponse;
        internal ModernEditorBrowserFixture()
        {
            Assets = System.IO.Path.Combine(Editor.Module.Root, "assets"); System.IO.Directory.CreateDirectory(Assets); System.IO.File.WriteAllText(System.IO.Path.Combine(Assets, "index.html"), "owned local editor");
            Window.BrowserAssetsDirectory = Assets; Editor.Base.Ready(false);
            Settings.Values["AreDevToolsEnabled"] = Settings.Values["AreDefaultContextMenusEnabled"] = Settings.Values["IsStatusBarEnabled"] = Settings.Values["AreHostObjectsAllowed"] = 1;
            Core.Values["Settings"] = Settings.Proxy; Core.Values["Source"] = ModernEditorWindow.Origin;
            Core.Operation = (name, values) =>
            {
                if (name == "AddHostObjectToScript") { Assert.IsNotNull(values[1]); return null; }
                if (name == "Navigate") { Navigation = (string)values[0]; return null; }
                if (name == "SetVirtualHostNameToFolderMapping") { Assert.AreEqual("editor.vbai.local", values[0]); Assert.AreEqual(Assets, values[1]); StringAssert.Contains(values[2].ToString(), "DENY_CORS"); return null; }
                if (name == "AddWebResourceRequestedFilter") { Assert.AreEqual("*", values[0]); return null; }
                if (name == "AddScriptToExecuteOnDocumentCreated") { Scripts++; TranslationAdded?.Invoke(); OwnedWebRaw.Complete(values[1], 0, "owned-script"); return null; }
                if (name == "ExecuteScript") { Executes++; OwnedWebRaw.Complete(values[1], 0, "null"); return null; }
                throw new AssertFailedException("Unexpected owned core operation " + name);
            };
            Environment.Operation = (name, values) => { if (name != "CreateWebResourceResponse") throw new AssertFailedException("Unexpected owned environment " + name); Assert.AreEqual(403, values[1]); Assert.AreEqual("Forbidden", values[2]); CreatedResponse = new OwnedWebRaw(OwnedWebRaw.Raw("ICoreWebView2WebResourceResponse")); CreatedResponse.Values["StatusCode"] = values[1]; return CreatedResponse.Proxy; };
            Controller.Values["CoreWebView2"] = Core.Proxy; Controller.Values["IsVisible"] = 0; Controller.Values["Bounds"] = Activator.CreateInstance(OwnedWebRaw.Raw("tagRECT")); Controller.Values["ZoomFactor"] = 1d; Controller.Values["ParentWindow"] = IntPtr.Zero;
            Controller.Operation = (name, values) => { if (name == "Close") { NativeCloses++; return null; } if (name == "NotifyParentWindowPositionChanged" || name == "SetBoundsAndZoomFactor") return null; throw new AssertFailedException("Unexpected owned controller operation " + name); };
            Window.CreateBrowserEnvironment = cache => { Assert.IsTrue(cache.EndsWith(System.Diagnostics.Process.GetCurrentProcess().Id.ToString())); EnvironmentCreated?.Invoke(); return Task.FromResult(OwnedWebRaw.Wrap<CoreWebView2Environment>(Environment.Proxy)); };
            Window.EnsureBrowserEnvironment = (browser, environment) => { try { Attach(browser); Assert.IsNotNull(browser.CoreWebView2); CoreEnsured?.Invoke(); return Task.CompletedTask; } catch (Exception error) { Diagnostic = error.ToString(); throw; } };
        }
        internal void Attach(WebView2 browser)
        {
            typeof(WebView2).GetField("_coreWebView2Controller", BindingFlags.NonPublic | BindingFlags.Instance).SetValue(browser, OwnedWebRaw.Wrap<CoreWebView2Controller>(Controller.Proxy));
            typeof(WebView2).GetField("<IsInitialized>k__BackingField", BindingFlags.NonPublic | BindingFlags.Instance).SetValue(browser, true);
        }
        internal Task Initialize() => (Task)Editor.Private("InitializeBrowser");
        internal void Message(string payload, string source = ModernEditorWindow.Origin)
        {
            var raw = new OwnedWebRaw(OwnedWebRaw.Raw("ICoreWebView2WebMessageReceivedEventArgs")); raw.Values["Source"] = source; raw.Values["webMessageAsJson"] = payload;
            Editor.Private("MessageReceived", null, OwnedWebRaw.Wrap<CoreWebView2WebMessageReceivedEventArgs>(raw.Proxy));
        }
        public void Dispose() { Editor.Dispose(); }
    }
}