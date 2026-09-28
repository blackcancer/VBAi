using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using CodexVBE;
using CodexVBE.Tests.Infrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace CodexVBE.Tests.Unit
{
    public sealed partial class ChatWindowStateTests
    {
        private sealed class RuntimeScope : IDisposable
        {
            internal readonly Dictionary<string, object> Defaults = new Dictionary<string, object>();
            internal readonly string Root = Path.Combine(Path.GetTempPath(), "CodexChat-" + Guid.NewGuid().ToString("N"));
            internal readonly LlmSettings Settings = new LlmSettings();
            internal Func<Request, Response> Host;
            internal readonly VbeSession Session;
            internal readonly VbeSessionTests.FakeVbe Vbe = new VbeSessionTests.FakeVbe();
            internal readonly VbeSessionTests.FakeModule Module = new VbeSessionTests.FakeModule("new");
            internal RuntimeTransport Transport = new RuntimeTransport();
            internal int Saves;
            private readonly SynchronizationContext originalContext = SynchronizationContext.Current;
            private readonly ThemeScope theme = new ThemeScope();
            private readonly LocalizationScope culture = new LocalizationScope();
            internal RuntimeScope()
            {
                SynchronizationContext.SetSynchronizationContext(new System.Windows.Threading.DispatcherSynchronizationContext());
                var vbe = Vbe; var project = new VbeSessionTests.FakeProject { Name = "P", FileName = @"C:\Temp\P.xlsm", Mode = 2 }; project.VBComponents.Items.Add(new VbeSessionTests.FakeComponent { Name = "M", Type = 1, CodeModule = Module }); vbe.VBProjects.Add(project); Session = new VbeSession(vbe);
                foreach (var field in typeof(ChatWindow).GetFields(BindingFlags.Static | BindingFlags.NonPublic).Where(x => typeof(Delegate).IsAssignableFrom(x.FieldType))) Defaults[field.Name] = field.GetValue(null);
                ChatWindow.ReadSettings = () => Settings; ChatWindow.WriteSettings = s => Saves++;
                ChatWindow.HistoryPath = () => Path.Combine(Root, "chat.db");
                ChatWindow.ReadModelCatalogue = (p, s) => Task.FromResult(new[] { new LlmModelOption("model", "Model", true, "medium", new[] { new LlmEffortOption("medium", "Medium"), new LlmEffortOption("high", "High") }) });
                ChatWindow.TransportFactory = () => Transport;
                Host = r => Response.Success(r.Command == "list_projects" ? (object)new[] { new { Name = "P", FileName = @"C:\Temp\P.xlsm" } } : r.Command == "debug_state" ? new { SelectedProject = "P", SelectedProjectPath = @"C:\Temp\P.xlsm", ActiveModule = "M", Selection = new { StartLine = 1 } } : r.Command == "read_module" ? new { Code = "Sub A()\nEnd Sub", Sha256 = "sha" } : r.Command == "code_panes" ? (object)new { ActiveCodePane = new { Properties = new { Project = "P", ProjectPath = @"C:\Temp\P.xlsm", Module = "M", Selection = new { StartLine = 1, EndLine = 2, StartColumn = 1, EndColumn = 8 } } } } : new { });
                ChatWindow.ReadHost = (s, r) => Host(r);
                ChatWindow.InvokeTool = (t, n, a) => Task.FromResult(new JavaScriptSerializer().Serialize(Response.Success(new { Compiled = true })));
                ChatWindow.ShowModal = (d, o) => System.Windows.Forms.DialogResult.Cancel;
                ChatWindow.ShowSaveDialog = (d, o) => System.Windows.Forms.DialogResult.Cancel;
                ChatWindow.ShowNotice = (o, t, c, b, i) => System.Windows.Forms.DialogResult.OK;
                ChatWindow.WriteClipboard = s => { };
            }
            public void Dispose()
            {
                foreach (var entry in Defaults) typeof(ChatWindow).GetField(entry.Key, BindingFlags.Static | BindingFlags.NonPublic).SetValue(null, entry.Value);
                theme.Dispose(); culture.Dispose(); SynchronizationContext.SetSynchronizationContext(originalContext);
                if (Directory.Exists(Root)) Directory.Delete(Root, true);
            }
        }
        public sealed class UnavailableReferenceHost { public object VBProjects { get { throw new IOException("reference catalogue unavailable"); } } }
        private sealed class RuntimeHttpHandler : System.Net.Http.HttpMessageHandler
        {
            internal Action BeforeResponse;
            internal string Body;
            internal bool Streaming;
            protected override Task<System.Net.Http.HttpResponseMessage> SendAsync(System.Net.Http.HttpRequestMessage request, CancellationToken cancellationToken)
            {
                BeforeResponse?.Invoke();
                var content = new System.Net.Http.StringContent(Body ?? "{\"choices\":[{\"message\":{\"role\":\"assistant\",\"content\":\"answer\"}}]}");
                if (Streaming) content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("text/event-stream");
                return Task.FromResult(new System.Net.Http.HttpResponseMessage(System.Net.HttpStatusCode.OK) { Content = content });
            }
        }
        private sealed class RuntimeTransport : ICodexAppServerTransport
        {
            private readonly JavaScriptSerializer json = new JavaScriptSerializer();
            public event Action<string> LineReceived;
            public event Action<Exception> Exited;
            public bool IsRunning { get; private set; }
            internal bool FailModels, FailTurn, Complete = true;
            internal Action BeforeComplete;
            internal object[] Models = { new { model = "model", displayName = "Model", isDefault = true, defaultReasoningEffort = "medium", supportedReasoningEfforts = new[] { new { reasoningEffort = "medium", description = "Medium" } } } };
            public void Start() { IsRunning = true; }
            public void Dispose() { IsRunning = false; }
            internal void Emit(object value) { LineReceived?.Invoke(json.Serialize(value)); }
            internal void Exit() { Exited?.Invoke(new IOException("closed")); }
            public void Send(string line)
            {
                var msg = (IDictionary<string, object>)json.DeserializeObject(line);
                if (!msg.ContainsKey("id") || !msg.ContainsKey("method")) return;
                var id = msg["id"]; var method = Convert.ToString(msg["method"]);
                if (method == "model/list" && FailModels) { Emit(new { id, error = new { message = "models failed" } }); return; }
                object result = method == "account/read" ? (object)new { account = new { type = "chatgpt" } } : method == "model/list" ? new { data = Models, nextCursor = (string)null } : method == "thread/start" || method == "thread/resume" ? (object)new { thread = new { id = "thread" } } : method == "turn/start" ? (object)new { turn = new { id = "turn" } } : new { };
                if (method == "turn/start") Emit(new { method = "turn/started", @params = new { threadId = "thread", turn = new { id = "turn" } } });
                Emit(new { id, result });
                if (method == "turn/start" && Complete) { BeforeComplete?.Invoke(); Emit(new { method = "item/agentMessage/delta", @params = new { threadId = "thread", itemId = "answer", delta = "answer" } }); Emit(new { method = "item/completed", @params = new { threadId = "thread", item = new { type = "agentMessage", phase = "final", id = "answer", text = "answer" } } }); Emit(new { method = "turn/completed", @params = new { threadId = "thread", turn = new { status = FailTurn ? "failed" : "completed", error = FailTurn ? new { message = "turn failed" } : null } } }); }
                if (method == "turn/interrupt") Emit(new { method = "turn/completed", @params = new { threadId = "thread", turn = new { status = "interrupted", error = (object)null } } });
            }
        }
        private static T Visual<T>(DependencyObject root) where T : DependencyObject { if (root is T value) return value; for (int i = 0; i < System.Windows.Media.VisualTreeHelper.GetChildrenCount(root); i++) { var result = Visual<T>(System.Windows.Media.VisualTreeHelper.GetChild(root, i)); if (result != null) return result; } return null; }
        private static void RaiseScroll(ChatWindow window, UIElement source, double vertical, double extent) { var args = (ScrollChangedEventArgs)Activator.CreateInstance(typeof(ScrollChangedEventArgs), BindingFlags.Instance | BindingFlags.NonPublic, null, new object[] { new Vector(0, vertical), new Vector(0, vertical), new Size(100, 100), new Vector(0, extent), new Size(100, 30), new Vector(0, 0) }, null); args.RoutedEvent = ScrollViewer.ScrollChangedEvent; source.RaiseEvent(args); }
        private static void TimerTick(System.Windows.Threading.DispatcherTimer timer) { var method = typeof(System.Windows.Threading.DispatcherTimer).GetMethod("FireTick", BindingFlags.Instance | BindingFlags.NonPublic); method.Invoke(timer, new object[method.GetParameters().Length]); }
        private static void Click(System.Windows.Forms.Control control) { typeof(System.Windows.Forms.Control).GetMethod("OnClick", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(control, new object[] { EventArgs.Empty }); }
        private static IEnumerable<FrameworkElement> Descendants(FrameworkElement element)
        {
            yield return element;
            foreach (var child in LogicalTreeHelper.GetChildren(element).OfType<FrameworkElement>()) foreach (var nested in Descendants(child)) yield return nested;
        }
        private static void WpfClick(System.Windows.Controls.Button button) { button.RaiseEvent(new RoutedEventArgs(System.Windows.Controls.Button.ClickEvent)); }
        private static KeyEventArgs RunKey(ChatWindow window, Key key)
        {
            var prompt = Get<System.Windows.Controls.TextBox>(window, "prompt");
            using (var source = new System.Windows.Interop.HwndSource(new System.Windows.Interop.HwndSourceParameters("test") { Width = 1, Height = 1, PositionX = -10000, PositionY = -10000 }))
            {
                // Source lifetime only supplies keyboard event routing; no input is synthesized.
                var args = new KeyEventArgs(Keyboard.PrimaryDevice, source, 0, key) { RoutedEvent = Keyboard.PreviewKeyDownEvent }; Call(window, "PromptKeyDown", prompt, args); return args;
            }
        }
    }
}
