using System;
using System.Collections.Generic;
using System.CodeDom.Compiler;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using CodexVBE;
using Microsoft.CSharp;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Infrastructure
{
    internal sealed class LlmBoundaryScope : IDisposable
    {
        private readonly Dictionary<FieldInfo, object> defaults = new Dictionary<FieldInfo, object>();
        private readonly Dictionary<string, string> environment = new Dictionary<string, string>();
        private readonly ThemeScope theme = new ThemeScope();
        private readonly LocalizationScope culture = new LocalizationScope();
        private readonly SynchronizationContext context = SynchronizationContext.Current;
        internal readonly string Root = Path.Combine(Path.GetTempPath(), "CodexLlm-" + Guid.NewGuid().ToString("N"));
        internal readonly List<string> Notices = new List<string>();
        internal readonly List<string> Logins = new List<string>();
        internal int Saves;
        internal LlmBoundaryScope()
        {
            Directory.CreateDirectory(Root);
            foreach (var field in typeof(LlmSettingsWindow).GetFields(BindingFlags.Static | BindingFlags.NonPublic).Where(x => typeof(Delegate).IsAssignableFrom(x.FieldType))) defaults[field] = field.GetValue(null);
            var handler = typeof(LlmChatClient).GetField("HttpHandlerFactory", BindingFlags.Static | BindingFlags.NonPublic); defaults[handler] = handler.GetValue(null);
            foreach (var name in LlmProvider.All.SelectMany(p => new[] { p.KeyVariable, p.ModelVariable, p.EndpointVariable }).Concat(new[] { "AZURE_OPENAI_ENTRA_TOKEN", "CODEXVBE_COPILOT_CLI", "CODEXVBE_CODEX_CLI", "CODEXVBE_TEST_COPILOT_MODE", "CODEXVBE_TEST_COPILOT_MARKER" }).Where(x => x != null).Distinct()) { environment[name] = Environment.GetEnvironmentVariable(name); Environment.SetEnvironmentVariable(name, null); }
            LlmSettingsWindow.WriteSettings = s => Saves++;
            LlmSettingsWindow.StartCopilotLogin = () => Logins.Add("copilot"); LlmSettingsWindow.StartCodexLogin = () => Logins.Add("codex");
            LlmSettingsWindow.ReadCopilotStatus = () => Task.FromResult("Copilot fixture connected");
            LlmSettingsWindow.ReadCodexStatus = () => Task.FromResult(new CodexAccountStatus(true, "ChatGPT fixture connected"));
            LlmSettingsWindow.ShowNotice = (o, t, c, b, i) => { Notices.Add(t); return DialogResult.OK; };
            LlmChatClient.HttpHandlerFactory = () => new LlmHttpFixture("{\"data\":[]}");
        }
        internal LlmSettingsWindow Window(LlmSettings settings)
        {
            if (Thread.CurrentThread.GetApartmentState() == ApartmentState.STA) SynchronizationContext.SetSynchronizationContext(new System.Windows.Threading.DispatcherSynchronizationContext());
            var window = new LlmSettingsWindow(settings) { Left = -10000, Top = -10000, ShowInTaskbar = false };
            Set(window, "githubService", new GitHubAccountService((command, token) => Task.FromResult("")));
            return window;
        }
        internal void UseCopilot(string mode = "normal")
        {
            var executable = Path.Combine(Root, "copilot-fixture.exe");
            if (!File.Exists(executable))
                using (var compiler = new CSharpCodeProvider())
                {
                    var options = new CompilerParameters(new[] { "System.dll", "System.Core.dll", "System.Web.Extensions.dll" }, executable) { GenerateExecutable = true, GenerateInMemory = false, CompilerOptions = "/target:winexe /optimize+" };
                    var result = compiler.CompileAssemblyFromSource(options, CopilotFixtureProgram.Source);
                    Assert.IsFalse(result.Errors.HasErrors, string.Join("\n", result.Errors.Cast<CompilerError>().Select(x => x.ToString())));
                }
            Environment.SetEnvironmentVariable("CODEXVBE_COPILOT_CLI", executable); Environment.SetEnvironmentVariable("CODEXVBE_TEST_COPILOT_MODE", mode); Environment.SetEnvironmentVariable("CODEXVBE_TEST_COPILOT_MARKER", Path.Combine(Root, "login.marker"));
        }
        internal static LlmProvider Provider(string name) { return LlmProvider.All.Single(x => x.Name == name); }
        internal static object Call(object target, string name, params object[] args) { return target.GetType().GetMethods(BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance).Single(x => x.Name == name && x.GetParameters().Length == args.Length).Invoke(target, args); }
        internal static T Get<T>(object target, string name) { return (T)target.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(target); }
        internal static void Set(object target, string name, object value) { target.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).SetValue(target, value); }
        internal static void Click(Control button) { typeof(Control).GetMethod("OnClick", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(button, new object[] { EventArgs.Empty }); }
        internal static void Pump(Task task)
        {
            var deadline = DateTime.UtcNow.AddSeconds(12);
            while (!task.IsCompleted && DateTime.UtcNow < deadline) { Application.DoEvents(); Thread.Sleep(1); }
            Assert.IsTrue(task.IsCompleted, "The isolated provider operation timed out."); task.GetAwaiter().GetResult();
        }
        internal static IDictionary<string, object> Object(object data) { return new JavaScriptSerializer().DeserializeObject(new JavaScriptSerializer().Serialize(data)) as IDictionary<string, object>; }
        public void Dispose()
        {
            foreach (var pair in defaults) pair.Key.SetValue(null, pair.Value);
            foreach (var pair in environment) Environment.SetEnvironmentVariable(pair.Key, pair.Value);
            theme.Dispose(); culture.Dispose(); SynchronizationContext.SetSynchronizationContext(context);
            // Process.Kill returns before Windows releases the executable image.
            for (int attempt = 0; Directory.Exists(Root); attempt++)
            {
                try { Directory.Delete(Root, true); }
                catch (IOException) { if (attempt >= 100) throw; Thread.Sleep(10); }
                catch (UnauthorizedAccessException) { if (attempt >= 100) throw; Thread.Sleep(10); }
            }
        }
    }

    internal sealed class LlmHttpFixture : HttpMessageHandler
    {
        internal sealed class Reply
        {
            internal string Body, MediaType = "application/json";
            internal bool OmitContentType;
            internal HttpStatusCode Status = HttpStatusCode.OK;
            internal Reply(string body) { Body = body; }
        }
        internal readonly Queue<Reply> Replies = new Queue<Reply>();
        internal readonly List<Uri> Uris = new List<Uri>();
        internal readonly List<string> Bodies = new List<string>();
        internal readonly List<Dictionary<string, string>> Headers = new List<Dictionary<string, string>>();
        internal Action BeforeResponse;
        internal LlmHttpFixture(params string[] bodies) { foreach (var body in bodies) Replies.Enqueue(new Reply(body)); }
        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken token)
        {
            Uris.Add(request.RequestUri); Bodies.Add(request.Content == null ? "" : await request.Content.ReadAsStringAsync()); Headers.Add(request.Headers.ToDictionary(x => x.Key, x => string.Join(",", x.Value), StringComparer.OrdinalIgnoreCase)); BeforeResponse?.Invoke();
            if (Replies.Count == 0) throw new InvalidOperationException("No fixture response remains.");
            var reply = Replies.Dequeue(); var content = new StringContent(reply.Body, System.Text.Encoding.UTF8, reply.MediaType); if (reply.OmitContentType) content.Headers.ContentType = null; return new HttpResponseMessage(reply.Status) { Content = content };
        }
    }
    internal sealed class LlmQueuedContext : SynchronizationContext
    {
        private readonly Queue<Action> work = new Queue<Action>();
        public override void Post(SendOrPostCallback callback, object state) { lock (work) work.Enqueue(() => callback(state)); }
        internal int Count { get { lock (work) return work.Count; } }
        internal void Drain() { while (true) { Action action; lock (work) { if (work.Count == 0) return; action = work.Dequeue(); } action(); } }
    }
}
