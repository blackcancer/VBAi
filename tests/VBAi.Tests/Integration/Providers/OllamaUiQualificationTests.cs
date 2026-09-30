namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using System.Net;
    using System.Net.Http;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;
    using System.Web.Script.Serialization;
    using System.Windows;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Forms = System.Windows.Forms;

    // Shares the isolated synthetic VBE/settings/history fixture; HTTP is deliberately not mocked.
    public sealed partial class ChatWindowStateTests
    {
        [STATestMethod, TestCategory("OllamaUi"), DoNotParallelize]
        public void LocalOllamaShownChatStreamsStopsAndCompletesNextSend()
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_OLLAMA_UI_TESTS") != "1")
                Assert.Inconclusive("Set VBAi_RUN_OLLAMA_UI_TESTS=1 for detached real chat UI plus local Ollama; VBE remains simulated.");
            string model = Environment.GetEnvironmentVariable("VBAi_TEST_OLLAMA_MODEL") ?? "qwen2.5:3b";
            var provider = LlmProvider.All.Single(item => item.IsOllama);
            using (var runtime = new RuntimeScope())
            using (var wire = new OllamaWireScope())
            {
                runtime.Settings.ProviderName = provider.Name;
                runtime.Settings.OllamaEndpoint = "http://127.0.0.1:11434/v1/chat/completions";
                runtime.Settings.OllamaModel = model;
                runtime.Settings.VbeEditApproval = "ReadOnly";
                // This ephemeral value shadows any inherited API-key environment variable.
                // It is sent only to the fixed loopback endpoint and never persisted as user settings.
                runtime.Settings.SetKey(provider, "vbai-synthetic-local-ui-test");
                ChatWindow.ReadModelCatalogue = LlmChatClient.ListModelsAsync;
                int refusedTools = 0;
                var toolCalls = new List<object>();
                ChatWindow.InvokeTool = (tools, name, arguments) => {
                    refusedTools++;
                    toolCalls.Add(new { Name = name, Arguments = arguments, Result = "Refused by synthetic fixture", Utc = DateTime.UtcNow.ToString("o") });
                    return Task.FromResult(new JavaScriptSerializer().Serialize(Response.Failure(
                        "Synthetic UI test: tool execution is disabled. Answer the user's generic text request without tools.")));
                };
                using (var window = LoadedWindow(runtime.Session))
                {
                    Assert.IsNull(window.HttpHandlerOverride, "Live qualification must retain the production HTTP transport.");
                    window.Show();
                    var prompt = Get<System.Windows.Controls.TextBox>(window, "prompt");
                    var send = Get<Forms.Button>(window, "send");
                    var providers = Get<Forms.ComboBox>(window, "providerPicker");
                    var models = Get<Forms.ComboBox>(window, "modelPicker");
                    var modes = Get<Forms.ComboBox>(window, "modePicker");
                    WaitOllamaUi(window, () => !Get<bool>(window, "loadingScope") &&
                        (models.SelectedItem as LlmModelOption)?.Id == model, 45, "select installed local model");
                    Get<Forms.CheckBox>(window, "verifyAfterEdit").Checked = false;
                    modes.SelectedItem = ChatMode.Discussion;
                    int streamedCharacters = 0;
                    var observations = new List<object>();
                    var elapsed = Stopwatch.StartNew();
                    long lastObservation = -1000;
                    string lastStage = null;
                    Action<string> observe = stage => {
                        if (stage == lastStage && elapsed.ElapsedMilliseconds - lastObservation < 500) return;
                        lastStage = stage; lastObservation = elapsed.ElapsedMilliseconds;
                        if (observations.Count >= 512) return;
                        try { observations.Add(OllamaUiSnapshot(window, stage, elapsed.ElapsedMilliseconds)); }
                        catch (Exception diagnostic) { observations.Add(new { Stage = stage, DiagnosticError = diagnostic.ToString() }); }
                    };
                    Exception primaryFailure = null;
                    try
                    {
                        prompt.Text = "This is a synthetic interface test unrelated to VBA. Do not use tools. Write a long numbered list of 1000 everyday objects, beginning immediately with item 1 and continuing without introductory remarks.";
                        send.PerformClick();
                        observe("first Send returned");
                        Assert.IsTrue(Get<bool>(window, "busy"), "The normal Send button must start the turn.");
                        Assert.AreEqual("", prompt.Text);
                        Assert.IsFalse(providers.Enabled);
                        Assert.IsFalse(models.Enabled);
                        Assert.IsFalse(modes.Enabled);
                        WaitOllamaUi(window, () => {
                            streamedCharacters = Get<Dictionary<string, Forms.RichTextBox>>(window, "liveTexts").Values
                                .Where(text => !text.IsDisposed && text.Visible).Sum(text => text.TextLength);
                            return Get<bool>(window, "busy") && streamedCharacters > 0;
                        }, 120, "render real streamed text while the turn is still active", observe);
                        Assert.IsTrue(send.Enabled, "Stop must remain available during streaming.");
                        send.PerformClick();
                        Assert.IsTrue(Get<bool>(window, "stopRequested"));
                        WaitOllamaUi(window, () => !Get<bool>(window, "busy"), 30, "finish cancellation", observe);
                        Assert.IsTrue(providers.Enabled && models.Enabled && modes.Enabled && send.Enabled);
                        Assert.IsNull(Get<LlmChatClient>(window, "activeHttpClient"));
                        Assert.IsTrue(Get<List<ChatEntry>>(window, "transcriptEntries").Any(entry =>
                            (entry.Text ?? "").Contains(UiText.Get("Response interrupted. Changes already applied can still be undone in the chat."))), "Cancellation must remain visible in the transcript.");

                        prompt.Text = "The previous response was intentionally stopped. This is a synthetic UI test unrelated to VBA. Do not use tools. Reply with exactly UI_READY_42 and nothing else.";
                        send.PerformClick();
                        Assert.IsTrue(Get<bool>(window, "busy"));
                        WaitOllamaUi(window, () => !Get<bool>(window, "busy"), 120, "complete the next real model response", observe);
                        Assert.IsFalse(Get<bool>(window, "stopRequested"));
                        Assert.IsTrue(providers.Enabled && models.Enabled && modes.Enabled && send.Enabled);
                        Assert.IsNull(Get<LlmChatClient>(window, "activeHttpClient"));
                        var answer = Get<List<ChatEntry>>(window, "transcriptEntries").LastOrDefault(entry =>
                            entry.Speaker == "Assistant" && (entry.Text ?? "").Contains("UI_READY_42"));
                        Assert.IsNotNull(answer, "The next model response must complete in the transcript.");
                        WaitOllamaUi(window, () => Get<Dictionary<ChatEntry, FrameworkElement>>(window, "entryViews")
                            .TryGetValue(answer, out var view) && NativeDescendants(view).OfType<Forms.RichTextBox>()
                                .Any(text => text.Visible && text.Text.Contains("UI_READY_42")),
                            10, "materialize the completed response in visible native text controls", observe);
                        Console.WriteLine("OllamaUi: model={0}; streamed visible characters before stop={1}; cancellation and next send completed; refused synthetic tool calls={2}; assembly={3}; MVID={4}. Detached real UI and production HTTP; simulated VBE, no native Office/SOLIDWORKS qualification.",
                            model, streamedCharacters, refusedTools, typeof(LlmChatClient).Assembly.Location, typeof(LlmChatClient).Module.ModuleVersionId);
                    }
                    catch (Exception error)
                    {
                        primaryFailure = error;
                        observe("failure before cleanup");
                        WriteOllamaUiFailure(window, model, error, observations, toolCalls);
                        throw;
                    }
                    finally
                    {
                        try
                        {
                            if (Get<bool>(window, "busy"))
                            {
                                prompt.Clear();
                                send.PerformClick();
                                WaitOllamaUi(window, () => !Get<bool>(window, "busy"), 30, "stop unfinished local response during cleanup", observe);
                            }
                        }
                        catch (Exception cleanup) when (primaryFailure != null)
                        { Console.WriteLine("Ollama UI cleanup also failed; primary scenario failure retained: " + cleanup); }
                    }
                }
            }
        }

        private static object OllamaUiSnapshot(ChatWindow window, string stage, long elapsedMilliseconds)
        {
            var texts = Get<Dictionary<string, Forms.RichTextBox>>(window, "liveTexts");
            var entries = Get<List<ChatEntry>>(window, "transcriptEntries");
            return new {
                Stage = stage, ElapsedMilliseconds = elapsedMilliseconds,
                Busy = Get<bool>(window, "busy"), StopRequested = Get<bool>(window, "stopRequested"),
                Status = Get<Forms.Label>(window, "status").Text, WindowVisible = window.Visible,
                ActiveHttpClient = Get<LlmChatClient>(window, "activeHttpClient") != null,
                StreamDiagnostics = (Get<LlmChatClient>(window, "activeHttpClient")?.LastStreamDiagnostics ?? Get<StreamDiagnostics>(window, "lastHttpStreamDiagnostics"))?.Snapshot(),
                EntryCount = entries.Count, TranscriptCharacters = entries.Sum(entry => (entry.Text ?? "").Length),
                EntryViews = Get<Dictionary<ChatEntry, FrameworkElement>>(window, "entryViews").Count,
                PendingStreamCharacters = Get<Dictionary<ChatEntry, StringBuilder>>(window, "pendingStreamText").Values.Sum(text => text.Length),
                LiveTexts = texts.Select(item => new { StreamId = item.Key, Disposed = item.Value.IsDisposed,
                    Visible = !item.Value.IsDisposed && item.Value.Visible,
                    Characters = item.Value.IsDisposed ? 0 : item.Value.TextLength }).ToArray()
            };
        }

        private static void WriteOllamaUiFailure(ChatWindow window, string model, Exception error,
            List<object> observations, List<object> toolCalls)
        {
            try
            {
                // RuntimeScope has a new synthetic VBE project and isolated history/settings. Do not
                // collect credentials, settings objects, user files, or the desktop outside this window.
                var transcript = Get<List<ChatEntry>>(window, "transcriptEntries").Select(entry => new {
                    entry.Speaker, entry.StreamId, entry.TurnId, entry.Text,
                    ActivityStatus = entry.Activity?.Status
                }).ToArray();
                var views = Get<Dictionary<ChatEntry, FrameworkElement>>(window, "entryViews").Select(item => new {
                    item.Key.Speaker, item.Key.StreamId, ViewType = item.Value.GetType().FullName,
                    item.Value.IsVisible, item.Value.ActualWidth, item.Value.ActualHeight,
                    NativeTexts = NativeDescendants(item.Value).OfType<Forms.RichTextBox>().Select(text => new {
                        Disposed = text.IsDisposed, Visible = !text.IsDisposed && text.Visible,
                        Text = text.IsDisposed ? null : text.Text, Width = text.Width, Height = text.Height
                    }).ToArray()
                }).ToArray();
                var result = new { State = "FAIL", Model = model, Failure = error.ToString(),
                    Scope = "Detached synthetic ChatWindow, real loopback Ollama HTTP; no native host or tool execution",
                    Mvid = typeof(LlmChatClient).Module.ModuleVersionId.ToString("D"),
                    Observations = observations, Transcript = transcript, EntryViews = views, RefusedToolCalls = toolCalls };
                string root = Environment.GetEnvironmentVariable("VBAi_OLLAMA_UI_RESULTS") ??
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ollama-ui-diagnostics");
                Directory.CreateDirectory(root);
                string path = Path.Combine(root, "failure-" + Guid.NewGuid().ToString("N") + ".json");
                File.WriteAllText(path, new JavaScriptSerializer { MaxJsonLength = 4 * 1024 * 1024 }.Serialize(result));
                Console.WriteLine("Ollama UI synthetic failure diagnostics: " + path);
                Console.WriteLine("Ollama UI terminal snapshot: " + new JavaScriptSerializer().Serialize(OllamaUiSnapshot(window, "failure", 0)));
            }
            catch (Exception diagnostic) { Console.WriteLine("Ollama UI diagnostics failed; original assertion retained: " + diagnostic); }
        }

        private static void WaitOllamaUi(ChatWindow window, Func<bool> completed, int seconds, string stage, Action<string> observe = null)
        {
            var watch = Stopwatch.StartNew();
            while (!completed() && watch.Elapsed.TotalSeconds < seconds)
            {
                observe?.Invoke(stage);
                Forms.Application.DoEvents();
                Thread.Sleep(10);
            }
            observe?.Invoke(stage + " ended");
            Assert.IsTrue(completed(), "Ollama UI timeout at " + stage + "; status=" + Get<Forms.Label>(window, "status").Text);
        }

        // Optional qualification instrumentation. It observes bytes as the production parser reads
        // them, without eagerly reading the response, replacing it, or logging authentication headers.
        private sealed class OllamaWireScope : IDisposable
        {
            private readonly Func<HttpMessageHandler> previous = LlmChatClient.HttpHandlerFactory;
            private readonly bool enabled = Environment.GetEnvironmentVariable("VBAi_OLLAMA_UI_CAPTURE_WIRE") == "1";
            internal OllamaWireScope()
            {
                if (!enabled) return;
                string root = Environment.GetEnvironmentVariable("VBAi_OLLAMA_UI_RESULTS") ??
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ollama-ui-diagnostics");
                Directory.CreateDirectory(root);
                LlmChatClient.HttpHandlerFactory = () => {
                    var inner = previous();
                    Assert.IsInstanceOfType<HttpClientHandler>(inner, "Wire capture must retain the real HTTP transport.");
                    Assert.IsFalse(((HttpClientHandler)inner).AllowAutoRedirect);
                    return new OllamaWireHandler(inner, root);
                };
            }
            public void Dispose() { if (enabled) LlmChatClient.HttpHandlerFactory = previous; }
        }

        private sealed class OllamaWireHandler : DelegatingHandler
        {
            private readonly string root;
            internal OllamaWireHandler(HttpMessageHandler inner, string root) : base(inner) { this.root = root; }
            protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken token)
            {
                if (request.RequestUri.Scheme != "http" || !request.RequestUri.IsLoopback || request.RequestUri.Port != 11434)
                    throw new InvalidOperationException("Synthetic wire capture requires the fixed local Ollama endpoint.");
                string prefix = Path.Combine(root, "wire-" + Guid.NewGuid().ToString("N"));
                if (request.Content != null)
                {
                    // The client creates StringContent; observing it does not consume the network stream.
                    byte[] bytes = await request.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
                    SaveWire(prefix + "-request.json", bytes.Take(OllamaTeeStream.Limit).ToArray());
                }
                var response = await base.SendAsync(request, token).ConfigureAwait(false);
                SaveWire(prefix + "-metadata.json", Encoding.UTF8.GetBytes(new JavaScriptSerializer().Serialize(new {
                    Utc = DateTime.UtcNow.ToString("o"), Method = request.Method.Method,
                    Path = request.RequestUri.AbsolutePath, Status = (int)response.StatusCode,
                    ContentType = response.Content?.Headers.ContentType?.ToString(),
                    Mvid = typeof(LlmChatClient).Module.ModuleVersionId.ToString("D"),
                    Scope = "Synthetic fixture bodies only; no headers, credentials or user history recorded",
                    CaptureByteLimit = OllamaTeeStream.Limit
                })));
                if (response.Content != null) response.Content = new OllamaWireContent(response.Content, prefix);
                return response;
            }
        }

        private static void SaveWire(string path, byte[] bytes)
        {
            try { File.WriteAllBytes(path, bytes); }
            catch (Exception error) { Console.WriteLine("Passive Ollama capture failed: " + error.GetType().Name); }
        }

        private sealed class OllamaWireContent : HttpContent
        {
            private readonly HttpContent original;
            private readonly string prefix;
            private Stream stream;
            internal OllamaWireContent(HttpContent original, string prefix)
            {
                this.original = original; this.prefix = prefix;
                foreach (var header in original.Headers) Headers.TryAddWithoutValidation(header.Key, header.Value);
            }
            protected override async Task<Stream> CreateContentReadStreamAsync()
            {
                return stream ?? (stream = new OllamaTeeStream(await original.ReadAsStreamAsync().ConfigureAwait(false), prefix));
            }
            protected override async Task SerializeToStreamAsync(Stream target, TransportContext context)
            { await (await CreateContentReadStreamAsync().ConfigureAwait(false)).CopyToAsync(target).ConfigureAwait(false); }
            protected override bool TryComputeLength(out long length) { length = 0; return false; }
            protected override void Dispose(bool disposing)
            {
                if (disposing) { stream?.Dispose(); original.Dispose(); }
                base.Dispose(disposing);
            }
        }

        private sealed class OllamaTeeStream : Stream
        {
            internal const int Limit = 1024 * 1024;
            private readonly Stream inner;
            private readonly string prefix;
            private readonly MemoryStream captured = new MemoryStream();
            private readonly object captureGate = new object();
            private long observed;
            private bool ended, disposed;
            internal OllamaTeeStream(Stream inner, string prefix) { this.inner = inner; this.prefix = prefix; }
            private int Observe(byte[] buffer, int offset, int count)
            {
                lock (captureGate)
                {
                    if (disposed) return count;
                    observed += count; ended |= count == 0;
                    int keep = Math.Min(count, Limit - (int)captured.Length);
                    if (keep > 0) captured.Write(buffer, offset, keep);
                }
                return count;
            }
            public override int Read(byte[] buffer, int offset, int count) { return Observe(buffer, offset, inner.Read(buffer, offset, count)); }
            public override async Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken token)
            { return Observe(buffer, offset, await inner.ReadAsync(buffer, offset, count, token).ConfigureAwait(false)); }
            protected override void Dispose(bool disposing)
            {
                if (disposing)
                {
                    lock (captureGate)
                    {
                        if (disposed) return;
                        disposed = true;
                        SaveWire(prefix + "-response.bin", captured.ToArray());
                        SaveWire(prefix + "-read-summary.json", Encoding.UTF8.GetBytes(new JavaScriptSerializer().Serialize(new {
                            ObservedBytes = observed, CapturedBytes = captured.Length, Truncated = observed > Limit, EndOfStreamObserved = ended
                        })));
                        captured.Dispose();
                    }
                    inner.Dispose();
                }
                base.Dispose(disposing);
            }
            public override bool CanRead => inner.CanRead;
            public override bool CanSeek => false;
            public override bool CanWrite => false;
            public override long Length => throw new NotSupportedException();
            public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
            public override void Flush() => inner.Flush();
            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
            public override void SetLength(long value) => throw new NotSupportedException();
            public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        }
    }
}
