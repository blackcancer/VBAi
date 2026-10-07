using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Infrastructure.Diagnostics
{
    /// <summary>Explicit opt-in evidence for fixture-owned synthetic HTTP bodies; never a provider transport substitute.</summary>
    internal sealed class OllamaSyntheticWireCapture : IDisposable
    {
        internal const int Limit = 1024 * 1024;
        internal const string UiFlag = "VBAi_OLLAMA_UI_CAPTURE_WIRE";
        internal const string UiResults = "VBAi_OLLAMA_UI_RESULTS";
        internal const string HeadlessFlag = "VBAi_OLLAMA_HEADLESS_CAPTURE_WIRE";
        internal const string HeadlessResults = "VBAi_OLLAMA_HEADLESS_RESULTS";
        private readonly Func<HttpMessageHandler> previous;
        private readonly Action<Func<HttpMessageHandler>> setFactory;
        private readonly string root, fixture;
        private readonly Uri endpoint;
        private readonly Action<string, byte[]> save;
        private readonly Func<HttpContent, string, Action<string, byte[]>, HttpContent> wrapResponse;
        private readonly string scopeId = Guid.NewGuid().ToString("N");
        private string phase = "fixture-start";
        private bool disposed;
        internal bool Enabled { get; }

        internal static OllamaSyntheticWireCapture ForUiFixture()
            => FromEnvironment(true, "detached-ui", Environment.GetEnvironmentVariable);

        internal static OllamaSyntheticWireCapture ForHeadlessFixture(string fixture)
        {
            if (fixture != "cancellation-recovery" && fixture != "synthetic-tool-roundtrip")
                throw new InvalidOperationException("Only the two synthetic headless fixtures may enable capture.");
            return FromEnvironment(false, fixture, Environment.GetEnvironmentVariable);
        }

        internal static OllamaSyntheticWireCapture FromEnvironment(bool ui, string fixture, Func<string, string> read,
            Func<Func<HttpMessageHandler>> getFactory = null, Action<Func<HttpMessageHandler>> setFactory = null,
            Action<string, byte[]> save = null)
        {
            bool enabled = read(ui ? UiFlag : HeadlessFlag) == "1";
            string root = enabled ? read(ui ? UiResults : HeadlessResults) ??
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ui ? "ollama-ui-diagnostics" : "ollama-headless-diagnostics") : null;
            return new OllamaSyntheticWireCapture(enabled, true, root,
                enabled ? OllamaQualificationEndpoint.Parse(read(OllamaQualificationEndpoint.EnvironmentName)) : null, fixture,
                getFactory ?? (() => LlmChatClient.HttpHandlerFactory), setFactory ?? (value => LlmChatClient.HttpHandlerFactory = value), save);
        }

        internal OllamaSyntheticWireCapture(bool enabled, bool syntheticFixture, string root, Uri endpoint,
            string fixture, Func<Func<HttpMessageHandler>> getFactory, Action<Func<HttpMessageHandler>> setFactory,
            Action<string, byte[]> save = null,
            Func<HttpContent, string, Action<string, byte[]>, HttpContent> wrapResponse = null)
        {
            Enabled = enabled;
            if (!enabled) return;
            if (!syntheticFixture) throw new InvalidOperationException("Capture requires an explicitly synthetic fixture.");
            if (fixture != "detached-ui" && fixture != "cancellation-recovery" && fixture != "synthetic-tool-roundtrip")
                throw new InvalidOperationException("Unknown synthetic fixture scope.");
            if (endpoint == null || !endpoint.IsAbsoluteUri)
                throw new InvalidOperationException("An exact canonical loopback endpoint is required.");
            OllamaQualificationEndpoint.Parse(endpoint.AbsoluteUri);
            if (String.IsNullOrWhiteSpace(root) || !Path.IsPathRooted(root) || Path.GetPathRoot(root).Length != 3)
                throw new InvalidOperationException("A local absolute diagnostic root is required.");
            this.root = Path.GetFullPath(root); this.endpoint = endpoint; this.fixture = fixture;
            this.setFactory = setFactory ?? throw new ArgumentNullException(nameof(setFactory));
            previous = (getFactory ?? throw new ArgumentNullException(nameof(getFactory)))();
            if (previous == null) throw new InvalidOperationException("The original HTTP factory is missing.");
            this.save = save ?? SaveNew;
            this.wrapResponse = wrapResponse ?? ((content, prefix, write) => new Content(content, prefix, write));
            if (save == null) Directory.CreateDirectory(this.root);
            try { setFactory(() => CreateHandler(previous())); }
            catch (Exception primary)
            {
                try { setFactory(previous); }
                catch (Exception restoration) { throw new AggregateException("Synthetic capture factory installation and restoration both failed.", primary, restoration); }
                throw;
            }
        }

        internal HttpMessageHandler CreateHandler(HttpMessageHandler inner)
        {
            var real = inner as HttpClientHandler;
            if (real == null || real.AllowAutoRedirect)
            {
                try { inner?.Dispose(); } catch (Exception diagnostic) { Notice(diagnostic); }
                throw new InvalidOperationException("Synthetic capture requires the real HttpClientHandler with redirects disabled.");
            }
            return new Handler(real, this);
        }

        internal void SetPhase(string value)
        {
            if (Enabled) phase = value?.Substring(0, Math.Min(value.Length, 256));
        }

        /// <summary>Records the synthetic parsed argument shape before an assertion, without coercion or native tool dispatch.</summary>
        internal void RecordArguments(string stage, object rawArguments, object parsedArguments)
        {
            if (!Enabled) return;
            try
            {
                var dictionary = parsedArguments as IDictionary<string, object>;
                object marker = dictionary != null && dictionary.TryGetValue("marker", out var found) ? found : null;
                var value = new
                {
                    ScopeId = scopeId,
                    Fixture = fixture,
                    Stage = stage,
                    Utc = DateTime.UtcNow.ToString("o"),
                    RawArgumentType = rawArguments?.GetType().FullName,
                    RawArgumentJson = rawArguments,
                    ParsedArgumentType = parsedArguments?.GetType().FullName,
                    ParsedArgumentJson = parsedArguments,
                    MarkerType = marker?.GetType().FullName,
                    MarkerJson = marker,
                    NativeToolsExecuted = 0,
                    Scope = "Fixture-owned synthetic arguments only; no coercion or historical causal claim"
                };
                byte[] bytes = Json(value);
                string prefix = Prefix("arguments");
                Save(prefix + ".json", bytes.Take(Limit).ToArray());
                Save(prefix + "-summary.json", Json(new { Bytes = bytes.Length, CapturedBytes = Math.Min(bytes.Length, Limit), Truncated = bytes.Length > Limit }));
            }
            catch (Exception diagnostic) { Notice(diagnostic); }
        }

        internal void RecordCompletion(string stage, LlmChatClient client, Exception error = null)
        {
            if (!Enabled) return;
            try
            {
                Save(Prefix("completion") + ".json", Json(new
                {
                    ScopeId = scopeId,
                    Fixture = fixture,
                    Stage = stage,
                    Utc = DateTime.UtcNow.ToString("o"),
                    Mvid = typeof(LlmChatClient).Module.ModuleVersionId.ToString("D"),
                    StreamDiagnostics = client?.LastStreamDiagnostics?.Snapshot(),
                    ErrorType = error?.GetType().FullName,
                    ErrorHResult = error?.HResult,
                    NativeToolsExecuted = 0
                }));
            }
            catch (Exception diagnostic) { Notice(diagnostic); }
        }

        internal async Task<IDictionary<string, object>> CompleteAsync(string stage, LlmChatClient client, IList<object> messages, object[] tools)
        {
            SetPhase(stage);
            try
            {
                var result = await client.CompleteAsync(messages, tools).ConfigureAwait(false);
                RecordCompletion(stage + " returned", client);
                return result;
            }
            catch (Exception error)
            {
                RecordCompletion(stage + " threw", client, error);
                throw;
            }
        }

        internal async Task<string> ObserveRequestAsync(HttpRequestMessage request)
        {
            OllamaQualificationEndpoint.RequireWireUri(request?.RequestUri, endpoint);
            bool chat = request.RequestUri.AbsolutePath == endpoint.AbsolutePath;
            if ((chat && request.Method != HttpMethod.Post) || (!chat && request.Method != HttpMethod.Get))
                throw new InvalidOperationException("Capture permits only POST chat and GET catalogue.");
            if (!chat && request.Content != null) throw new InvalidOperationException("Synthetic catalogue requests have no body.");
            if (request.Content != null && !(request.Content is StringContent))
                throw new InvalidOperationException("Synthetic requests must have buffered StringContent, never an arbitrary source stream.");
            string prefix = Prefix("wire");
            try
            {
                byte[] bytes = request.Content == null ? new byte[0] : await request.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
                if (request.Content != null) Save(prefix + "-request.json", bytes.Take(Limit).ToArray());
                Save(prefix + "-intent.json", Json(new
                {
                    ScopeId = scopeId,
                    Fixture = fixture,
                    Phase = phase,
                    Utc = DateTime.UtcNow.ToString("o"),
                    Method = request.Method.Method,
                    Path = request.RequestUri.AbsolutePath,
                    LoopbackPort = endpoint.Port,
                    RequestBytes = bytes.Length,
                    CapturedRequestBytes = Math.Min(bytes.Length, Limit),
                    RequestTruncated = bytes.Length > Limit,
                    CaptureByteLimit = Limit,
                    Mvid = typeof(LlmChatClient).Module.ModuleVersionId.ToString("D"),
                    Scope = "Synthetic fixture bodies only; no headers, credentials or user history recorded"
                }));
            }
            catch (Exception diagnostic) { Notice(diagnostic); }
            return prefix;
        }

        internal void ObserveResponse(HttpResponseMessage response, string prefix)
        {
            try
            {
                Save(prefix + "-metadata.json", Json(new
                {
                    ScopeId = scopeId,
                    Utc = DateTime.UtcNow.ToString("o"),
                    Status = (int)response.StatusCode,
                    LoopbackPort = endpoint.Port,
                    ContentType = response.Content?.Headers.ContentType?.MediaType,
                    Mvid = typeof(LlmChatClient).Module.ModuleVersionId.ToString("D"),
                    CaptureByteLimit = Limit,
                    ContentPresent = response.Content != null,
                    Scope = "Synthetic fixture bodies only; no headers, credentials or user history recorded"
                }));
                if (response.Content != null)
                {
                    var replacement = wrapResponse(response.Content, prefix, Save);
                    if (replacement == null) throw new InvalidOperationException("A passive response wrapper may not remove the original body.");
                    response.Content = replacement;
                }
            }
            catch (Exception diagnostic) { Notice(diagnostic); }
        }

        private string Prefix(string kind) => Path.Combine(root, kind + "-" + Guid.NewGuid().ToString("N"));
        internal static byte[] Json(object value) => Encoding.UTF8.GetBytes(new JavaScriptSerializer { MaxJsonLength = 4 * Limit }.Serialize(value));
        private static void SaveNew(string path, byte[] bytes)
        {
            using (var file = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.Read)) file.Write(bytes, 0, bytes.Length);
        }
        private void Save(string path, byte[] bytes) { try { save(path, bytes); } catch (Exception error) { Notice(error); } }
        internal void RecordDeliveryError(string prefix, Exception error)
        {
            try { Save(prefix + "-delivery-error.json", Json(new { ErrorType = error.GetType().FullName, ErrorHResult = error.HResult, Outcome = "Unknown delivery; no retry inferred" })); }
            catch (Exception diagnostic) { Notice(diagnostic); }
        }
        internal static void Notice(Exception error)
        { try { Console.WriteLine("Passive Ollama synthetic evidence failed: " + error.GetType().Name); } catch { /* Diagnostics cannot replace the provider outcome. */ } }
        public void Dispose() { if (disposed) return; disposed = true; if (Enabled) setFactory(previous); }

        private sealed class Handler : DelegatingHandler
        {
            private readonly OllamaSyntheticWireCapture owner;
            internal Handler(HttpClientHandler real, OllamaSyntheticWireCapture owner) : base(real) { this.owner = owner; }
            protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken token)
            {
                string prefix = await owner.ObserveRequestAsync(request).ConfigureAwait(false);
                try
                {
                    var response = await base.SendAsync(request, token).ConfigureAwait(false);
                    owner.ObserveResponse(response, prefix);
                    return response;
                }
                catch (Exception error)
                {
                    owner.RecordDeliveryError(prefix, error);
                    throw;
                }
            }
        }

        internal sealed class Content : HttpContent
        {
            private readonly HttpContent original;
            private readonly string prefix;
            private readonly Action<string, byte[]> save;
            private Stream stream;
            private bool disposed;
            internal Content(HttpContent original, string prefix, Action<string, byte[]> save)
            {
                this.original = original; this.prefix = prefix; this.save = save;
                foreach (var header in original.Headers) Headers.TryAddWithoutValidation(header.Key, header.Value);
            }
            protected override async Task<Stream> CreateContentReadStreamAsync()
                => stream ?? (stream = new TeeStream(await original.ReadAsStreamAsync().ConfigureAwait(false), prefix, save, true));
            protected override async Task SerializeToStreamAsync(Stream target, TransportContext context)
                => await (await CreateContentReadStreamAsync().ConfigureAwait(false)).CopyToAsync(target).ConfigureAwait(false);
            protected override bool TryComputeLength(out long length) { length = 0; return false; }
            protected override void Dispose(bool disposing)
            {
                if (!disposing || disposed) return;
                disposed = true;
                try
                {
                    if (stream != null) stream.Dispose();
                    else
                    {
                        try { save(prefix + "-response.bin", new byte[0]); } catch (Exception error) { Notice(error); }
                        try
                        {
                            save(prefix + "-read-summary.json", Json(new
                            {
                                ObservedBytes = 0,
                                CapturedBytes = 0,
                                Truncated = false,
                                EndOfStreamObserved = false,
                                ResponseReadStarted = false,
                                ReadInFlightAtDispose = false
                            }));
                        }
                        catch (Exception error) { Notice(error); }
                    }
                }
                finally { try { original.Dispose(); } finally { base.Dispose(disposing); } }
            }
        }

        internal sealed class TeeStream : Stream
        {
            private readonly Stream inner;
            private readonly string prefix;
            private readonly Action<string, byte[]> save;
            private readonly bool leaveOpen;
            private readonly MemoryStream captured = new MemoryStream();
            private readonly object gate = new object();
            private long observed;
            private bool ended, disposed;
            private int pending;
            private int readCalls;
            private Exception readError;
            internal TeeStream(Stream inner, string prefix, Action<string, byte[]> save, bool leaveOpen = false)
            { this.inner = inner; this.prefix = prefix; this.save = save; this.leaveOpen = leaveOpen; }
            private void BeginRead() { lock (gate) { if (disposed) throw new ObjectDisposedException(nameof(TeeStream)); pending++; readCalls++; } }
            private int Observe(byte[] buffer, int offset, int count, int requested)
            {
                lock (gate)
                {
                    if (!disposed)
                    {
                        observed += count; ended |= count == 0 && requested > 0;
                        int keep = Math.Min(count, Limit - (int)captured.Length);
                        // Capturing an already delivered read may not replace its transport outcome.
                        try { if (keep > 0) captured.Write(buffer, offset, keep); }
                        catch (Exception diagnostic) { Notice(diagnostic); }
                    }
                }
                return count;
            }
            private void Failed(Exception error) { lock (gate) { if (!disposed) readError = error; } }
            private void EndRead() { lock (gate) { pending--; } }
            public override int Read(byte[] buffer, int offset, int count)
            {
                BeginRead();
                try { return Observe(buffer, offset, inner.Read(buffer, offset, count), count); }
                catch (Exception error) { Failed(error); throw; }
                finally { EndRead(); }
            }
            public override async Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken token)
            {
                BeginRead();
                try { return Observe(buffer, offset, await inner.ReadAsync(buffer, offset, count, token).ConfigureAwait(false), count); }
                catch (Exception error) { Failed(error); throw; }
                finally { EndRead(); }
            }
            protected override void Dispose(bool disposing)
            {
                if (!disposing) return;
                lock (gate)
                {
                    if (disposed) return;
                    disposed = true;
                    try { save(prefix + "-response.bin", captured.ToArray()); } catch (Exception error) { Notice(error); }
                    try
                    {
                        save(prefix + "-read-summary.json", Json(new
                        {
                            ObservedBytes = observed,
                            CapturedBytes = captured.Length,
                            Truncated = observed > Limit,
                            EndOfStreamObserved = ended,
                            ReadInFlightAtDispose = pending > 0,
                            ResponseReadStarted = readCalls > 0,
                            ReadErrorType = readError?.GetType().FullName,
                            ReadErrorHResult = readError?.HResult
                        }));
                    }
                    catch (Exception error) { Notice(error); }
                    captured.Dispose();
                }
                try { if (!leaveOpen) inner.Dispose(); } finally { base.Dispose(disposing); }
            }
            public override bool CanRead => !disposed && inner.CanRead;
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
