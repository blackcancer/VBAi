using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace VBAi.Tests.Integration
{
    /// <summary>Explicit synthetic loopback observer for requests originating inside the installed native host.</summary>
    public sealed class OllamaQualificationProxy : IDisposable
    {
        private const int CaptureLimit = 1024 * 1024;
        private readonly HttpListener listener = new HttpListener();
        private readonly HttpClient client;
        private readonly Uri backend;
        private readonly string root;
        private readonly Task accepting;
        private readonly List<Task> requests = new List<Task>();
        private int sequence;
        private bool disposed;
        public string Endpoint { get; private set; }

        public OllamaQualificationProxy(string backendEndpoint, int port, string evidenceRoot)
        {
            var selected = OllamaQualificationEndpoint.Parse(backendEndpoint);
            backend = new Uri(selected.GetLeftPart(UriPartial.Authority));
            if (port <= 0 || port >= 65536 || port == selected.Port) throw new ArgumentOutOfRangeException(nameof(port));
            root = Path.GetFullPath(evidenceRoot);
            if (!Path.IsPathRooted(evidenceRoot) || Directory.Exists(root)) throw new InvalidOperationException("A fresh absolute proxy evidence directory is required.");
            Directory.CreateDirectory(root);
            Endpoint = "http://127.0.0.1:" + port + "/v1/chat/completions";
            client = new HttpClient(new HttpClientHandler { AllowAutoRedirect = false, UseProxy = false })
            {
                Timeout = TimeSpan.FromSeconds(150)
            };
            listener.Prefixes.Add("http://127.0.0.1:" + port + "/"); listener.Start();
            accepting = Accept();
        }

        private async Task Accept()
        {
            while (listener.IsListening)
            {
                HttpListenerContext context;
                try { context = await listener.GetContextAsync().ConfigureAwait(false); }
                catch (Exception) when (disposed) { break; }
                Task request = Forward(context, Interlocked.Increment(ref sequence));
                lock (requests) requests.Add(request);
            }
        }

        private async Task Forward(HttpListenerContext context, int number)
        {
            string prefix = Path.Combine(root, number.ToString("D4"));
            var receipt = new Dictionary<string, object>
            {
                ["State"] = "STARTED",
                ["Sequence"] = number,
                ["Path"] = context.Request.Url.AbsolutePath,
                ["StartedUtc"] = DateTime.UtcNow.ToString("o")
            };
            Action flush = () => File.WriteAllText(prefix + "-receipt.json", new JavaScriptSerializer().Serialize(receipt));
            flush();
            using (var captured = new MemoryStream())
            {
                try
                {
                    if (number > 64) throw new InvalidOperationException("Frozen qualification request budget exceeded.");
                    string route = context.Request.Url.AbsolutePath;
                    if (context.Request.RemoteEndPoint == null || !IPAddress.IsLoopback(context.Request.RemoteEndPoint.Address) ||
                        (route != "/api/tags" && route != "/v1/chat/completions") ||
                        context.Request.Url.Query.Length != 0 || (route == "/api/tags" ? context.Request.HttpMethod != "GET" : context.Request.HttpMethod != "POST"))
                        throw new InvalidOperationException("Only exact synthetic loopback catalogue/chat routes are authorized.");
                    byte[] body;
                    using (var requestBody = new MemoryStream())
                    {
                        var buffer = new byte[8192]; int count;
                        using (var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(10)))
                        {
                            while ((count = await context.Request.InputStream.ReadAsync(buffer, 0, buffer.Length, deadline.Token).ConfigureAwait(false)) != 0)
                            {
                                if (requestBody.Length + count > CaptureLimit) throw new InvalidOperationException("Synthetic request exceeds the capture boundary.");
                                requestBody.Write(buffer, 0, count);
                            }
                        }
                        body = requestBody.ToArray();
                    }
                    if (route == "/v1/chat/completions")
                    {
                        string text = new UTF8Encoding(false, true).GetString(body);
                        if (text.IndexOf("synthetic", StringComparison.OrdinalIgnoreCase) < 0)
                            throw new InvalidOperationException("The explicit synthetic qualification marker is absent; body is not retained or forwarded.");
                        File.WriteAllBytes(prefix + "-request.json", body);
                    }
                    using (var request = new HttpRequestMessage(route == "/api/tags" ? HttpMethod.Get : HttpMethod.Post, new Uri(backend, route)))
                    {
                        if (route != "/api/tags")
                        {
                            request.Content = new ByteArrayContent(body);
                            request.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/json");
                        }
                        using (var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(150)))
                        using (var response = await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, deadline.Token).ConfigureAwait(false))
                        using (var stream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false))
                        {
                            receipt["HttpStatus"] = (int)response.StatusCode;
                            context.Response.StatusCode = (int)response.StatusCode;
                            context.Response.ContentType = response.Content.Headers.ContentType?.ToString() ?? "application/octet-stream";
                            context.Response.SendChunked = true;
                            var buffer = new byte[8192]; int count; long bytes = 0;
                            while ((count = await stream.ReadAsync(buffer, 0, buffer.Length, deadline.Token).ConfigureAwait(false)) != 0)
                            {
                                bytes += count;
                                int keep = Math.Min(count, CaptureLimit - (int)captured.Length);
                                if (keep > 0) captured.Write(buffer, 0, keep);
                                receipt["ObservedResponseBytes"] = bytes; receipt["CaptureTruncated"] = bytes > CaptureLimit;
                                await context.Response.OutputStream.WriteAsync(buffer, 0, count, deadline.Token).ConfigureAwait(false);
                                await context.Response.OutputStream.FlushAsync(deadline.Token).ConfigureAwait(false);
                            }
                            receipt["BackendEofObserved"] = true; receipt["State"] = "COMPLETE";
                        }
                    }
                }
                catch (Exception error)
                {
                    receipt["State"] = "ERROR_OR_EARLY_DISPOSAL"; receipt["Error"] = error.ToString();
                    try { context.Response.StatusCode = 502; } catch (InvalidOperationException) { }
                }
                finally
                {
                    if (captured.Length > 0) File.WriteAllBytes(prefix + "-response.bin", captured.ToArray());
                    receipt["CompletedUtc"] = DateTime.UtcNow.ToString("o"); flush();
                    try { context.Response.Close(); } catch (HttpListenerException) { }
                }
            }
        }

        public bool AllRequestsSettled
        {
            get { lock (requests) return requests.TrueForAll(request => request.IsCompleted); }
        }

        public void Dispose()
        {
            if (disposed) return;
            if (!AllRequestsSettled) throw new InvalidOperationException("Proxy requests are unsettled; retain the qualification server.");
            disposed = true; listener.Stop();
            if (!accepting.Wait(TimeSpan.FromSeconds(10))) throw new InvalidOperationException("Proxy accept worker is unsettled.");
            listener.Close(); client.Dispose();
        }
    }
}
