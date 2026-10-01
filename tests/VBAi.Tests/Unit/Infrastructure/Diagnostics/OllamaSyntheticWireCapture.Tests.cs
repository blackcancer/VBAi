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
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Infrastructure.Diagnostics;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Detached synthetic capture contracts only; none of these cases opens a socket or host.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class OllamaSyntheticWireCaptureTests
    {
        private static readonly Uri Endpoint = new Uri("http://127.0.0.1:52541/v1/chat/completions");
        private static string Root => Path.Combine(Path.GetTempPath(), "VBAiSyntheticWire-" + Guid.NewGuid().ToString("N"));

        [TestMethod]
        public void DisabledScopeDoesNotReadFactoryValidateConfigurationOrCreateEvidence()
        {
            int calls = 0;
            using (var scope = new OllamaSyntheticWireCapture(false, false, null, null, null,
                () => { calls++; throw new Exception(); }, value => calls++)) {
                scope.SetPhase("ignored"); scope.RecordArguments("ignored", new object(), new object());
                scope.RecordCompletion("ignored", null); scope.Dispose();
                Assert.IsFalse(scope.Enabled);
            }
            Assert.AreEqual(0, calls);
        }

        [DataTestMethod, DataRow(null), DataRow(""), DataRow("true"), DataRow("0"), DataRow(" 1")]
        public void EnvironmentCaptureRequiresExactExplicitOne(string flag)
        {
            var reads = new List<string>();
            using (var scope = OllamaSyntheticWireCapture.FromEnvironment(false, "synthetic-tool-roundtrip",
                name => { reads.Add(name); return flag; }, () => { Assert.Fail("Disabled capture read the factory."); return null; }, value => Assert.Fail()))
                Assert.IsFalse(scope.Enabled);
            CollectionAssert.AreEqual(new[] { OllamaSyntheticWireCapture.HeadlessFlag }, reads);
        }

        [DataTestMethod, DataRow(true), DataRow(false)]
        public void UiAndHeadlessFlagsSelectTheirOwnResultsAndRestoreTheFactory(bool ui)
        {
            var sink = new Sink();
            Func<HttpMessageHandler> original = () => new HttpClientHandler { AllowAutoRedirect = false };
            Func<HttpMessageHandler> current = original;
            var values = new Dictionary<string, string> {
                [ui ? OllamaSyntheticWireCapture.UiFlag : OllamaSyntheticWireCapture.HeadlessFlag] = "1",
                [ui ? OllamaSyntheticWireCapture.UiResults : OllamaSyntheticWireCapture.HeadlessResults] = Root,
                [OllamaQualificationEndpoint.EnvironmentName] = Endpoint.AbsoluteUri
            };
            using (var scope = OllamaSyntheticWireCapture.FromEnvironment(ui, ui ? "detached-ui" : "synthetic-tool-roundtrip",
                name => values.TryGetValue(name, out var value) ? value : null, () => current, value => current = value, sink.Write)) {
                Assert.IsTrue(scope.Enabled); Assert.AreNotSame(original, current);
                scope.RecordArguments("synthetic", "{}", new Dictionary<string, object>());
            }
            Assert.AreSame(original, current); Assert.AreEqual(2, sink.Files.Count);
        }

        [DataTestMethod, DataRow(""), DataRow("relative"), DataRow("C:relative"), DataRow("\\\\server\\share")]
        public void InvalidEvidenceRootRefusesBeforeFactoryMutation(string root)
            => Refused(root, Endpoint, "synthetic-tool-roundtrip", true);

        [DataTestMethod, DataRow("http://localhost:52541/v1/chat/completions"),
            DataRow("https://127.0.0.1:52541/v1/chat/completions"), DataRow("http://127.0.0.1:52541/api/chat"),
            DataRow("http://name:secret@127.0.0.1:52541/v1/chat/completions"),
            DataRow("http://127.0.0.1:52541/v1/chat/completions?secret=synthetic"),
            DataRow("http://127.0.0.1:52541/v1/chat/completions#synthetic")]
        public void InvalidSelectedEndpointRefusesBeforeFactoryMutation(string raw)
            => Refused(Root, new Uri(raw), "synthetic-tool-roundtrip", true);

        [TestMethod]
        public void UnknownFixtureAndNonSyntheticScopeRefuseBeforeFactoryMutation()
        {
            Refused(Root, Endpoint, "unknown", true);
            Refused(Root, Endpoint, "synthetic-tool-roundtrip", false);
            Refused(Root, null, "synthetic-tool-roundtrip", true);
            Refused(Root, new Uri("relative", UriKind.Relative), "synthetic-tool-roundtrip", true);
            Assert.ThrowsException<InvalidOperationException>(() => OllamaSyntheticWireCapture.ForHeadlessFixture("native-project"));
        }

        private static void Refused(string root, Uri endpoint, string fixture, bool synthetic)
        {
            int factory = 0;
            Assert.ThrowsException<InvalidOperationException>(() => new OllamaSyntheticWireCapture(true, synthetic, root, endpoint, fixture,
                () => { factory++; return () => new HttpClientHandler(); }, value => factory++, new Sink().Write));
            Assert.AreEqual(0, factory);
        }

        [TestMethod]
        public void ScopeRestoreIsIdempotentAndPriorFactorySurvivesAnAssertionFailure()
        {
            Func<HttpMessageHandler> original = () => new HttpClientHandler { AllowAutoRedirect = false };
            Func<HttpMessageHandler> current = original; int sets = 0;
            var primary = new InvalidOperationException("synthetic assertion");
            try {
                using (var scope = new OllamaSyntheticWireCapture(true, true, Root, Endpoint, "synthetic-tool-roundtrip",
                    () => current, value => { sets++; current = value; }, new Sink().Write)) {
                    scope.Dispose(); scope.Dispose(); throw primary;
                }
            } catch (Exception observed) { Assert.AreSame(primary, observed); }
            Assert.AreSame(original, current); Assert.AreEqual(2, sets);
        }

        [TestMethod]
        public void MissingFactoryIsRejectedWithoutInstallingAnOverride()
        {
            int sets = 0;
            Assert.ThrowsException<InvalidOperationException>(() => new OllamaSyntheticWireCapture(true, true, Root, Endpoint,
                "synthetic-tool-roundtrip", () => null, value => sets++, new Sink().Write));
            Assert.AreEqual(0, sets);
        }

        [TestMethod]
        public void PartialFactoryInstallationFailureRestoresPriorFactoryAndPreservesTheOriginalException()
        {
            Func<HttpMessageHandler> original = () => new HttpClientHandler();
            Func<HttpMessageHandler> current = original; int sets = 0;
            var primary = new IOException("synthetic factory installation");
            try {
                new OllamaSyntheticWireCapture(true, true, Root, Endpoint, "synthetic-tool-roundtrip", () => current,
                    value => { current = value; if (++sets == 1) throw primary; }, new Sink().Write);
                Assert.Fail();
            } catch (Exception observed) { Assert.AreSame(primary, observed); }
            Assert.AreSame(original, current); Assert.AreEqual(2, sets);
        }

        [TestMethod]
        public void FactoryInstallationAndRestorationErrorsRemainDistinct()
        {
            var primary = new IOException("synthetic installation"); var cleanup = new InvalidOperationException("synthetic restoration"); int sets = 0;
            var error = Assert.ThrowsException<AggregateException>(() => new OllamaSyntheticWireCapture(true, true, Root, Endpoint,
                "synthetic-tool-roundtrip", () => (() => new HttpClientHandler()), value => { if (++sets == 1) throw primary; throw cleanup; }, new Sink().Write));
            Assert.AreSame(primary, error.InnerExceptions[0]); Assert.AreSame(cleanup, error.InnerExceptions[1]);
        }

        [TestMethod]
        public async Task RealWrappedHandlerRefusesForeignUriBeforeAnyHttpDispatch()
        {
            using (var fixture = new Fixture())
            using (var invoker = new HttpMessageInvoker(fixture.Scope.CreateHandler(new HttpClientHandler { AllowAutoRedirect=false })))
            using (var request = new HttpRequestMessage(HttpMethod.Post, "http://fixture.invalid/v1/chat/completions") { Content=new StringContent("synthetic") }) {
                await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => invoker.SendAsync(request, CancellationToken.None));
                Assert.AreEqual(0, fixture.Sink.Files.Count);
            }
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void CaptureRejectsMockOrRedirectingTransportAndDisposesIt(bool redirects)
        {
            using (var fixture = new Fixture()) {
                var real = new TrackingTransport { AllowAutoRedirect = true };
                var fake = new FakeHandler();
                HttpMessageHandler inner = redirects ? (HttpMessageHandler)real : fake;
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Scope.CreateHandler(inner));
                Assert.AreEqual(1, redirects ? real.Disposals : fake.Disposals);
                Assert.AreEqual(0, fake.Requests); real.Dispose(); fake.Dispose();
            }
        }

        [TestMethod]
        public void RealRedirectDisabledTransportRemainsWrappedAndOwnedWithoutDispatch()
        {
            using (var fixture = new Fixture()) {
                var real = new TrackingTransport { AllowAutoRedirect = false };
                using (var wrapper = fixture.Scope.CreateHandler(real)) Assert.AreEqual(0, real.Disposals);
                Assert.AreEqual(1, real.Disposals);
            }
        }

        [DataTestMethod, DataRow("http://127.0.0.1:11434/v1/chat/completions"),
            DataRow("http://fixture.invalid:52541/v1/chat/completions"), DataRow("http://127.0.0.1:52541/api/pull"),
            DataRow("http://127.0.0.1:52541/api/tags?secret=synthetic")]
        public async Task ForeignPortHostRouteAndQueryRefuseBeforeBodyCapture(string raw)
        {
            using (var fixture = new Fixture())
            using (var request = new HttpRequestMessage(HttpMethod.Post, raw) { Content = new StringContent("synthetic") }) {
                await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => fixture.Scope.ObserveRequestAsync(request));
                Assert.AreEqual(0, fixture.Sink.Files.Count);
            }
        }

        [DataTestMethod, DataRow("GET", "/v1/chat/completions"), DataRow("POST", "/api/tags"), DataRow("DELETE", "/v1/chat/completions")]
        public async Task WrongMethodRefusesBeforeCapture(string method, string path)
        {
            using (var fixture = new Fixture())
            using (var request = new HttpRequestMessage(new HttpMethod(method), "http://127.0.0.1:52541" + path)) {
                await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => fixture.Scope.ObserveRequestAsync(request));
                Assert.AreEqual(0, fixture.Sink.Files.Count);
            }
        }

        [TestMethod]
        public async Task UnbufferedRequestAndCatalogueBodyCannotBeCaptured()
        {
            using (var fixture = new Fixture()) {
                var stream = new CountingStream(Encoding.UTF8.GetBytes("synthetic"));
                using (var request = new HttpRequestMessage(HttpMethod.Post, Endpoint) { Content = new StreamContent(stream) })
                    await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => fixture.Scope.ObserveRequestAsync(request));
                Assert.AreEqual(0, stream.Reads);
                using (var request = new HttpRequestMessage(HttpMethod.Get, "http://127.0.0.1:52541/api/tags") { Content = new StringContent("synthetic") })
                    await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => fixture.Scope.ObserveRequestAsync(request));
                Assert.AreEqual(0, fixture.Sink.Files.Count);
            }
        }

        [TestMethod]
        public async Task RequestEvidenceKeepsExactUtf8AndPhaseButNeverCopiesCredentialHeaders()
        {
            using (var fixture = new Fixture())
            using (var request = new HttpRequestMessage(HttpMethod.Post, Endpoint) { Content = new StringContent("{\"synthetic\":\"été\"}", Encoding.UTF8) }) {
                request.Headers.TryAddWithoutValidation("Authorization", "Bearer private-header-not-for-evidence");
                request.Content.Headers.TryAddWithoutValidation("X-Private", "private-content-header-not-for-evidence");
                fixture.Scope.SetPhase("before-marker-assertion");
                await fixture.Scope.ObserveRequestAsync(request);
                CollectionAssert.AreEqual(await request.Content.ReadAsByteArrayAsync(), fixture.Sink.Single("-request.json"));
                StringAssert.Contains(fixture.Sink.Text, "before-marker-assertion");
                Assert.IsFalse(fixture.Sink.Text.Contains("private-header-not-for-evidence"));
                Assert.IsFalse(fixture.Sink.Text.Contains("private-content-header-not-for-evidence"));
                Assert.AreEqual(52541, fixture.Sink.Object("-intent.json")["LoopbackPort"]);
            }
        }

        [TestMethod]
        public async Task OversizedRequestEvidenceIsBoundedWithoutChangingTheRequestBody()
        {
            using (var fixture = new Fixture())
            using (var request = new HttpRequestMessage(HttpMethod.Post, Endpoint) { Content = new StringContent(new string('x', OllamaSyntheticWireCapture.Limit + 7)) }) {
                await fixture.Scope.ObserveRequestAsync(request);
                Assert.AreEqual(OllamaSyntheticWireCapture.Limit + 7, (await request.Content.ReadAsByteArrayAsync()).Length);
                Assert.AreEqual(OllamaSyntheticWireCapture.Limit, fixture.Sink.Single("-request.json").Length);
                Assert.AreEqual(true, fixture.Sink.Object("-intent.json")["RequestTruncated"]);
            }
        }

        [TestMethod]
        public async Task CatalogueIntentContainsNoInventedRequestBody()
        {
            using (var fixture = new Fixture())
            using (var request = new HttpRequestMessage(HttpMethod.Get, "http://127.0.0.1:52541/api/tags")) {
                await fixture.Scope.ObserveRequestAsync(request);
                Assert.AreEqual(1, fixture.Sink.Files.Count);
                Assert.AreEqual("GET", fixture.Sink.Object("-intent.json")["Method"]);
            }
        }

        [TestMethod]
        public async Task RequestDiagnosticIoFailureDoesNotEscapeOrReplaceBufferedContent()
        {
            using(var fixture=new Fixture(save:(path,bytes)=>{throw new IOException("synthetic diagnostic");}))
            using(var request=new HttpRequestMessage(HttpMethod.Post,Endpoint){Content=new StringContent("synthetic unchanged")}) {
                string prefix=await fixture.Scope.ObserveRequestAsync(request);
                Assert.IsTrue(Path.GetFileName(prefix).StartsWith("wire-"));
                Assert.AreEqual("synthetic unchanged",await request.Content.ReadAsStringAsync());
                fixture.Scope.RecordDeliveryError(prefix,new IOException("synthetic original delivery"));
            }
        }

        [TestMethod]
        public async Task RequestObservationErrorDoesNotThrowBeforeTheOriginalTransportWouldSeeTheRequest()
        {
            using(var fixture=new Fixture())
            using(var request=new HttpRequestMessage(HttpMethod.Post,Endpoint){Content=new StringContent("synthetic")}) {
                request.Content.Dispose();
                Assert.IsNotNull(await fixture.Scope.ObserveRequestAsync(request));
                Assert.AreEqual(0,fixture.Sink.Files.Count);
            }
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public async Task PassiveResponseWrapperFailureLeavesOriginalHttpOutcomeAndBodyUnchanged(bool returnNull)
        {
            using (var fixture = new Fixture(wrap: (content, prefix, save) => {
                if (returnNull) return null; throw new IOException("synthetic evidence error");
            }))
            using (var response = new HttpResponseMessage(HttpStatusCode.Accepted) { Content = new StringContent("synthetic original body") }) {
                var original = response.Content;
                fixture.Scope.ObserveResponse(response, "synthetic");
                Assert.AreSame(original, response.Content); Assert.AreEqual(HttpStatusCode.Accepted, response.StatusCode);
                Assert.AreEqual("synthetic original body", await response.Content.ReadAsStringAsync());
            }
        }

        [TestMethod]
        public void MissingResponseBodyOnlyRecordsMetadata()
        {
            using (var fixture = new Fixture())
            using (var response = new HttpResponseMessage(HttpStatusCode.NoContent)) {
                fixture.Scope.ObserveResponse(response, "synthetic");
                Assert.IsNull(response.Content); Assert.AreEqual(false, fixture.Sink.Object("-metadata.json")["ContentPresent"]);
            }
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public async Task TeePreservesReadOffsetsActualBytesAndEofForSyncAndAsync(bool asynchronous)
        {
            var sink = new Sink(); byte[] expected = Encoding.UTF8.GetBytes("synthetic été\nsecond");
            using (var inner = new CountingStream(expected))
            using (var tee = new OllamaSyntheticWireCapture.TeeStream(inner, "synthetic", sink.Write)) {
                var actual = new MemoryStream(); var buffer = new byte[9]; int count;
                do {
                    count = asynchronous ? await tee.ReadAsync(buffer, 2, 5, CancellationToken.None) : tee.Read(buffer, 2, 5);
                    actual.Write(buffer, 2, count);
                } while (count > 0);
                CollectionAssert.AreEqual(expected, actual.ToArray());
            }
            CollectionAssert.AreEqual(expected, sink.Single("-response.bin"));
            Assert.AreEqual(true, sink.Object("-read-summary.json")["EndOfStreamObserved"]);
            Assert.AreEqual(false, sink.Object("-read-summary.json")["Truncated"]);
        }

        [TestMethod]
        public void EarlyDisposeAndZeroLengthReadDoNotInventEofAndAreIdempotent()
        {
            var sink = new Sink(); var inner = new CountingStream(new byte[] { 1, 2, 3 });
            var tee = new OllamaSyntheticWireCapture.TeeStream(inner, "synthetic", sink.Write);
            Assert.AreEqual(0, tee.Read(new byte[1], 0, 0));
            Assert.AreEqual(1, tee.Read(new byte[1], 0, 1));
            tee.Dispose(); tee.Dispose();
            Assert.AreEqual(false, sink.Object("-read-summary.json")["EndOfStreamObserved"]);
            Assert.AreEqual(1, inner.Disposals); Assert.AreEqual(2, sink.Files.Count);
            Assert.IsFalse(tee.CanRead); Assert.IsFalse(tee.CanWrite); Assert.IsFalse(tee.CanSeek);
            Assert.ThrowsException<ObjectDisposedException>(() => tee.Read(new byte[1], 0, 1));
        }

        [TestMethod]
        public void TeeLimitDoesNotTruncateBytesDeliveredToTheProductionReader()
        {
            var sink = new Sink(); byte[] all = Enumerable.Repeat((byte)42, OllamaSyntheticWireCapture.Limit + 13).ToArray();
            using (var tee = new OllamaSyntheticWireCapture.TeeStream(new MemoryStream(all), "synthetic", sink.Write)) {
                using (var received = new MemoryStream()) { tee.CopyTo(received); CollectionAssert.AreEqual(all, received.ToArray()); }
            }
            Assert.AreEqual(OllamaSyntheticWireCapture.Limit, sink.Single("-response.bin").Length);
            Assert.AreEqual(all.Length, sink.Object("-read-summary.json")["ObservedBytes"]);
            Assert.AreEqual(true, sink.Object("-read-summary.json")["Truncated"]);
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public async Task ReadExceptionsRemainTheOriginalInstanceAndDoNotBecomeEof(bool cancellation)
        {
            var sink = new Sink(); Exception primary = cancellation ? (Exception)new OperationCanceledException("synthetic") : new IOException("synthetic");
            using (var tee = new OllamaSyntheticWireCapture.TeeStream(new FailingReadStream(primary), "synthetic", sink.Write)) {
                try { await tee.ReadAsync(new byte[3], 0, 3, CancellationToken.None); Assert.Fail(); }
                catch (Exception observed) { Assert.AreSame(primary, observed); }
            }
            Assert.AreEqual(false, sink.Object("-read-summary.json")["EndOfStreamObserved"]);
            Assert.AreEqual(primary.GetType().FullName, sink.Object("-read-summary.json")["ReadErrorType"]);
        }

        [TestMethod]
        public async Task PendingReadAtDisposeIsRecordedWithoutInventingTerminalOrCapturingLateBytes()
        {
            var sink = new Sink(); var inner = new PendingReadStream();
            var tee = new OllamaSyntheticWireCapture.TeeStream(inner, "synthetic", sink.Write, true);
            var pending = tee.ReadAsync(new byte[1], 0, 1, CancellationToken.None);
            tee.Dispose();
            Assert.AreEqual(true, sink.Object("-read-summary.json")["ReadInFlightAtDispose"]);
            Assert.AreEqual(false, sink.Object("-read-summary.json")["EndOfStreamObserved"]);
            inner.Result.SetResult(1); Assert.AreEqual(1, await pending);
            Assert.AreEqual(0, sink.Single("-response.bin").Length); inner.Dispose();
        }

        [TestMethod]
        public void EvidenceWriteFailuresNeverPreventUnderlyingDisposeOrMaskItsPrimaryException()
        {
            var primary = new IOException("synthetic original dispose"); var inner = new FailingDisposeStream(primary);
            var tee = new OllamaSyntheticWireCapture.TeeStream(inner, "synthetic", (path, bytes) => { throw new UnauthorizedAccessException(); });
            try { tee.Dispose(); Assert.Fail(); } catch (Exception observed) { Assert.AreSame(primary, observed); }
            Assert.AreEqual(1, inner.Disposals); tee.Dispose(); Assert.AreEqual(1, inner.Disposals);
        }

        [TestMethod]
        public async Task ContentDoesNotReadAheadPreservesHeadersAndRecordsUnreadDisposal()
        {
            var sink = new Sink(); var stream = new CountingStream(Encoding.UTF8.GetBytes("synthetic"));
            var original = new StreamContent(stream); original.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("text/event-stream");
            using (var wrapped = new OllamaSyntheticWireCapture.Content(original, "synthetic", sink.Write)) {
                Assert.AreEqual(0, stream.Reads); Assert.AreEqual("text/event-stream", wrapped.Headers.ContentType.MediaType);
                var available = await wrapped.ReadAsStreamAsync(); Assert.AreEqual(0, stream.Reads);
                Assert.IsTrue(available.CanRead);
            }
            Assert.AreEqual(false, sink.Object("-read-summary.json")["EndOfStreamObserved"]);
            Assert.AreEqual(false, sink.Object("-read-summary.json")["ResponseReadStarted"]);
        }

        [TestMethod]
        public void NeverOpenedContentStillKeepsAnExplicitUnreadSummary()
        {
            var sink = new Sink(); var stream = new CountingStream(new byte[] { 42 });
            var content = new OllamaSyntheticWireCapture.Content(new StreamContent(stream), "synthetic", sink.Write);
            content.Dispose(); content.Dispose();
            Assert.AreEqual(0, stream.Reads); Assert.AreEqual(2, sink.Files.Count);
            Assert.AreEqual(false, sink.Object("-read-summary.json")["ResponseReadStarted"]);
        }

        [TestMethod]
        public async Task BufferedStringReadIsUnchangedAndAddsNoExtraInnerDisposalComparedWithOriginalContent()
        {
            byte[] bytes = Encoding.UTF8.GetBytes("synthetic utf8 été"); var sink = new Sink();
            var baseline = new CountingStream(bytes); var baselineContent = new StreamContent(baseline);
            await baselineContent.ReadAsStreamAsync(); baselineContent.Dispose();
            var captured = new CountingStream(bytes);
            var wrapper = new OllamaSyntheticWireCapture.Content(new StreamContent(captured), "synthetic", sink.Write);
            Assert.AreEqual("synthetic utf8 été", await wrapper.ReadAsStringAsync());
            wrapper.Dispose(); wrapper.Dispose();
            Assert.AreEqual(baseline.Disposals, captured.Disposals);
            CollectionAssert.AreEqual(bytes, sink.Single("-response.bin"));
            Assert.AreEqual(true, sink.Object("-read-summary.json")["EndOfStreamObserved"]);
        }

        [TestMethod]
        public void TeeUnsupportedOperationsStayUnsupportedAndFlushStillDelegates()
        {
            using (var inner = new CountingStream(new byte[] { 1 }))
            using (var tee = new OllamaSyntheticWireCapture.TeeStream(inner, "synthetic", new Sink().Write)) {
                Assert.ThrowsException<NotSupportedException>(() => { var ignored = tee.Length; });
                Assert.ThrowsException<NotSupportedException>(() => { var ignored = tee.Position; });
                Assert.ThrowsException<NotSupportedException>(() => tee.Position = 1);
                Assert.ThrowsException<NotSupportedException>(() => tee.Seek(0, SeekOrigin.Begin));
                Assert.ThrowsException<NotSupportedException>(() => tee.SetLength(1));
                Assert.ThrowsException<NotSupportedException>(() => tee.Write(new byte[1], 0, 1));
                tee.Flush(); Assert.AreEqual(1, inner.Flushes);
            }
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public void SyntheticArgumentsRetainObjectVersusScalarWithoutCoercingTheMarker(bool nested)
        {
            using (var fixture = new Fixture()) {
                object marker = nested ? (object)new Dictionary<string, object> { ["value"] = "VB_AI_42" } : "VB_AI_42";
                var parsed = new Dictionary<string, object> { ["marker"] = marker };
                string raw = new JavaScriptSerializer().Serialize(parsed);
                fixture.Scope.RecordArguments("before-unchanged-assertion", raw, parsed);
                var data = fixture.Sink.Object(".json", path => Path.GetFileName(path).StartsWith("arguments-") && !path.EndsWith("-summary.json"));
                Assert.AreEqual(marker.GetType().FullName, data["MarkerType"]);
                Assert.AreSame(marker, parsed["marker"]); Assert.AreEqual(raw, data["RawArgumentJson"]);
                Assert.AreEqual(0, data["NativeToolsExecuted"]);
                if (nested) Assert.IsInstanceOfType<Dictionary<string, object>>(data["MarkerJson"]);
                else Assert.AreEqual("VB_AI_42", data["MarkerJson"]);
            }
        }

        [TestMethod]
        public void DiagnosticSerializationAndDiskErrorsCannotEscapeArgumentOrCompletionRecording()
        {
            using (var fixture = new Fixture(save: (path, bytes) => { throw new IOException("synthetic evidence"); })) {
                var cyclic = new Dictionary<string, object>(); cyclic["self"] = cyclic;
                fixture.Scope.RecordArguments("cycle", cyclic, cyclic);
                fixture.Scope.RecordArguments("disk", "{}", new Dictionary<string, object>());
                fixture.Scope.RecordCompletion("synthetic failure", null, new IOException("secret exception message not logged"));
            }
        }

        [TestMethod]
        public void DeliveryErrorEvidenceIsContentFreeAndDoesNotTreatAnExceptionAsCompletion()
        {
            using(var fixture=new Fixture()) {
                fixture.Scope.RecordDeliveryError("synthetic",new IOException("private-exception-message-not-for-evidence"));
                var data=fixture.Sink.Object("-delivery-error.json");
                Assert.AreEqual(typeof(IOException).FullName,data["ErrorType"]);
                StringAssert.Contains((string)data["Outcome"],"Unknown delivery");
                Assert.IsFalse(fixture.Sink.Text.Contains("private-exception-message-not-for-evidence"));
            }
        }

        [TestMethod]
        public void DefaultDurableWriterCreatesUniqueBoundedSyntheticFilesInItsExplicitRoot()
        {
            string root=Root;
            Func<HttpMessageHandler> current=()=>new HttpClientHandler{AllowAutoRedirect=false};
            try {
                using(var scope=new OllamaSyntheticWireCapture(true,true,root,Endpoint,"synthetic-tool-roundtrip",()=>current,value=>current=value)) {
                    scope.RecordArguments("first","{}",new Dictionary<string,object>());
                    scope.RecordArguments("second","{}",new Dictionary<string,object>());
                }
                var files=Directory.GetFiles(root);
                Assert.AreEqual(4,files.Length);
                Assert.IsTrue(files.All(file=>new FileInfo(file).Length<=OllamaSyntheticWireCapture.Limit));
                Assert.IsTrue(files.All(file=>new JavaScriptSerializer().DeserializeObject(File.ReadAllText(file,Encoding.UTF8))!=null));
            } finally { if(Directory.Exists(root))Directory.Delete(root,true); }
        }

        [DataTestMethod, DataRow(false), DataRow(true)]
        public async Task CompletionEvidenceDoesNotChangeTheParsedResponseOrOriginalFailure(bool fail)
        {
            var handler = new FakeHandler { Error = fail ? new IOException("synthetic original send") : null };
            var provider = LlmProvider.All.Single(item => item.IsOllama);
            var settings = new LlmSettings { OllamaEndpoint = Endpoint.AbsoluteUri }; settings.SetKey(provider, "synthetic-only");
            using (var fixture = new Fixture(save: (path, bytes) => { throw new IOException("synthetic diagnostic"); }))
            using (var client = new LlmChatClient(provider, settings, "synthetic-model", handler)) {
                if (fail) {
                    try { await fixture.Scope.CompleteAsync("synthetic", client, new object[] { new { role="user", content="synthetic" } }, new object[0]); Assert.Fail(); }
                    catch (Exception observed) { Assert.AreSame(handler.Error, observed); }
                } else {
                    var result = await fixture.Scope.CompleteAsync("synthetic", client, new object[] { new { role="user", content="synthetic" } }, new object[0]);
                    Assert.AreEqual("READY", result["content"]);
                }
                Assert.AreEqual(1, handler.Requests);
            }
        }

        private sealed class Sink
        {
            internal readonly Dictionary<string, byte[]> Files = new Dictionary<string, byte[]>();
            internal void Write(string path, byte[] bytes) { Assert.IsFalse(Files.ContainsKey(path)); Files.Add(path, bytes); }
            internal byte[] Single(string suffix) => Files.Single(item => item.Key.EndsWith(suffix)).Value;
            internal string Text => String.Join("\n", Files.Values.Select(Encoding.UTF8.GetString));
            internal IDictionary<string, object> Object(string suffix, Func<string, bool> predicate = null)
                => new JavaScriptSerializer { MaxJsonLength = 4 * OllamaSyntheticWireCapture.Limit }.DeserializeObject(Encoding.UTF8.GetString(
                    Files.Single(item => item.Key.EndsWith(suffix) && (predicate == null || predicate(item.Key))).Value)) as IDictionary<string, object>;
        }
        private sealed class Fixture : IDisposable
        {
            internal readonly Sink Sink = new Sink();
            internal readonly OllamaSyntheticWireCapture Scope;
            internal Fixture(Action<string, byte[]> save = null, Func<HttpContent,string,Action<string,byte[]>,HttpContent> wrap = null)
            {
                Func<HttpMessageHandler> current = () => new HttpClientHandler { AllowAutoRedirect = false };
                Scope = new OllamaSyntheticWireCapture(true, true, Root, Endpoint, "synthetic-tool-roundtrip", () => current,
                    value => current = value, save ?? Sink.Write, wrap);
            }
            public void Dispose() => Scope.Dispose();
        }
        private class CountingStream : MemoryStream
        {
            internal int Reads, Disposals, Flushes;
            internal CountingStream(byte[] bytes) : base(bytes) { }
            public override int Read(byte[] buffer,int offset,int count) { Reads++; return base.Read(buffer,offset,count); }
            public override void Flush() { Flushes++; base.Flush(); }
            protected override void Dispose(bool disposing) { if(disposing) Disposals++; base.Dispose(disposing); }
        }
        private sealed class FailingReadStream : MemoryStream
        {
            private readonly Exception error;
            internal FailingReadStream(Exception error) { this.error = error; }
            public override Task<int> ReadAsync(byte[] buffer,int offset,int count,CancellationToken token) => Task.FromException<int>(error);
        }
        private sealed class PendingReadStream : MemoryStream
        {
            internal readonly TaskCompletionSource<int> Result = new TaskCompletionSource<int>();
            public override Task<int> ReadAsync(byte[] buffer,int offset,int count,CancellationToken token) => Result.Task;
        }
        private sealed class FailingDisposeStream : MemoryStream
        {
            private readonly Exception error;
            internal int Disposals;
            internal FailingDisposeStream(Exception error) { this.error = error; }
            protected override void Dispose(bool disposing) { if(disposing){Disposals++;throw error;}base.Dispose(disposing); }
        }
        private sealed class TrackingTransport : HttpClientHandler
        {
            internal int Disposals;
            protected override void Dispose(bool disposing) { if(disposing)Disposals++;base.Dispose(disposing); }
        }
        private sealed class FakeHandler : HttpMessageHandler
        {
            internal int Requests, Disposals;
            internal Exception Error;
            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request,CancellationToken token)
            {
                Requests++;
                return Error != null ? Task.FromException<HttpResponseMessage>(Error) : Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
                    Content = new StringContent("{\"choices\":[{\"message\":{\"role\":\"assistant\",\"content\":\"READY\"}}]}",Encoding.UTF8,"application/json") });
            }
            protected override void Dispose(bool disposing) { if(disposing)Disposals++;base.Dispose(disposing); }
        }
    }
}
