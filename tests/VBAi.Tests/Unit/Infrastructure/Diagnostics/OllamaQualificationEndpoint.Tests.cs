using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class OllamaQualificationEndpointTests
    {
        [TestMethod]
        public void MissingOverrideRetainsTheHistoricalEndpoint()
            => Assert.AreEqual(OllamaQualificationEndpoint.Default, OllamaQualificationEndpoint.Parse(null).AbsoluteUri);

        [DataTestMethod, DataRow(11434), DataRow(52541), DataRow(65535)]
        public void ExplicitLocalPortPreservesChatAndCatalogueCaptureOnly(int port)
        {
            var selected = OllamaQualificationEndpoint.Parse("http://127.0.0.1:" + port + "/v1/chat/completions");
            Assert.AreEqual(port, selected.Port);
            OllamaQualificationEndpoint.RequireWireUri(selected, selected);
            OllamaQualificationEndpoint.RequireWireUri(new Uri("http://127.0.0.1:" + port + "/api/tags"), selected);
        }

        [DataTestMethod, DataRow(""), DataRow("http://localhost:11434/v1/chat/completions"),
            DataRow("https://127.0.0.1:11434/v1/chat/completions"), DataRow("http://[::1]:11434/v1/chat/completions"),
            DataRow("http://fixture.invalid:11434/v1/chat/completions"), DataRow("http://127.0.0.1:11434/api/chat"),
            DataRow("http://name:secret@127.0.0.1:11434/v1/chat/completions"),
            DataRow("http://127.0.0.1:11434/v1/chat/completions?token=synthetic"),
            DataRow("http://127.0.0.1:11434/v1/chat/completions#synthetic"),
            DataRow(" http://127.0.0.1:11434/v1/chat/completions"), DataRow("http://127.0.0.1:0/v1/chat/completions"),
            DataRow("http://127.0.0.1:65536/v1/chat/completions")]
        public void EndpointRefusesRemoteAmbiguousCredentialOrDifferentRoute(string configured)
            => Assert.ThrowsException<InvalidOperationException>(() => OllamaQualificationEndpoint.Parse(configured));

        [DataTestMethod, DataRow("http://127.0.0.1:11434/api/tags"), DataRow("http://localhost:52541/api/tags"),
            DataRow("http://127.0.0.1:52541/api/pull"), DataRow("http://127.0.0.1:52541/api/tags?synthetic=1"),
            DataRow("http://127.0.0.1:52541/api/tags#synthetic"), DataRow("http://name:secret@127.0.0.1:52541/api/tags"),
            DataRow("https://127.0.0.1:52541/api/tags")]
        public void WireCaptureRefusesAnotherPortHostRouteOrCredentials(string raw)
            => Assert.ThrowsException<InvalidOperationException>(() => OllamaQualificationEndpoint.RequireWireUri(new Uri(raw),
                OllamaQualificationEndpoint.Parse("http://127.0.0.1:52541/v1/chat/completions")));

        [TestMethod]
        public void MissingOrRelativeWireUriCannotCrossTheCaptureBoundary()
        {
            var selected = OllamaQualificationEndpoint.Parse(null);
            Assert.ThrowsException<InvalidOperationException>(() => OllamaQualificationEndpoint.RequireWireUri(null, selected));
            Assert.ThrowsException<InvalidOperationException>(() => OllamaQualificationEndpoint.RequireWireUri(selected, null));
            Assert.ThrowsException<InvalidOperationException>(() => OllamaQualificationEndpoint.RequireWireUri(new Uri("api/tags", UriKind.Relative), selected));
        }
    }
}
