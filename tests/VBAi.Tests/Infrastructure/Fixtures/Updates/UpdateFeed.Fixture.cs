using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
namespace VBAi.Tests.Unit
{
    internal sealed class OwnedUpdateContent : HttpContent
    {
        private readonly byte[] bytes;
        internal OwnedUpdateContent(string text) { bytes = System.Text.Encoding.UTF8.GetBytes(text); }
        protected override bool TryComputeLength(out long length) { length = 0; return false; }
        protected override Task SerializeToStreamAsync(Stream stream, TransportContext context) => stream.WriteAsync(bytes, 0, bytes.Length);
        protected override Task<Stream> CreateContentReadStreamAsync() => Task.FromResult<Stream>(new MemoryStream(bytes, false));
    }
    internal sealed class OwnedUpdateProgress : IProgress<int>
    {
        private readonly Action<int> report;
        internal OwnedUpdateProgress(Action<int> report) { this.report = report; }
        public void Report(int value) => report(value);
    }
}