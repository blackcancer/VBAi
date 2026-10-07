namespace VBAi.Tests.Unit
{
    using System.Security.Cryptography.X509Certificates;
    using VBAi;
    internal sealed class CertificateDispatchStore : VbeSession.ISigningStore
    {
        public X509Certificate2Collection Certificates { get; } = new X509Certificate2Collection();
        internal bool Disposed;
        internal OpenFlags? Opened;
        public void Open(OpenFlags flags) { Opened = flags; }
        public void Dispose() { Disposed = true; }
    }
}
