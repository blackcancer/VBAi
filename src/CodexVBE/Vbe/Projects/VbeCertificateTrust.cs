using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;

namespace CodexVBE
{
    /// <summary>Évalue la chaîne d’un certificat de signature sans récupération réseau ni accès à sa clé privée.</summary>
    internal static class VbeCertificateTrust
    {
        internal delegate bool ChainReader(IntPtr engine, IntPtr certificate, IntPtr time, IntPtr store, ref ChainParameters parameters, uint flags, IntPtr reserved, out IntPtr chain);
        internal delegate bool PolicyReader(IntPtr policy, IntPtr chain, ref PolicyParameters parameters, ref PolicyStatus status);
        /// <summary>Native chain, policy and allocation boundaries; defaults retain offline Windows validation.</summary>
        internal static ChainReader ReadChain = CertGetCertificateChain;
        internal static PolicyReader ReadPolicy = CertVerifyCertificateChainPolicy;
        internal static Action<IntPtr> ReleaseChain = CertFreeCertificateChain;
        internal static Func<int, IntPtr> AllocatePointers = Marshal.AllocHGlobal;
        /// <summary>Construit une chaîne Windows locale et applique la politique Authenticode sans ignorer les erreurs.</summary>
        /// <param name="certificate">Certificat public à examiner.</param>
        /// <returns>Diagnostic daté ; il ne vérifie aucune signature VBA ni aucun digest de document.</returns>
        internal static object Evaluate(X509Certificate2 certificate)
        {
            if (certificate == null) throw new ArgumentNullException(nameof(certificate));
            IntPtr oid = Marshal.StringToHGlobalAnsi("1.3.6.1.5.5.7.3.3"), pointers = IntPtr.Zero, chain = IntPtr.Zero;
            try
            {
                pointers = AllocatePointers(IntPtr.Size); Marshal.WriteIntPtr(pointers, oid);
                var parameters = new ChainParameters { Size = Marshal.SizeOf(typeof(ChainParameters)),
                    Usage = new UsageMatch { Usage = new EnhancedUsage { Count = 1, Identifiers = pointers } } };
                // Cache-only chain AND revocation, no automatic root update, excluding root from revocation.
                const uint flags = 0x00000004u | 0x00000100u | 0x40000000u | 0x80000000u;
                if (!ReadChain(IntPtr.Zero, certificate.Handle, IntPtr.Zero, IntPtr.Zero,
                    ref parameters, flags, IntPtr.Zero, out chain)) throw new Win32Exception(Marshal.GetLastWin32Error());
                var context = (ChainHeader)Marshal.PtrToStructure(chain, typeof(ChainHeader));
                var policy = new PolicyParameters { Size = Marshal.SizeOf(typeof(PolicyParameters)) };
                var status = new PolicyStatus { Size = Marshal.SizeOf(typeof(PolicyStatus)) };
                if (!ReadPolicy(new IntPtr(2), chain, ref policy, ref status))
                    throw new Win32Exception(Marshal.GetLastWin32Error());
                uint errors = context.ErrorStatus;
                const uint revocationUnknown = 0x40u | 0x01000000u;
                string state = errors == 0 && status.Error == 0 ? "Trusted" :
                    (errors & ~revocationUnknown) == 0 && (errors & revocationUnknown) != 0 &&
                    (status.Error == 0 || status.Error == 0x80092013u || status.Error == 0x800B010Eu) ? "IndeterminateOffline" : "NotTrusted";
                return new { certificate.Thumbprint, certificate.Subject, certificate.Issuer,
                    CheckedAtUtc = DateTime.UtcNow.ToString("o"), State = state,
                    Trusted = state == "Trusted", ChainErrorStatus = "0x" + errors.ToString("X8"),
                    PolicyError = "0x" + status.Error.ToString("X8"), status.ChainIndex, status.ElementIndex,
                    OfflineOnly = true, PrivateKeyAccessed = false, MacroSignatureVerified = false,
                    Limit = "Current Windows trust and cached revocation only; no VBA digest, signer binding or timestamp verification. Missing cached evidence is not trust." };
            }
            finally
            {
                if (chain != IntPtr.Zero) ReleaseChain(chain);
                if (pointers != IntPtr.Zero) Marshal.FreeHGlobal(pointers);
                Marshal.FreeHGlobal(oid);
            }
        }
        [StructLayout(LayoutKind.Sequential)] internal struct EnhancedUsage { internal uint Count; internal IntPtr Identifiers; }
        [StructLayout(LayoutKind.Sequential)] internal struct UsageMatch { internal uint Type; internal EnhancedUsage Usage; }
        [StructLayout(LayoutKind.Sequential)] internal struct ChainParameters
        {
            internal int Size; internal UsageMatch Usage; internal UsageMatch IssuancePolicy;
            internal uint UrlTimeout; internal int CheckFreshness; internal uint Freshness;
            internal IntPtr CacheResync; internal IntPtr StrongSignature; internal uint StrongFlags;
        }
        [StructLayout(LayoutKind.Sequential)] internal struct ChainHeader { internal int Size; internal uint ErrorStatus; internal uint InfoStatus; }
        [StructLayout(LayoutKind.Sequential)] internal struct PolicyParameters { internal int Size; internal uint Flags; internal IntPtr Extra; }
        [StructLayout(LayoutKind.Sequential)] internal struct PolicyStatus
        { internal int Size; internal uint Error; internal int ChainIndex; internal int ElementIndex; internal IntPtr Extra; }
        [DllImport("crypt32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool CertGetCertificateChain(IntPtr engine, IntPtr certificate, IntPtr time, IntPtr store,
            ref ChainParameters parameters, uint flags, IntPtr reserved, out IntPtr chain);
        [DllImport("crypt32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool CertVerifyCertificateChainPolicy(IntPtr policy, IntPtr chain, ref PolicyParameters parameters, ref PolicyStatus status);
        [DllImport("crypt32.dll")] private static extern void CertFreeCertificateChain(IntPtr chain);
    }
}
