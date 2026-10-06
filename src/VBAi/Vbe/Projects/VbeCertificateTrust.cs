using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;

namespace VBAi
{

    /// <summary>Évalue la chaîne d’un certificat de signature sans récupération réseau ni accès à sa clé privée.</summary>
    internal static class VbeCertificateTrust
    {

        /// <summary>Signature injectable de la construction de chaîne Crypt32.</summary>
        /// <param name="engine">Moteur de chaîne facultatif.</param>
        /// <param name="certificate">Contexte du certificat public à vérifier.</param>
        /// <param name="time">Instant de validation facultatif.</param>
        /// <param name="store">Magasin supplémentaire facultatif.</param>
        /// <param name="parameters">Options de construction de chaîne.</param>
        /// <param name="flags">Drapeaux Crypt32 de validation.</param>
        /// <param name="reserved">Paramètre réservé par l’API.</param>
        /// <param name="chain">Reçoit la chaîne construite.</param>
        /// <returns><see langword="true"/> si Crypt32 a construit une chaîne.</returns>
        internal delegate bool ChainReader(IntPtr engine, IntPtr certificate, IntPtr time, IntPtr store, ref ChainParameters parameters, uint flags, IntPtr reserved, out IntPtr chain);

        /// <summary>Signature injectable de la vérification Authenticode d’une chaîne.</summary>
        /// <param name="policy">Identifiant de la politique Crypt32.</param>
        /// <param name="chain">Chaîne à vérifier.</param>
        /// <param name="parameters">Options de vérification de politique.</param>
        /// <param name="status">Reçoit le résultat et les indices d’erreur.</param>
        /// <returns><see langword="true"/> si l’appel de vérification a abouti.</returns>
        internal delegate bool PolicyReader(IntPtr policy, IntPtr chain, ref PolicyParameters parameters, ref PolicyStatus status);

        /// <summary>Native chain, policy and allocation boundaries; defaults retain offline Windows validation.</summary>
        internal static ChainReader ReadChain = CertGetCertificateChain;

        /// <summary>Point d’injection de la vérification de politique Authenticode.</summary>
        internal static PolicyReader ReadPolicy = CertVerifyCertificateChainPolicy;

        /// <summary>Libère la chaîne Crypt32 créée par le lecteur courant.</summary>
        internal static Action<IntPtr> ReleaseChain = CertFreeCertificateChain;

        /// <summary>Alloue le tampon natif des identifiants d’usage étendu.</summary>
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

        /// <summary>Structure CRYPT_OID_INFO contenant la liste d’identifiants d’usage.</summary>
        [StructLayout(LayoutKind.Sequential)] internal struct EnhancedUsage {

/// <summary>Nombre d’identifiants OID présents.</summary>
internal uint Count;

/// <summary>Pointeur vers le tableau natif d’OID.</summary>
internal IntPtr Identifiers; }

        /// <summary>Association entre un mode de correspondance et des usages étendus.</summary>
        [StructLayout(LayoutKind.Sequential)] internal struct UsageMatch {

/// <summary>Mode de correspondance défini par Crypt32.</summary>
internal uint Type;

/// <summary>Liste des usages recherchés.</summary>
internal EnhancedUsage Usage; }

        /// <summary>Paramètres CERT_CHAIN_PARA utilisés pour construire une chaîne de certificat.</summary>
        [StructLayout(LayoutKind.Sequential)] internal struct ChainParameters
        {

            /// <summary>Taille de cette structure native en octets.</summary>
            internal int Size;

/// <summary>Usages étendus exigés pour la chaîne.</summary>
internal UsageMatch Usage;

/// <summary>Politique d’émission exigée, laissée vide pour l’évaluation Authenticode.</summary>
internal UsageMatch IssuancePolicy;

            /// <summary>Délai maximal accordé aux requêtes d’URL, nul en validation hors ligne.</summary>
            internal uint UrlTimeout;

/// <summary>Indique si la fraîcheur de la chaîne doit être vérifiée.</summary>
internal int CheckFreshness;

/// <summary>Durée de fraîcheur demandée lorsque ce contrôle est activé.</summary>
internal uint Freshness;

            /// <summary>Instant de resynchronisation du cache, nul pour ce contrôle hors ligne.</summary>
            internal IntPtr CacheResync;

/// <summary>Paramètres de vérification de signature forte facultatifs.</summary>
internal IntPtr StrongSignature;

/// <summary>Drapeaux complémentaires de signature forte.</summary>
internal uint StrongFlags;
        }

        /// <summary>En-tête CERT_CHAIN_CONTEXT lu pour obtenir les indicateurs d’erreur.</summary>
        [StructLayout(LayoutKind.Sequential)] internal struct ChainHeader {

/// <summary>Taille de la structure native.</summary>
internal int Size;

/// <summary>Indicateurs d’erreur de la chaîne.</summary>
internal uint ErrorStatus;

/// <summary>Indicateurs informatifs de la chaîne.</summary>
internal uint InfoStatus; }

        /// <summary>Paramètres utilisés par CertVerifyCertificateChainPolicy.</summary>
        [StructLayout(LayoutKind.Sequential)] internal struct PolicyParameters {

/// <summary>Taille de la structure native.</summary>
internal int Size;

/// <summary>Drapeaux propres à la politique.</summary>
internal uint Flags;

/// <summary>Données supplémentaires facultatives.</summary>
internal IntPtr Extra; }

        /// <summary>Résultat POLICY_STATUS comprenant l’erreur et son emplacement dans la chaîne.</summary>
        [StructLayout(LayoutKind.Sequential)] internal struct PolicyStatus
        {

/// <summary>Taille de la structure native.</summary>
internal int Size;

/// <summary>Code d’erreur de politique.</summary>
internal uint Error;

/// <summary>Index du maillon fautif, s’il est fourni.</summary>
internal int ChainIndex;

/// <summary>Index de l’élément fautif dans le maillon.</summary>
internal int ElementIndex;

/// <summary>Données supplémentaires retournées par Crypt32.</summary>
internal IntPtr Extra; }

        /// <summary>Construit la chaîne locale d’un certificat selon les paramètres fournis.</summary>
        /// <param name="engine">Moteur de chaîne facultatif.</param>
        /// <param name="certificate">Contexte du certificat.</param>
        /// <param name="time">Instant de vérification facultatif.</param>
        /// <param name="store">Magasin supplémentaire facultatif.</param>
        /// <param name="parameters">Paramètres de construction de chaîne.</param>
        /// <param name="flags">Options de validation Crypt32.</param>
        /// <param name="reserved">Paramètre réservé.</param>
        /// <param name="chain">Reçoit le contexte de chaîne.</param>
        /// <returns><see langword="true"/> si la chaîne a été construite.</returns>
        [DllImport("crypt32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool CertGetCertificateChain(IntPtr engine, IntPtr certificate, IntPtr time, IntPtr store,
            ref ChainParameters parameters, uint flags, IntPtr reserved, out IntPtr chain);

        /// <summary>Évalue la chaîne avec la politique de certificat Crypt32 demandée.</summary>
        /// <param name="policy">Identifiant de politique.</param>
        /// <param name="chain">Contexte de chaîne construit.</param>
        /// <param name="parameters">Paramètres de politique.</param>
        /// <param name="status">Reçoit le résultat de la vérification.</param>
        /// <returns><see langword="true"/> si l’appel de politique a abouti.</returns>
        [DllImport("crypt32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool CertVerifyCertificateChainPolicy(IntPtr policy, IntPtr chain, ref PolicyParameters parameters, ref PolicyStatus status);

        /// <summary>Libère le contexte de chaîne alloué par Crypt32.</summary>
        /// <param name="chain">Contexte de chaîne à détruire.</param>
        [DllImport("crypt32.dll")] private static extern void CertFreeCertificateChain(IntPtr chain);
    }
}
