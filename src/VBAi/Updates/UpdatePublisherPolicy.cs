using System;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

namespace VBAi
{

    /// <summary>Release-owned certificate pins. An update feed or pending job cannot grant trust.</summary>
    internal static class UpdatePublisherPolicy
    {
        // Intentionally empty until the maintainer supplies and approves a signing certificate.
        // Certificate rotation must ship an explicit overlap policy in a trusted product build.
        /// <summary>Trusted product signer-certificate SHA-256 pins shipped in this build; currently empty, so product installers are refused.</summary>
        private static readonly string[] CertificateSha256 = new string[0];

        /// <summary>Checks the signed file's raw certificate digest against the build-owned publisher pins.</summary>
        /// <param name="path">Candidate installer file from which the signer certificate is extracted.</param>
        /// <returns>True for a pinned signer; false when no pin is configured, no pin matches, or certificate extraction fails cryptographically.</returns>
        /// <remarks>This is publisher identity checking, not Windows trust verification; the installer runner requires both.</remarks>
        internal static bool Accepts(string path)
        {
            if (CertificateSha256.Length == 0) return false;
            try
            {
                using (var certificate = X509Certificate.CreateFromSignedFile(path))
                using (var hash = SHA256.Create())
                {
                    string actual = BitConverter.ToString(hash.ComputeHash(certificate.GetRawCertData())).Replace("-", "");
                    return Array.Exists(CertificateSha256, expected => string.Equals(expected, actual, StringComparison.OrdinalIgnoreCase));
                }
            }
            catch (CryptographicException) { return false; }
        }
    }
}
