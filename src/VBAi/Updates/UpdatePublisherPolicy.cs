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
        private static readonly string[] CertificateSha256 = new string[0];

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
