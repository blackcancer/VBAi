using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Security.Cryptography;

namespace VBAi.Tests.Integration
{
    internal sealed partial class OfficeVbeFixture
    {
        /// <summary>Inspects a bounded immutable snapshot of a frozen owned package while Word holds its write handle.</summary>
        internal static void InspectWordControlPackage(string path, string expectedSha256, Action<Package> inspect)
        {
            if (string.IsNullOrWhiteSpace(path) || !Path.IsPathRooted(path) || inspect == null ||
                expectedSha256 == null || expectedSha256.Length != 64 || expectedSha256.Any(c => !Uri.IsHexDigit(c)))
                throw new ArgumentException("An absolute owned package, frozen SHA-256 and inspector are required.");
            const int maximum = 2 * 1024 * 1024;
            using (var snapshot = new MemoryStream())
            {
                using (var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    if (file.Length == 0 || file.Length > maximum) throw new IOException("The generated control package is outside its size bound.");
                    var buffer = new byte[4096]; int read;
                    while ((read = file.Read(buffer, 0, buffer.Length)) != 0)
                    {
                        if (snapshot.Length + read > maximum) throw new IOException("The control package grew outside its size bound.");
                        snapshot.Write(buffer, 0, read);
                    }
                }
                snapshot.Position = 0;
                string actual;
                using (var sha = SHA256.Create()) actual = BitConverter.ToString(sha.ComputeHash(snapshot)).Replace("-", "");
                if (!string.Equals(actual, expectedSha256, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("The owned package changed before structural inspection.");
                snapshot.Position = 0;
                using (var package = Package.Open(snapshot, FileMode.Open, FileAccess.Read)) inspect(package);
            }
        }
    }
}
