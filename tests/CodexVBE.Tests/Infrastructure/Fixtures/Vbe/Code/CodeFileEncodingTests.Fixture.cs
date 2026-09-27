namespace CodexVBE.Tests.Unit
{
    using System;
    using System.IO;
    using System.Security.Cryptography;
    using System.Text;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class CodeFileEncodingTests
    {
        private static string Hash(string value)
        {
            using (var sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(value))).Replace("-", "").ToLowerInvariant();
        }

        private static byte[] Combine(byte[] prefix, byte[] body)
        {
            var result = new byte[prefix.Length + body.Length];
            Buffer.BlockCopy(prefix, 0, result, 0, prefix.Length);
            Buffer.BlockCopy(body, 0, result, prefix.Length, body.Length);
            return result;
        }
    }
}
