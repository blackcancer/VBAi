namespace VBAi.Tests.Unit
{
    using System;
    using System.Security.Cryptography;
    using System.Text;

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
