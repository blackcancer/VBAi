using System;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    /// <summary>Compares independent owner-bridge exports with the installed Git checkpoint, retaining raw provenance.</summary>
    internal static class EmbeddedGitSnapshotOracle
    {
        internal static void Verify(VbaGitSnapshot baseline, VbaGitSnapshot checkpoint)
        {
            Assert.IsNotNull(baseline); Assert.IsNotNull(checkpoint);
            var expected = baseline.ComparisonFiles(); var actual = checkpoint.ComparisonFiles();
            CollectionAssert.AreEquivalent(expected.Keys.ToArray(), actual.Keys.ToArray());
            foreach (var pair in expected)
                CollectionAssert.AreEqual(pair.Value, actual[pair.Key], "Owner-bridge baseline differs from checkpoint: " + pair.Key);
        }

        internal static object Describe(VbaGitSnapshot snapshot)
        {
            var compared = snapshot.ComparisonFiles();
            return snapshot.Files.Select(pair => new { Path = pair.Key, Bytes = pair.Value.Length,
                RawSha256 = Hash(pair.Value), ComparisonSha256 = Hash(compared[pair.Key]),
                OleOffsets = pair.Key.EndsWith(".frm", StringComparison.Ordinal) ? Offsets(VbaGitSnapshot.Utf8.GetString(pair.Value)) : new int[0] }).ToArray();
        }

        internal static int[] Offsets(string text)
        {
            int code = Regex.Match(text, "^Attribute VB_Name = ", RegexOptions.Multiline).Index;
            return Regex.Matches(text.Substring(0, code), "^[ \\t]*OleObjectBlob[ \\t]*=[ \\t]*\"[^\"\\r\\n]+\"[ \\t]*:[ \\t]*([0-9a-f]+)[ \\t]*$", RegexOptions.Multiline | RegexOptions.IgnoreCase)
                .Cast<Match>().Select(match => checked((int)uint.Parse(match.Groups[1].Value, NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture)))
                .Distinct().OrderBy(x => x).ToArray();
        }

        private static string Hash(byte[] bytes) { using (var hash = SHA256.Create()) return BitConverter.ToString(hash.ComputeHash(bytes)).Replace("-", ""); }
    }
}
