using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;
using VBAi;

namespace VBAi.Tests.Integration
{
    internal sealed partial class ExcelVbeFixture
    {
        /// <summary>Retains the prepared baseline and one post-import capture before any exact comparison or native readback.</summary>
        internal static void RetainEmbeddedImportRawEvidence(string ownedRoot,
            IDictionary<string, byte[]> target, IDictionary<string, byte[]> actual, Action<object> evidence)
        {
            if (string.IsNullOrWhiteSpace(ownedRoot) || !Path.IsPathRooted(ownedRoot) ||
                !string.Equals(Path.GetPathRoot(ownedRoot), Path.GetPathRoot(Path.GetFullPath(ownedRoot)),
                    StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("An absolute owned evidence root is required.", nameof(ownedRoot));
            if (target == null) throw new ArgumentNullException(nameof(target));
            if (actual == null) throw new ArgumentNullException(nameof(actual));
            if (evidence == null) throw new ArgumentNullException(nameof(evidence));
            string root = Path.Combine(ownedRoot, "post-import-raw");
            if (Directory.Exists(root) || File.Exists(root)) throw new IOException("Post-import raw evidence already exists; no overwrite.");
            Directory.CreateDirectory(root);
            var targetRows = RetainEmbeddedRawSide(root, "target", target);
            var actualRows = RetainEmbeddedRawSide(root, "actual", actual);
            var receipt = new { Phase = "PostImportRawSnapshotsRetained", Root = root,
                TargetRole = "PreparedBaseline", ActualRole = "ActualPostImportCapture",
                Target = targetRows, Actual = actualRows, CaptureCount = 1, ImportReplay = false };
            File.WriteAllText(Path.Combine(root, "raw-evidence.json"),
                new JavaScriptSerializer().Serialize(receipt), new UTF8Encoding(false));
            evidence(receipt);
        }

        private static object[] RetainEmbeddedRawSide(string root, string side, IDictionary<string, byte[]> files)
        {
            string directory = Path.Combine(root, side);
            Directory.CreateDirectory(directory);
            var rows = new List<object>();
            foreach (var pair in files.OrderBy(item => item.Key, StringComparer.Ordinal))
            {
                if (string.IsNullOrWhiteSpace(pair.Key) || pair.Key == "." || pair.Key == ".." ||
                    pair.Key != Path.GetFileName(pair.Key) || pair.Key.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 ||
                    pair.Value == null)
                    throw new InvalidOperationException("The raw snapshot contains an unsafe or absent file.");
                string path = Path.Combine(directory, pair.Key);
                using (var output = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.Read))
                    output.Write(pair.Value, 0, pair.Value.Length);
                string hash;
                using (var sha = SHA256.Create())
                    hash = BitConverter.ToString(sha.ComputeHash(pair.Value)).Replace("-", "");
                bool binary = pair.Key.EndsWith(".frx", StringComparison.Ordinal);
                rows.Add(new { Name = pair.Key, Path = path, Bytes = pair.Value.Length, Sha256 = hash,
                    Encoding = binary ? "binary" : "UTF-8",
                    OleOffsets = pair.Key.EndsWith(".frm", StringComparison.Ordinal)
                        ? EmbeddedGitSnapshotOracle.Offsets(VbaGitSnapshot.Utf8.GetString(pair.Value)) : new int[0] });
            }
            return rows.ToArray();
        }

        /// <summary>Reads native state once even if the exact snapshot assertion fails, preserving both errors.</summary>
        internal static void VerifyEmbeddedImportReadbacks(Action assertExactSnapshot, Action readAndAssertNative)
        {
            if (assertExactSnapshot == null) throw new ArgumentNullException(nameof(assertExactSnapshot));
            if (readAndAssertNative == null) throw new ArgumentNullException(nameof(readAndAssertNative));
            Exception primary = null, readback = null;
            try { assertExactSnapshot(); } catch (Exception error) { primary = error; }
            try { readAndAssertNative(); } catch (Exception error) { readback = error; }
            if (primary != null && readback != null)
                throw new AggregateException("Exact imported snapshot and independent native readback both failed; neither action was replayed.", primary, readback);
            if (primary != null) ExceptionDispatchInfo.Capture(primary).Throw();
            if (readback != null) ExceptionDispatchInfo.Capture(readback).Throw();
        }
    }
}
