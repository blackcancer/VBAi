using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;

namespace VBAi.Tests.Integration
{
    /// <summary>Test-only attestation and one fresh copy of the retained inert LabelButton workbook.</summary>
    internal static class RetainedRootFontWorkbook
    {
        internal const string SourceSha256 = "B5264F941E0FD398A9DE03B203DB7A31A6DB9FF939B894E7829A0203C76F4B73";
        internal const string SourceMarker = "EMBEDDED_67b6b3d058ac43b591316d7a00f584b4";
        private const string SourceSuffix = "\\artifacts\\root-font-owner-observation-20261002\\native-labelbutton-diagnostic\\hosts\\29a84356e87e4f03a08d7025b20dd947\\EmbeddedGit.xlsm";
        private static readonly IDictionary<string, string> ComparisonHashes = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["EmbeddedClass.cls"] = "2A19380F277D1E6E082F3EABC69D2E61F0E56158C9DE4200DDAD80BB8C12787A",
            ["EmbeddedForm.frm"] = "EF8E1C42D1B299412783590210B2154CFEA35F62DD7750DA05A6B01EB1C4AD68",
            ["EmbeddedForm.frx"] = "72568A6AEE7D62C2EED3B2BF32F83887895B098FD20279440902387F891B4FDE",
            ["EmbeddedModule.bas"] = "083A36E1609F2E60A46FD6F1592D10C8033F37B30D3827A230C9D2016AFD0DF7",
            ["Feuil1.vba"] = "E3B0C44298FC1C149AFBF4C8996FB92427AE41E4649B934CA495991B7852B855",
            ["ThisWorkbook.vba"] = "E3B0C44298FC1C149AFBF4C8996FB92427AE41E4649B934CA495991B7852B855"
        };

        internal static byte[] Tahoma825Descriptor()
        { return new byte[] { 1, 0, 0, 0, 144, 1, 68, 66, 1, 0, 6, 84, 97, 104, 111, 109, 97 }; }

        internal static string BaselineMarker(string profile, string correlationNonce)
        { return profile == RootFontObservationManifest.RetainedSyntheticTahoma825 ? SourceMarker : correlationNonce; }

        internal static T AttestFirstCapture<T>(Func<T> capture, Func<T, byte[]> verify, Action<T, byte[]> record)
        {
            T snapshot = capture();
            byte[] descriptor = verify(snapshot);
            record(snapshot, descriptor);
            return snapshot;
        }

        internal static string RequirePinnedSource(string path, Func<string, FileAttributes> metadata = null)
        {
            if (string.IsNullOrWhiteSpace(path))
                throw new InvalidOperationException("The retained workbook source must be explicitly declared.");
            path = RequireExactLocal(path);
            if (!path.EndsWith(SourceSuffix, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The retained workbook is not the pinned owned synthetic artifact.");
            RequireSafePath(path, true, metadata ?? File.GetAttributes);
            RequireHash(File.ReadAllBytes(path), SourceSha256);
            return path;
        }

        /// <summary>Revalidates bytes and destination, then creates one copy without overwriting anything.</summary>
        internal static void CopyCreateNew(string source, string destination, string fixtureRoot,
            string expectedSha256, Func<string, FileAttributes> metadata = null)
        {
            source = RequireExactLocal(source); destination = RequireExactLocal(destination);
            fixtureRoot = RequireExactLocal(fixtureRoot);
            if (string.Equals(source, destination, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(Path.GetDirectoryName(destination), fixtureRoot, StringComparison.OrdinalIgnoreCase) ||
                Path.GetFileName(destination) != "EmbeddedGit.xlsm")
                throw new InvalidOperationException("The retained workbook destination is outside the fresh fixture root.");
            var inspect = metadata ?? File.GetAttributes;
            RequireSafePath(source, true, inspect);
            RequireSafePath(fixtureRoot, true, inspect);
            RequireSafePath(destination, false, inspect);
            if ((inspect(source) & FileAttributes.Directory) != 0 ||
                (inspect(fixtureRoot) & FileAttributes.Directory) == 0)
                throw new InvalidOperationException("The source must be a file and the fixture root a directory.");
            byte[] bytes = File.ReadAllBytes(source);
            RequireHash(bytes, expectedSha256);
            using (var copy = new FileStream(destination, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            { copy.Write(bytes, 0, bytes.Length); copy.Flush(true); }
            RequireHash(File.ReadAllBytes(destination), expectedSha256);
        }

        /// <summary>Refuses a source/code/resource drift before any scalar font getter or menu action.</summary>
        internal static byte[] RequirePinnedBaseline(VbaGitSnapshot baseline)
        {
            if (baseline?.Manifest?.Components == null || baseline.Files == null ||
                baseline.Manifest.Components.Length != 5 ||
                baseline.Manifest.Components.Count(item => item.Name == "EmbeddedForm" && item.Type == 3 && item.HasResources) != 1 ||
                baseline.Manifest.Components.Count(item => item.Name == "EmbeddedModule" && item.Type == 1) != 1 ||
                baseline.Manifest.Components.Count(item => item.Name == "EmbeddedClass" && item.Type == 2) != 1 ||
                baseline.Manifest.Components.Count(item => item.Type == 100) != 2)
                throw new InvalidOperationException("The copied workbook component identity differs from its pinned synthetic source.");
            var compared = baseline.ComparisonFiles();
            if (!baseline.Files.TryGetValue("EmbeddedForm.frx", out byte[] resource) || resource.Length != 3096)
                throw new InvalidOperationException("The retained form must contain its exact 3096-byte resource.");
            foreach (var item in ComparisonHashes)
                if (!compared.TryGetValue(item.Key, out byte[] bytes) ||
                    !string.Equals(Sha(bytes), item.Value, StringComparison.Ordinal))
                    throw new InvalidOperationException("The copied workbook source or resource differs from the retained synthetic baseline: " + item.Key);
            var form = baseline.Manifest.Components.Single(item => item.Name == "EmbeddedForm");
            return RootFontObservationManifest.RequireRoot(baseline.FormFonts(form),
                "AfterInitialCapture", RootFontObservationManifest.RetainedSyntheticTahoma825);
        }

        private static string RequireExactLocal(string path)
        {
            string exact = ExcelOwnedBootstrapPlan.RequireLocalAbsolutePath(path);
            if (!string.Equals(path, exact, StringComparison.Ordinal) || path.IndexOf('/', 2) >= 0)
                throw new InvalidOperationException("The retained workbook path must be exact and local.");
            return path;
        }

        private static void RequireSafePath(string path, bool exists, Func<string, FileAttributes> metadata)
        {
            FileAttributes? actual;
            try { actual = metadata(path); }
            catch (FileNotFoundException) { actual = null; }
            catch (DirectoryNotFoundException) { actual = null; }
            if (actual.HasValue != exists || actual.HasValue && (actual.Value & FileAttributes.ReparsePoint) != 0)
                throw new InvalidOperationException("The retained workbook path is missing, occupied or a reparse point.");
            for (string parent = Path.GetDirectoryName(path); parent != null; parent = Path.GetDirectoryName(parent))
                if ((metadata(parent) & (FileAttributes.Directory | FileAttributes.ReparsePoint)) != FileAttributes.Directory)
                    throw new InvalidOperationException("The retained workbook path traverses a non-directory or reparse point.");
        }

        private static void RequireHash(byte[] bytes, string expected)
        {
            if (!string.Equals(Sha(bytes), expected, StringComparison.Ordinal))
                throw new InvalidOperationException("The retained workbook bytes differ from the pinned synthetic source.");
        }

        private static string Sha(byte[] bytes)
        {
            using (var hash = SHA256.Create())
                return BitConverter.ToString(hash.ComputeHash(bytes)).Replace("-", "");
        }
    }
}
