using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;

namespace VBAi.Tests.Integration
{
    /// <summary>Publishes one explicitly requested diagnostic claim from an already captured owned baseline.</summary>
    internal static class RootFontObservationManifest
    {
        internal const string OptIn = "VBAi_RUN_ROOT_FONT_OBSERVATION_TESTS";
        internal const string ManifestVariable = "VBAi_TEST_ROOT_FONT_OBSERVATION_MANIFEST";
        internal const string ModeVariable = "VBAi_TEST_ROOT_FONT_OBSERVATION_MODE";
        internal const string SeedProfileVariable = "VBAi_TEST_ROOT_FONT_OBSERVATION_SEED_PROFILE";
        internal const string SourceWorkbookVariable = "VBAi_TEST_ROOT_FONT_OBSERVATION_SOURCE_WORKBOOK";
        internal const string SyntheticExplicitArial9 = "SyntheticExplicitArial9";
        internal const string RetainedSyntheticTahoma825 = "RetainedSyntheticTahoma825";
        private static readonly string[] Layouts = { "LabelButton", "TextBox", "ComboBox", "ListBox", "CheckBox", "OptionButton",
            "ToggleButton", "ScrollBar", "SpinButton", "TabStrip", "Image", "FrameMultiPage" };

        internal sealed class Configuration
        {
            internal readonly string Path, Mode, SeedProfile, SourceWorkbook;
            internal Configuration(string path, string mode, string seedProfile, string sourceWorkbook = null)
            { Path = path; Mode = mode; SeedProfile = seedProfile; SourceWorkbook = sourceWorkbook; }
        }

        /// <summary>Runs before bootstrap; an absent flag reads no other configuration or filesystem metadata.</summary>
        internal static Configuration Prepare(Func<string, string> environment, string layout, bool persistence,
            Func<string, FileAttributes> metadata = null)
        {
            if (environment(OptIn) != "1") return null;
            if (persistence || layout == null || !Layouts.Contains(layout))
                throw new InvalidOperationException("Root font observation requires one declared import layout, not capture or persistence qualification.");
            string mode = environment(ModeVariable);
            if (mode != "ObserveWrites" && mode != "DistinctChildName" && mode != "AfterInitialCapture")
                throw new InvalidOperationException("An explicit supported root font observation mode is required.");
            string seedProfile = environment(SeedProfileVariable);
            if (seedProfile != null && (mode != "AfterInitialCapture" ||
                seedProfile != SyntheticExplicitArial9 && seedProfile != RetainedSyntheticTahoma825))
                throw new InvalidOperationException("The synthetic root font seed is restricted to the declared deferred diagnostic profile.");
            string sourceWorkbook = environment(SourceWorkbookVariable);
            if (seedProfile == RetainedSyntheticTahoma825)
            {
                if (layout != "LabelButton")
                    throw new InvalidOperationException("The retained synthetic workbook qualifies only LabelButton.");
                sourceWorkbook = RetainedRootFontWorkbook.RequirePinnedSource(sourceWorkbook,
                    metadata ?? File.GetAttributes);
            }
            else if (sourceWorkbook != null)
                throw new InvalidOperationException("A retained source workbook requires the exact retained font profile.");
            string path = ExactPath(environment(ManifestVariable));
            RequirePath(path, Entry.Absent, metadata ?? File.GetAttributes);
            RequirePath(System.IO.Path.GetDirectoryName(path), Entry.Directory, metadata ?? File.GetAttributes);
            return new Configuration(path, mode, seedProfile, sourceWorkbook);
        }

        /// <summary>Builds only managed data. It neither exports a form nor claims native font delivery.</summary>
        internal static Dictionary<string, object> Build(Configuration configuration, string evidenceRoot, string projectPath,
            VbaGitSnapshot baseline, Guid expectedCandidate, Guid loadedCandidate, Guid nonce)
        {
            if (configuration == null) throw new ArgumentNullException(nameof(configuration));
            if (configuration.Mode != "ObserveWrites" && configuration.Mode != "DistinctChildName" &&
                configuration.Mode != "AfterInitialCapture")
                throw new InvalidOperationException("Unsupported root font observation mode.");
            if (expectedCandidate == Guid.Empty || loadedCandidate != expectedCandidate || nonce == Guid.Empty)
                throw new InvalidOperationException("Root font observation must bind the exact loaded candidate and a fresh nonce.");
            projectPath = ExactPath(projectPath);
            evidenceRoot = ExactPath(evidenceRoot);
            if (baseline == null || baseline.Manifest?.Components == null)
                throw new InvalidOperationException("The independently captured baseline is required.");
            var forms = baseline.Manifest.Components.Where(item => item.Name == "EmbeddedForm" && item.Type == 3).ToArray();
            if (forms.Length != 1 || !forms[0].HasResources ||
                !baseline.Files.TryGetValue(forms[0].FileName, out byte[] source) || source == null || source.Length == 0 ||
                !baseline.Files.TryGetValue("EmbeddedForm.frx", out byte[] resource) || resource == null || resource.Length == 0)
                throw new InvalidOperationException("One resource-bearing EmbeddedForm with exact source bytes is required.");
            if (configuration.SeedProfile == RetainedSyntheticTahoma825)
                RetainedRootFontWorkbook.RequirePinnedBaseline(baseline);
            byte[] descriptor = RequireRoot(baseline.FormFonts(forms[0]), configuration.Mode, configuration.SeedProfile);
            return new Dictionary<string, object> {
                ["ProjectPath"] = projectPath, ["FormName"] = "EmbeddedForm", ["TargetFormSha256"] = Sha(source),
                ["TargetDescriptorHex"] = Hex(descriptor), ["CandidateMvid"] = expectedCandidate.ToString("D"),
                ["OutputRoot"] = System.IO.Path.Combine(evidenceRoot, nonce.ToString("N")), ["Nonce"] = nonce.ToString("N"),
                ["Mode"] = configuration.Mode, ["TemporaryName"] = configuration.Mode == "DistinctChildName" ? "Arial" : ""
            };
        }

        internal static byte[] RequireRoot(FormStreamPadding.FormFontBinding[] bindings, string mode = "ObserveWrites",
            string seedProfile = null)
        {
            var roots = bindings?.Where(item => item != null && item.OwnerPath == "" && item.Type == 7).ToArray();
            if (roots == null || roots.Length != 1 || roots[0].Descriptor == null ||
                roots[0].Descriptor.Length < 11 || roots[0].Descriptor.Length > 64 ||
                roots[0].Descriptor[10] == 0 || roots[0].Descriptor.Length != 11 + roots[0].Descriptor[10])
                throw new InvalidOperationException("Exactly one fully bounded root Type7 font descriptor is required.");
            if (mode == "DistinctChildName" && string.Equals(Encoding.ASCII.GetString(roots[0].Descriptor, 11,
                roots[0].Descriptor[10]), "Arial", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("DistinctChildName requires the fixed Arial face to differ from the target.");
            if (seedProfile != null && (mode != "AfterInitialCapture" ||
                seedProfile == SyntheticExplicitArial9 && !roots[0].Descriptor.SequenceEqual(SyntheticArial9Descriptor()) ||
                seedProfile == RetainedSyntheticTahoma825 && !roots[0].Descriptor.SequenceEqual(RetainedRootFontWorkbook.Tahoma825Descriptor()) ||
                seedProfile != SyntheticExplicitArial9 && seedProfile != RetainedSyntheticTahoma825))
                throw new InvalidOperationException("The saved and reopened synthetic root font differs from the declared exact seed.");
            return (byte[])roots[0].Descriptor.Clone();
        }

        internal static IDictionary<string, object> SyntheticArial9Values(string profile)
        {
            if (profile != SyntheticExplicitArial9)
                throw new InvalidOperationException("Unknown synthetic root font seed profile.");
            return new Dictionary<string, object> {
                ["Form.Font.Name"] = "Arial", ["Form.Font.Size"] = 9.00m,
                ["Form.Font.Weight"] = (short)400, ["Form.Font.Charset"] = (short)0,
                ["Form.Font.Italic"] = false, ["Form.Font.Underline"] = false,
                ["Form.Font.Strikethrough"] = false
            };
        }

        internal static byte[] SyntheticArial9Descriptor()
        { return new byte[] { 1, 0, 0, 0, 144, 1, 144, 95, 1, 0, 5, 65, 114, 105, 97, 108 }; }

        /// <summary>Rechecks inherited configuration and local paths, then claims the file once without creating output directories.</summary>
        internal static void Publish(Configuration configuration, Dictionary<string, object> manifest,
            Func<string, string> environment, Func<string, FileAttributes> metadata = null)
        {
            if (configuration == null) throw new ArgumentNullException(nameof(configuration));
            if (environment(OptIn) != "1" || environment(ModeVariable) != configuration.Mode ||
                environment(SeedProfileVariable) != configuration.SeedProfile ||
                environment(SourceWorkbookVariable) != configuration.SourceWorkbook ||
                !string.Equals(environment(ManifestVariable), configuration.Path, StringComparison.Ordinal))
                throw new InvalidOperationException("Root font diagnostic configuration changed after the inherited bootstrap preflight.");
            if (configuration.SeedProfile == RetainedSyntheticTahoma825)
                RetainedRootFontWorkbook.RequirePinnedSource(configuration.SourceWorkbook, metadata ?? File.GetAttributes);
            var inspect = metadata ?? File.GetAttributes;
            RequirePath(configuration.Path, Entry.Absent, inspect);
            RequirePath(System.IO.Path.GetDirectoryName(configuration.Path), Entry.Directory, inspect);
            RequirePath(Convert.ToString(manifest["ProjectPath"]), Entry.File, inspect);
            string output = ExactPath(Convert.ToString(manifest["OutputRoot"]));
            RequirePath(output, Entry.Absent, inspect);
            RequirePath(System.IO.Path.GetDirectoryName(output), Entry.Directory, inspect);
            if (Convert.ToString(manifest["Mode"]) != configuration.Mode ||
                !Guid.TryParseExact(Convert.ToString(manifest["Nonce"]), "N", out Guid nonce) || nonce == Guid.Empty ||
                System.IO.Path.GetFileName(output) != nonce.ToString("N"))
                throw new InvalidOperationException("The manifest output claim must remain nonce-bound and mode-bound.");
            byte[] bytes = new UTF8Encoding(false, true).GetBytes(new JavaScriptSerializer().Serialize(manifest));
            if (bytes.Length > 16384) throw new InvalidOperationException("Diagnostic manifest exceeds the product's bounded contract.");
            using (var stream = new FileStream(configuration.Path, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            { stream.Write(bytes, 0, bytes.Length); stream.Flush(true); }
        }

        private enum Entry { Absent, File, Directory }

        private static void RequirePath(string path, Entry expected, Func<string, FileAttributes> metadata)
        {
            path = ExactPath(path);
            FileAttributes? leaf = Attributes(path, metadata);
            if (leaf.HasValue && ((leaf.Value & FileAttributes.ReparsePoint) != 0 || expected == Entry.Absent ||
                ((leaf.Value & FileAttributes.Directory) != 0) != (expected == Entry.Directory)) ||
                !leaf.HasValue && expected != Entry.Absent)
                throw new InvalidOperationException("Diagnostic path is stale, missing, a filesystem link or has the wrong entry type.");
            for (string parent = System.IO.Path.GetDirectoryName(path); parent != null; parent = System.IO.Path.GetDirectoryName(parent))
            {
                var attributes = Attributes(parent, metadata);
                if (attributes.HasValue && ((attributes.Value & FileAttributes.ReparsePoint) != 0 ||
                    (attributes.Value & FileAttributes.Directory) == 0))
                    throw new InvalidOperationException("Diagnostic path traverses a filesystem link or non-directory.");
            }
        }

        private static FileAttributes? Attributes(string path, Func<string, FileAttributes> metadata)
        {
            try { return metadata(path); }
            catch (FileNotFoundException) { return null; }
            catch (DirectoryNotFoundException) { return null; }
        }

        private static string ExactPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || path.Length < 3 || !char.IsLetter(path[0]) || path[1] != ':' ||
                path[2] != '\\' || path.IndexOf(':', 2) >= 0 ||
                !string.Equals(path, System.IO.Path.GetFullPath(path), StringComparison.Ordinal))
                throw new InvalidOperationException("An exact absolute local diagnostic path without alternate streams is required.");
            return path;
        }
        private static string Hex(byte[] bytes) { return BitConverter.ToString(bytes).Replace("-", ""); }
        private static string Sha(byte[] bytes) { using (var hash = SHA256.Create()) return Hex(hash.ComputeHash(bytes)); }
    }
}
