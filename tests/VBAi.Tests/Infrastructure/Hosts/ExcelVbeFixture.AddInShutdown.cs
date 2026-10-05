using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;

namespace VBAi.Tests.Integration
{
    internal sealed partial class ExcelVbeFixture
    {
        internal const string AddInShutdownProductEnvironmentName = "VBAi_TEST_ADDIN_SHUTDOWN_PRODUCT";
        private AddInShutdownDiagnostic.Identity addInShutdownIdentity;
        private string addInShutdownRoot;
        private Exception addInShutdownFailure;

        /// <summary>Arms only an explicitly launched, already identified native host; reads no project or additional COM object.</summary>
        private void ArmAddInShutdownObservation(uint owningThread, IDictionary<string, object> loaded)
        {
            string configured = Environment.GetEnvironmentVariable(AddInShutdownDiagnostic.EnvironmentName);
            if (string.IsNullOrWhiteSpace(configured)) return;
            if (!owned || ownedProcess == null || ownedProcess.Id != ProcessId || !retainEvidence || ownedImagePath == null)
                throw new InvalidOperationException("Shutdown observation requires the exact explicitly launched durable fixture.");
            string hash = Environment.GetEnvironmentVariable("VBAi_TEST_EMBEDDED_GIT_SHA256");
            string product = RequireAddInShutdownProduct(loaded,
                Environment.GetEnvironmentVariable(AddInShutdownProductEnvironmentName),
                Environment.GetEnvironmentVariable("VBAi_TEST_EMBEDDED_GIT_MVID"), hash,
                typeof(VbeSession).Assembly.Location, typeof(VbeSession).Module.ModuleVersionId, ReadAddInShutdownProductHash);
            var identity = new AddInShutdownDiagnostic.Identity { ProcessId = ProcessId,
                ProcessStartedUtc = ownedProcess.StartTime.ToUniversalTime().ToString("o"), HostImagePath = ownedImagePath(),
                ProductPath = product, ProductMvid = typeof(VbeSession).Module.ModuleVersionId.ToString("D"), ProductSha256 = hash, ThreadId = owningThread };
            string root = AddInShutdownDiagnostic.RequireRoot(configured, Path.GetTempPath());
            string request = PublishAddInShutdownRequest(root, identity);
            addInShutdownRoot = root;
            addInShutdownIdentity = identity;
            WriteEvidence("addin-shutdown-request.json", new { RequestPath = request, Nonce = Path.GetFileName(root),
                ExpectedIdentity = identity, FixtureRoot = Root, PublicationReturned = true, Utc = DateTime.UtcNow.ToString("o") });
        }

        /// <summary>Compares independent manifest pins to installed status and local test-reference bytes, allowing different deployment paths.</summary>
        internal static string RequireAddInShutdownProduct(IDictionary<string, object> loaded, string expectedProduct,
            string expectedMvid, string expectedHash, string localProduct, Guid localMvid, Func<string, string> readHash)
        {
            string product = ExcelOwnedBootstrapPlan.RequireLocalAbsolutePath(expectedProduct);
            Guid pinnedMvid;
            if (!string.Equals(product, expectedProduct, StringComparison.OrdinalIgnoreCase) ||
                !Guid.TryParseExact(expectedMvid, "D", out pinnedMvid) || pinnedMvid != localMvid ||
                expectedHash == null || expectedHash.Length != 64 || expectedHash.Any(value => !Uri.IsHexDigit(value)) ||
                loaded == null || !loaded.TryGetValue("AssemblyPath", out var path) ||
                !string.Equals(path as string, product, StringComparison.OrdinalIgnoreCase) ||
                !loaded.TryGetValue("AssemblyModuleVersionId", out var mvid) ||
                !string.Equals(mvid as string, pinnedMvid.ToString("D"), StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The installed shutdown observer must match the independently pinned product path/MVID and local test image.");
            if (readHash == null) throw new ArgumentNullException(nameof(readHash));
            string local = ExcelOwnedBootstrapPlan.RequireLocalAbsolutePath(localProduct);
            if (!string.Equals(local, localProduct, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The local test-reference product path must be canonical.");
            string nativeHash = readHash(product);
            string localHash = string.Equals(local, product, StringComparison.OrdinalIgnoreCase) ? nativeHash : readHash(local);
            if (!string.Equals(nativeHash, expectedHash, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(localHash, expectedHash, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Installed and local test-reference product bytes must both match the frozen SHA-256.");
            return product;
        }

        private static string ReadAddInShutdownProductHash(string path)
        {
            using (var file = System.IO.File.OpenRead(path))
            using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(file)).Replace("-", "");
        }

        /// <summary>Publishes a canonical CreateNew request exactly once; file collisions never authorize overwrite.</summary>
        internal static string PublishAddInShutdownRequest(string root, AddInShutdownDiagnostic.Identity identity)
        {
            root = AddInShutdownDiagnostic.RequireRoot(root, Path.GetTempPath());
            AddInShutdownDiagnostic.RequireSameIdentity(identity, identity);
            string path = Path.Combine(root, "request-" + identity.ProcessId.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".json");
            string temporary = path + ".tmp";
            AddInShutdownDiagnostic.RequireNoReparse(path); AddInShutdownDiagnostic.RequireNoReparse(temporary);
            string text = new JavaScriptSerializer().Serialize(new { Version = 1, Nonce = Path.GetFileName(root), Identity = identity });
            AddInShutdownDiagnostic.DecodeRequest(text, Path.GetFileName(root));
            using (var file = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            using (var writer = new StreamWriter(file, new UTF8Encoding(false))) writer.Write(text);
            System.IO.File.Move(temporary, path);
            return path;
        }

        /// <summary>Reads a single bounded cleanup snapshot after the original exit wait; never waits for or replays cleanup.</summary>
        private void ObserveAddInShutdownTrace(IDictionary<string, object> diagnostics)
        {
            if (addInShutdownIdentity == null) return;
            string output = Path.Combine(addInShutdownRoot, "process-" + ProcessId.ToString(System.Globalization.CultureInfo.InvariantCulture));
            diagnostics["AddInShutdownObservationRoot"] = output;
            try
            {
                object trace = AddInShutdownDiagnosticReceipt.Read(output, Path.GetFileName(addInShutdownRoot), addInShutdownIdentity);
                diagnostics["AddInShutdownTrace"] = trace;
                WriteEvidence("addin-shutdown-observed.json", trace);
            }
            catch (Exception error)
            {
                addInShutdownFailure = error;
                diagnostics["AddInShutdownTraceError"] = error.ToString();
                RememberShutdownFailure(error);
            }
        }
    }
}
