using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;

namespace VBAi.Tests.Integration
{
    internal sealed partial class ExcelVbeFixture
    {
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
            string product = typeof(VbeSession).Assembly.Location;
            if (!loaded.TryGetValue("AssemblyPath", out var path) || !string.Equals(path as string, product, StringComparison.OrdinalIgnoreCase) ||
                !loaded.TryGetValue("AssemblyModuleVersionId", out var mvid) || mvid as string != typeof(VbeSession).Module.ModuleVersionId.ToString("D"))
                throw new InvalidOperationException("The installed shutdown observer must be the exact expected product path/MVID.");
            string hash;
            using (var file = System.IO.File.OpenRead(product))
            using (var sha = SHA256.Create()) hash = BitConverter.ToString(sha.ComputeHash(file)).Replace("-", "");
            var identity = new AddInShutdownDiagnostic.Identity { ProcessId = ProcessId,
                ProcessStartedUtc = ownedProcess.StartTime.ToUniversalTime().ToString("o"), HostImagePath = ownedImagePath(),
                ProductPath = product, ProductMvid = mvid as string, ProductSha256 = hash, ThreadId = owningThread };
            string root = AddInShutdownDiagnostic.RequireRoot(configured, Path.GetTempPath());
            string request = PublishAddInShutdownRequest(root, identity);
            addInShutdownRoot = root;
            addInShutdownIdentity = identity;
            WriteEvidence("addin-shutdown-request.json", new { RequestPath = request, Nonce = Path.GetFileName(root),
                ExpectedIdentity = identity, FixtureRoot = Root, PublicationReturned = true, Utc = DateTime.UtcNow.ToString("o") });
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
