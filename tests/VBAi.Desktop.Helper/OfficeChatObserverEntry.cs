using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Web.Script.Serialization;
using VBAi.Tests.Integration;

namespace VBAi.Desktop.Helper
{
    /// <summary>Explicit manifest for the owned Word UI observer; no host activation or shutdown is exposed.</summary>
    public sealed class OfficeChatObserverPlan
    {
        public int Version { get; set; }
        public string Nonce { get; set; }
        public string UserSid { get; set; }
        public string Helper { get; set; }
        public string HelperSha256 { get; set; }
        public string Tests { get; set; }
        public string TestsSha256 { get; set; }
        public string TestsMvid { get; set; }
        public string Product { get; set; }
        public string ProductSha256 { get; set; }
        public string ProductMvid { get; set; }
        public int ParentPid { get; set; }
        public string ParentBirthUtc { get; set; }
        public string ParentImage { get; set; }
        public int HostPid { get; set; }
        public string HostBirthUtc { get; set; }
        public string HostImage { get; set; }
        public string Label { get; set; }
        public string Project { get; set; }
        public string Module { get; set; }
        public string Marker { get; set; }
        public string Model { get; set; }
        public string Profile { get; set; }
        public string Culture { get; set; }
        public string Receipt { get; set; }
    }

    /// <summary>Durable observer result. A functional pass is distinct from both observer exit and host shutdown.</summary>
    public sealed class OfficeChatObserverReceipt
    {
        public string Nonce { get; set; }
        public int ProcessId { get; set; }
        public int HostPid { get; set; }
        public string HostBirthUtc { get; set; }
        public string TestsMvid { get; set; }
        public string ProductMvid { get; set; }
        public bool Completed { get; set; }
        public bool Settled { get; set; }
        public bool StreamedWhileBusy { get; set; }
        public bool CancellationAcknowledged { get; set; }
        public bool NextSendCompleted { get; set; }
        public bool NativeMarkerVisible { get; set; }
        public string Error { get; set; }
        public object[] Observations { get; set; }
    }

    /// <summary>Starts only the fixed, hash-bound test observer on the real desktop. It cannot run an arbitrary method.</summary>
    public static class OfficeChatObserverEntry
    {
        public static string Hash(string path)
        {
            using (var file = File.OpenRead(path))
            using (var sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(file)).Replace("-", "");
        }

        /// <summary>Rejects incomplete or replaced inputs before any UI Automation provider is acquired.</summary>
        public static void Validate(OfficeChatObserverPlan plan, string manifest, string ownImage,
            Func<string, bool> exists, Func<string, string> hash)
        {
            if (plan == null || exists == null || hash == null) throw new ArgumentNullException(nameof(plan));
            Guid nonce, testsMvid, productMvid;
            if (plan.Version != 1 || !Guid.TryParseExact(plan.Nonce, "N", out nonce) ||
                !Guid.TryParse(plan.TestsMvid, out testsMvid) || !Guid.TryParse(plan.ProductMvid, out productMvid) ||
                plan.ParentPid <= 0 || plan.HostPid <= 0 || plan.ParentPid == plan.HostPid ||
                string.IsNullOrEmpty(plan.UserSid) || string.IsNullOrWhiteSpace(plan.Label) ||
                string.IsNullOrWhiteSpace(plan.Project) || plan.Module != "Q028Marker" ||
                string.IsNullOrWhiteSpace(plan.Model) || string.IsNullOrWhiteSpace(plan.Profile) ||
                plan.Marker == null || !plan.Marker.StartsWith("NATIVE_", StringComparison.Ordinal) || plan.Marker.Length != 39 ||
                !Guid.TryParseExact(plan.Marker.Substring(7), "N", out nonce))
                throw new InvalidOperationException("An exact owned Word observer manifest is required.");
            CultureInfo.GetCultureInfo(plan.Culture ?? throw new InvalidOperationException("Culture required."));
            RequireBirth(plan.ParentBirthUtc); RequireBirth(plan.HostBirthUtc);
            foreach (var path in new[] { manifest, ownImage, plan.Helper, plan.Tests, plan.Product, plan.Receipt,
                plan.ParentImage, plan.HostImage })
                if (string.IsNullOrEmpty(path) || !Path.IsPathRooted(path) || path.StartsWith(@"\\", StringComparison.Ordinal))
                    throw new InvalidOperationException("Local absolute paths are required.");
            if (!SamePath(ownImage, plan.Helper) || Path.GetFileName(plan.Helper) != "VBAi.Desktop.Helper.exe" ||
                Path.GetFileName(plan.Tests) != "VBAi.Tests.dll" || Path.GetFileName(plan.Product) != "VBAi.dll" ||
                !string.Equals(Path.GetFileName(plan.HostImage), "WINWORD.EXE", StringComparison.OrdinalIgnoreCase) ||
                !SamePath(Path.GetDirectoryName(plan.Tests), Path.GetDirectoryName(plan.Helper)) ||
                !SamePath(Path.GetDirectoryName(plan.Product), Path.GetDirectoryName(plan.Tests)) ||
                !SamePath(Path.GetDirectoryName(plan.Receipt), Path.GetDirectoryName(manifest)) ||
                Path.GetFileName(plan.Receipt) != "observer-receipt.json" || exists(plan.Receipt))
                throw new InvalidOperationException("Observer inputs or one-shot output do not match the owned scope.");
            foreach (var pair in new[] { Tuple.Create(plan.Helper, plan.HelperSha256),
                Tuple.Create(plan.Tests, plan.TestsSha256), Tuple.Create(plan.Product, plan.ProductSha256) })
                if (!exists(pair.Item1) || string.IsNullOrEmpty(pair.Item2) || hash(pair.Item1) != pair.Item2)
                    throw new InvalidOperationException("A frozen observer assembly changed or is unavailable.");
        }

        private static bool SamePath(string left, string right) =>
            string.Equals(Path.GetFullPath(left), Path.GetFullPath(right), StringComparison.OrdinalIgnoreCase);

        private static DateTime RequireBirth(string value)
        {
            DateTime birth;
            if (!DateTime.TryParseExact(value, "o", CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out birth) ||
                birth.Kind != DateTimeKind.Utc) throw new InvalidOperationException("Original UTC process birth required.");
            return birth;
        }

        private static void RequireProcess(int pid, string birth, string image)
        {
            using (var process = Process.GetProcessById(pid))
                if (process.HasExited || process.StartTime.ToUniversalTime() != RequireBirth(birth) ||
                    !SamePath(process.MainModule.FileName, image))
                    throw new InvalidOperationException("Original observer parent or owned Word process changed.");
        }

        /// <summary>Requires the original creation-handle exit and a matching terminal receipt before host teardown.</summary>
        public static void RequireTerminal(OfficeChatObserverPlan plan, OfficeChatObserverReceipt receipt,
            int childPid, bool originalExitObserved, uint exitCode)
        {
            if (plan == null || receipt == null || childPid <= 0 || !originalExitObserved ||
                (exitCode != 0 && exitCode != 1) || receipt.Nonce != plan.Nonce || receipt.ProcessId != childPid ||
                receipt.HostPid != plan.HostPid || receipt.HostBirthUtc != plan.HostBirthUtc ||
                receipt.TestsMvid != plan.TestsMvid || receipt.ProductMvid != plan.ProductMvid ||
                !receipt.Completed || !receipt.Settled)
                throw new InvalidOperationException("The original observer has no verified settled terminal result; retain the host without replay.");
            bool pass = receipt.Error == null && receipt.StreamedWhileBusy && receipt.CancellationAcknowledged &&
                receipt.NextSendCompleted && receipt.NativeMarkerVisible;
            if ((exitCode == 0) != pass || (exitCode == 1 && string.IsNullOrWhiteSpace(receipt.Error)))
                throw new InvalidOperationException("Observer exit and functional receipt disagree; retain the host.");
        }

        internal static int Run(string manifest)
        {
            if (!Path.IsPathRooted(manifest) || !File.Exists(manifest) || new FileInfo(manifest).Length > 65536) return 3;
            string json = File.ReadAllText(manifest);
            var plan = new JavaScriptSerializer().Deserialize<OfficeChatObserverPlan>(json);
            Validate(plan, manifest, typeof(OfficeChatObserverEntry).Assembly.Location, File.Exists, Hash);
            if (Environment.GetEnvironmentVariable("VBAi_RUN_OLLAMA_OFFICE_TESTS") != "1" ||
                Environment.GetEnvironmentVariable("VBAi_RUN_OLLAMA_OFFICE_MAIN_DESKTOP_TESTS") != "1" ||
                !string.IsNullOrEmpty(Environment.GetEnvironmentVariable("VBAi_QUALIFICATION_DESKTOP")) ||
                !string.IsNullOrEmpty(Environment.GetEnvironmentVariable("VBAi_TEST_DESKTOP_NAME")) ||
                WindowsIdentity.GetCurrent().User.Value != plan.UserSid) return 3;
            IsolatedTestDesktop.RequireMainCurrent();
            RequireProcess(plan.ParentPid, plan.ParentBirthUtc, plan.ParentImage);
            RequireProcess(plan.HostPid, plan.HostBirthUtc, plan.HostImage);
            var tests = Assembly.LoadFrom(plan.Tests);
            var product = Assembly.LoadFrom(plan.Product);
            if (tests.ManifestModule.ModuleVersionId.ToString("D") != plan.TestsMvid ||
                product.ManifestModule.ModuleVersionId.ToString("D") != plan.ProductMvid) return 3;
            // Fixed signature only. The helper has no build reference back to its parent test assembly.
            var entry = tests.GetType("VBAi.Tests.Integration.OllamaOfficeObserverWorker", true)
                .GetMethod("Run", BindingFlags.Public | BindingFlags.Static, null, new[] { typeof(string) }, null);
            if (entry == null || entry.ReturnType != typeof(int)) return 3;
            return (int)entry.Invoke(null, new object[] { json });
        }
    }
}
