using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace VBAi.Tests.Integration
{
    /// <summary>Observes existing Outlook add-in cleanup without changing the original shutdown verdict.</summary>
    internal sealed partial class OutlookVbaTestFixture
    {
        private AddInShutdownDiagnostic.Identity outlookShutdownIdentity;
        private string outlookShutdownRoot;
        private bool outlookShutdownTraceRead;

        /// <summary>Checks the disposable host and chooses one independently observed VBE window.</summary>
        internal static IsolatedTestDesktop.MainWindow RequireOutlookShutdownScope(bool main, bool baselineAbsent,
            bool baselineVerified, int pid, IntPtr processHandle, string root, IsolatedTestDesktop.MainWindow[] windows)
        {
            if (!main || !baselineAbsent || !baselineVerified || pid <= 0 || processHandle == IntPtr.Zero ||
                string.IsNullOrWhiteSpace(root) || !Path.IsPathRooted(root) || windows == null || windows.Length > 256 ||
                windows.Any(window => window == null || window.ProcessId != (uint)pid))
                throw new InvalidOperationException("Outlook cleanup observation requires a verified disposable Main host and complete owned window inventory.");
            var frames = windows.Where(window => window.ClassName == "wndclass_desked_gsk").ToArray();
            if (frames.Length != 1 || frames[0].Handle <= 0 || frames[0].ThreadId == 0 || !frames[0].Visible)
                throw new InvalidOperationException("One verified visible VBE owner window is required for Outlook cleanup observation.");
            return frames[0];
        }

        /// <summary>Publishes one optional request using existing ready-state data and no new COM acquisition.</summary>
        internal void ArmOutlookShutdownObservation(IDictionary<string, object> loaded)
        {
            string configured = Environment.GetEnvironmentVariable(AddInShutdownDiagnostic.EnvironmentName);
            if (string.IsNullOrWhiteSpace(configured)) return;
            if (outlookShutdownIdentity != null) throw new InvalidOperationException("Outlook shutdown observation is already armed; no repeated publication.");
            bool main = Environment.GetEnvironmentVariable("VBAi_RUN_MAIN_DESKTOP_TESTS") == "1" && privateDesktopChild == null;
            if (!main || !baselineAbsent || !baselineVerified || process == null)
                throw new InvalidOperationException("Outlook shutdown observation is restricted to the ready disposable Main fixture.");
            VerifyOwnedProcess();
            var inventory = IsolatedTestDesktop.ReadMainWindows((uint)ProcessId);
            var vbe = RequireOutlookShutdownScope(main, baselineAbsent, baselineVerified, ProcessId, process.Handle, Root, inventory.Windows);
            string hash = Environment.GetEnvironmentVariable("VBAi_TEST_EMBEDDED_GIT_SHA256");
            string product = ExcelVbeFixture.RequireAddInShutdownProduct(loaded,
                Environment.GetEnvironmentVariable(ExcelVbeFixture.AddInShutdownProductEnvironmentName),
                Environment.GetEnvironmentVariable("VBAi_TEST_EMBEDDED_GIT_MVID"), hash,
                typeof(VbeSession).Assembly.Location, typeof(VbeSession).Module.ModuleVersionId, HashFile);
            var identity = new AddInShutdownDiagnostic.Identity
            {
                ProcessId = ProcessId, ProcessStartedUtc = process.StartTime.ToUniversalTime().ToString("o"),
                HostImagePath = process.MainModule.FileName, ProductPath = product, ProductMvid = typeof(VbeSession).Module.ModuleVersionId.ToString("D"),
                ProductSha256 = hash, ThreadId = vbe.ThreadId
            };
            string root = AddInShutdownDiagnostic.RequireRoot(configured, Path.GetTempPath());
            string request = ExcelVbeFixture.PublishAddInShutdownRequest(root, identity);
            outlookShutdownIdentity = identity;
            outlookShutdownRoot = root;
            Save("addin-shutdown-request.json", new { RequestPath = request, ExpectedIdentity = identity,
                Nonce = Path.GetFileName(root), PublicationReturned = true, NoAdditionalComAcquisition = true });
        }

        /// <summary>Reads a bounded trace once; trace success never overrides a failed Quit or exit observation.</summary>
        internal void ObserveOutlookShutdownTrace(List<string> failures)
        {
            if (outlookShutdownIdentity == null || outlookShutdownTraceRead) return;
            outlookShutdownTraceRead = true;
            string output = Path.Combine(outlookShutdownRoot, "process-" + outlookShutdownIdentity.ProcessId);
            var observation = new Dictionary<string, object>
            {
                ["ObservationRoot"] = output, ["OriginalShutdownVerdictUnchanged"] = true,
                ["HostExitProvenByTrace"] = false, ["NativeRcwReleaseProven"] = false
            };
            try { observation["Trace"] = AddInShutdownDiagnosticReceipt.Read(output, Path.GetFileName(outlookShutdownRoot), outlookShutdownIdentity); }
            catch (Exception error) { observation["TraceError"] = error.ToString(); failures.Add("Outlook add-in shutdown observation: " + error.Message); }
            try { Save("addin-shutdown-observed.json", observation); }
            catch (Exception error) { failures.Add("Outlook add-in shutdown observation receipt: " + error.Message); }
        }
    }
}
