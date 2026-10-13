using System;
using System.Collections.Generic;
using System.IO;
using System.Web.Script.Serialization;

namespace VBAi.Tests.Integration
{
    internal sealed partial class OfficeVbeFixture
    {
        private AddInShutdownDiagnostic.Identity wordShutdownIdentity;
        private string wordShutdownRoot;

        /// <summary>Admits an observer only for the original, durable Main Word fixture and its already observed VBE thread.</summary>
        internal static void RequireWordShutdownObservationScope(string kind, bool main, bool owned, int processId,
            IntPtr launchHandle, IntPtr processHandle, string evidenceRoot, IsolatedTestDesktop.MainWindow vbe)
        {
            if (kind != "Word" || !main || !owned || processId <= 0 || launchHandle == IntPtr.Zero ||
                processHandle == IntPtr.Zero || string.IsNullOrWhiteSpace(evidenceRoot))
                throw new InvalidOperationException("Shutdown observation requires the original owned Main Word fixture and durable evidence.");
            OfficeVbeFixture.RequireMainWordWindowObservation(vbe, new IntPtr(vbe?.Handle ?? 0),
                (uint)processId, vbe?.ThreadId ?? 0, true, "wndclass_desked_gsk");
        }

        /// <summary>Publishes one request from verified ready-state data without acquiring another COM reference.</summary>
        private void ArmWordShutdownObservation(IDictionary<string, object> loaded, IsolatedTestDesktop.MainWindow vbe)
        {
            string configured = Environment.GetEnvironmentVariable(AddInShutdownDiagnostic.EnvironmentName);
            if (string.IsNullOrWhiteSpace(configured)) return;
            RequireWordShutdownObservationScope(Kind, mainWordDesktop, owned, ProcessId,
                privateWordChild?.ProcessHandle ?? IntPtr.Zero, ownedProcess?.Handle ?? IntPtr.Zero, Root, vbe);
            RequireMainWordOriginal();
            string hash = Environment.GetEnvironmentVariable("VBAi_TEST_EMBEDDED_GIT_SHA256");
            string product = ExcelVbeFixture.RequireAddInShutdownProduct(loaded,
                Environment.GetEnvironmentVariable(ExcelVbeFixture.AddInShutdownProductEnvironmentName),
                Environment.GetEnvironmentVariable("VBAi_TEST_EMBEDDED_GIT_MVID"), hash,
                typeof(VbeSession).Assembly.Location, typeof(VbeSession).Module.ModuleVersionId, HashOwnedFile);
            var identity = new AddInShutdownDiagnostic.Identity
            {
                ProcessId = ProcessId, ProcessStartedUtc = mainWordBirth, HostImagePath = mainWordExecutable,
                ProductPath = product, ProductMvid = typeof(VbeSession).Module.ModuleVersionId.ToString("D"),
                ProductSha256 = hash, ThreadId = vbe.ThreadId
            };
            string root = AddInShutdownDiagnostic.RequireRoot(configured, Path.GetTempPath());
            string request = ExcelVbeFixture.PublishAddInShutdownRequest(root, identity);
            wordShutdownRoot = root;
            wordShutdownIdentity = identity;
            File.WriteAllText(Path.Combine(Root, "addin-shutdown-request.json"), new JavaScriptSerializer().Serialize(new
            {
                RequestPath = request, Nonce = Path.GetFileName(root), ExpectedIdentity = identity,
                FixtureRoot = Root, PublicationReturned = true, Utc = DateTime.UtcNow.ToString("o")
            }));
        }

        /// <summary>Reads the cleanup trace once after the original shutdown attempt, preserving its independent exit verdict.</summary>
        private void ObserveWordShutdownTrace()
        {
            if (wordShutdownIdentity == null) return;
            string output = Path.Combine(wordShutdownRoot, "process-" + wordShutdownIdentity.ProcessId);
            var observation = new Dictionary<string, object>
            {
                ["Phase"] = "WordAddInShutdownObservation", ["ObservationRoot"] = output,
                ["OriginalLifecycle"] = shutdownEvidence?.Record,
                ["HostExitProvenByTrace"] = false, ["NativeRcwReleaseProven"] = false
            };
            Exception failure = null;
            try { observation["Trace"] = AddInShutdownDiagnosticReceipt.Read(output, Path.GetFileName(wordShutdownRoot), wordShutdownIdentity); }
            catch (Exception error) { observation["TraceError"] = error.ToString(); failure = error; }
            steps.Add(observation);
            File.WriteAllText(Path.Combine(Root, "addin-shutdown-observed.json"), new JavaScriptSerializer().Serialize(observation));
            if (failure != null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
        }
    }
}
