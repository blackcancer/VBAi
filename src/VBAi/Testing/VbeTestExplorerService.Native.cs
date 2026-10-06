using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace VBAi
{

    /// <summary>Implements owner-thread native test dispatch and verifies callback registration.</summary>
    internal sealed partial class VbeTestExplorerService
    {

        /// <summary>Chooses returned-value assertions when native dispatch cannot provide trustworthy callback results.</summary>
        /// <returns>True for a non-native execution host, or when the native runtime registration is invalid.</returns>
        private bool PreferReturnedValues() => IsExecutionHost() && (!IsNativeExecutionHost() || NativeRuntimeReason() != null);

        /// <summary>Requires the owning thread and stops dispatch after cancellation outside cleanup phases.</summary>
        private void NativeExecutionGuard()
        {
            RequireOwner();
            active?.ExecutionGuard?.Invoke();
            if (active?.Stop.IsCancellationRequested == true && nativePhase != "TestCleanup" && nativePhase != "ModuleCleanup")
                throw new VbaTestInvocationException("The test run was stopped before native dispatch.", false);
        }

        /// <summary>Dispatches one test phase to the registered native callback on the owning thread.</summary>
        /// <param name="catalog">Current catalog used to validate project identity and revision before dispatch.</param>
        /// <param name="procedure">Discovered procedure selected for this phase.</param>
        /// <param name="phase">Supported lifecycle phase name, such as setup, invocation, or cleanup.</param>
        /// <returns>Callback result after owner-thread completion.</returns>
        /// <exception cref="VbaTestInvocationException">Dispatch is unavailable or its outcome is uncertain.</exception>
        private async Task<VbaTestResult> InvokeNativeAsync(VbaTestCatalog catalog, VbaTestDescriptor procedure, string phase)
        {
            nativePhase = phase;
            try
            {
                string reason = ExecutionUnavailableReason(catalog);
                if (reason != null) throw new VbaTestInvocationException(reason, false);
                return await AwaitOwner(nativeExecutionHost.InvokeAsync(catalog, procedure, phase));
            }
            catch (VbaTestInvocationException error) { if (error.Uncertain) outcomeUnknown = true; throw; }
            finally { nativePhase = null; }
        }

        // Read-only boundary: tests can verify every registration field without editing HKCR.
        /// <summary>Reads a 64-bit HKCR registration key or one named value without writing registry state.</summary>
        /// <param name="path">Registry subkey path relative to 64-bit Classes Root.</param>
        /// <param name="name">Value name, or null to test whether the key exists.</param>
        /// <returns>True/false for a null value name; otherwise the stored value or null when absent.</returns>
        internal static object ReadRuntimeRegistrationValue(string path, string name)
        {
            using (var classes = RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Registry64))
            using (var key = classes.OpenSubKey(path))
                return name == null ? (object)(key != null) : key?.GetValue(name);
        }

        /// <summary>Checks that the 64-bit callback registration matches the currently loaded VBAi assembly.</summary>
        /// <param name="readValue">Optional read-only registry accessor, used to inspect registration without changing it.</param>
        /// <returns>Null only when CLSID, ProgID, assembly path/name, class, CLR version, and threading model match; otherwise the refusal reason.</returns>
        internal static string NativeRuntimeRegistrationReason(Func<string, string, object> readValue = null)
        {
            try
            {
                readValue = readValue ?? ReadRuntimeRegistrationValue;
                const string server = @"CLSID\{5AF2F40B-939B-4CC6-A06C-F0C79841C031}\InprocServer32";
                const string progId = @"VBAi.TestRuntime\CLSID";
                if (!(bool)readValue(server, null)) return "Register the VBAi.TestRuntime callback from the same VBAi build before native test execution.";
                string codeBase = readValue(server, "CodeBase") as string;
                if (!Uri.TryCreate(codeBase, UriKind.Absolute, out Uri uri) || !uri.IsFile ||
                    !string.Equals(Path.GetFullPath(uri.LocalPath), Path.GetFullPath(typeof(VbaTestRuntime).Assembly.Location), StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(readValue(server, "Class") as string, typeof(VbaTestRuntime).FullName, StringComparison.Ordinal) ||
                    !string.Equals(readValue(server, "Assembly") as string, typeof(VbaTestRuntime).Assembly.FullName, StringComparison.Ordinal) ||
                    !string.Equals(readValue(server, "RuntimeVersion") as string, "v4.0.30319", StringComparison.Ordinal) ||
                    !string.Equals(readValue(server, "ThreadingModel") as string, "Both", StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(readValue(server, "") as string, "mscoree.dll", StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(readValue(progId, "") as string, "{5AF2F40B-939B-4CC6-A06C-F0C79841C031}", StringComparison.OrdinalIgnoreCase))
                    return "The registered VBAi.TestRuntime callback does not belong to this loaded VBAi build.";
                return null;
            }
            catch (Exception error) { return "The callback registration could not be verified: " + error.Message; }
        }
    }
}
