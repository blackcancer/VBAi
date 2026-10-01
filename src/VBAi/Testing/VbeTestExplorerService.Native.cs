using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace VBAi
{
    internal sealed partial class VbeTestExplorerService
    {
        private bool PreferReturnedValues() => IsExecutionHost() && (!IsNativeExecutionHost() || NativeRuntimeReason() != null);

        private void NativeExecutionGuard()
        {
            RequireOwner();
            active?.ExecutionGuard?.Invoke();
            if (active?.Stop.IsCancellationRequested == true && nativePhase != "TestCleanup" && nativePhase != "ModuleCleanup")
                throw new VbaTestInvocationException("The test run was stopped before native dispatch.", false);
        }

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
        internal static object ReadRuntimeRegistrationValue(string path, string name)
        {
            using (var classes = RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Registry64))
            using (var key = classes.OpenSubKey(path))
                return name == null ? (object)(key != null) : key?.GetValue(name);
        }

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
