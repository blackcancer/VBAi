using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace VBAi
{
    internal sealed partial class VbeTestExplorerService
    {
        private bool PreferReturnedValues() => IsExecutionHost() && (!IsNativeExecutionHost() || NativeRuntimeRegistrationReason() != null);

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

        internal static string NativeRuntimeRegistrationReason()
        {
            try
            {
                using (var classes = RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, RegistryView.Registry64))
                using (var server = classes.OpenSubKey(@"CLSID\{5AF2F40B-939B-4CC6-A06C-F0C79841C031}\InprocServer32"))
                using (var progId = classes.OpenSubKey(@"VBAi.TestRuntime\CLSID"))
                {
                    if (server == null) return "Register the VBAi.TestRuntime callback from the same VBAi build before native test execution.";
                    string codeBase = server.GetValue("CodeBase") as string;
                    if (!Uri.TryCreate(codeBase, UriKind.Absolute, out Uri uri) || !uri.IsFile ||
                        !string.Equals(Path.GetFullPath(uri.LocalPath), Path.GetFullPath(typeof(VbaTestRuntime).Assembly.Location), StringComparison.OrdinalIgnoreCase) ||
                        !string.Equals(server.GetValue("Class") as string, typeof(VbaTestRuntime).FullName, StringComparison.Ordinal) ||
                        !string.Equals(server.GetValue("Assembly") as string, typeof(VbaTestRuntime).Assembly.FullName, StringComparison.Ordinal) ||
                        !string.Equals(server.GetValue("RuntimeVersion") as string, "v4.0.30319", StringComparison.Ordinal) ||
                        !string.Equals(server.GetValue("ThreadingModel") as string, "Both", StringComparison.OrdinalIgnoreCase) ||
                        !string.Equals(server.GetValue("") as string, "mscoree.dll", StringComparison.OrdinalIgnoreCase) ||
                        !string.Equals(progId?.GetValue("") as string, "{5AF2F40B-939B-4CC6-A06C-F0C79841C031}", StringComparison.OrdinalIgnoreCase))
                        return "The registered VBAi.TestRuntime callback does not belong to this loaded VBAi build.";
                }
                return null;
            }
            catch (Exception error) { return "The callback registration could not be verified: " + error.Message; }
        }
    }
}
