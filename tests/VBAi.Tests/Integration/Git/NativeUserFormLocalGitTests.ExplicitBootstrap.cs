using System;
using System.IO;
using System.Runtime.ExceptionServices;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    public sealed partial class NativeUserFormLocalGitTests
    {
        /// <summary>Runs the unchanged full layout contract in a separately opted-in explicit Excel launch context.</summary>
        [STATestMethod, TestCategory("NativeUserFormExplicitBootstrap")]
        [DataRow("LabelButton")]
        [DataRow("TextBox")]
        [DataRow("ComboBox")]
        [DataRow("ListBox")]
        [DataRow("CheckBox")]
        [DataRow("OptionButton")]
        [DataRow("ToggleButton")]
        [DataRow("ScrollBar")]
        [DataRow("SpinButton")]
        [DataRow("TabStrip")]
        [DataRow("Image")]
        [DataRow("FrameMultiPage")]
        public void ExplicitOwnedLayoutCaptureImportRecoveryAndReopenPreserveNativeState(string layout)
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_USERFORM_EXPLICIT_BOOTSTRAP") != "1")
                Assert.Inconclusive("The distinct launch-context matrix requires VBAi_RUN_USERFORM_EXPLICIT_BOOTSTRAP=1.");
            RunLayout(layout, RunExplicit, "ExplicitXAutomation");
        }

        private static void RunExplicit(Action<ExcelVbeFixture> scenario, Action<ExcelVbeFixture> shutdownVerified)
        {
            if (!string.IsNullOrEmpty(Environment.GetEnvironmentVariable(PathVisibilityDiagnostic.EnvironmentName)))
                throw new InvalidOperationException("The layout matrix must not inherit a path-visibility/token manifest; no host was launched.");
            string output = ExcelOwnedBootstrapPlan.RequireLocalAbsolutePath(Environment.GetEnvironmentVariable("VBAi_EXCEL_RESULTS"));
            Directory.CreateDirectory(output);
            string trace = Path.Combine(output, "unused-inspection-" + Guid.NewGuid().ToString("N") + ".jsonl");
            // Reuse the existing exact-PID/seed/MVID bootstrap. No visibility/token manifest is supplied.
            var host = ExcelVbeFixture.StartOwnedWithTrace(trace);
            Exception failure = null;
            try { scenario(host); }
            catch (Exception error)
            {
                failure = error;
                // A timed-out asynchronous native operation must not be followed by a cleanup mutation.
                if (error is TimeoutException || error is IOException) host.PreserveForDiagnosticRecovery = true;
            }
            try { host.Dispose(); shutdownVerified?.Invoke(host); }
            catch (Exception cleanup)
            {
                if (failure != null) throw new AggregateException("Explicit Excel layout and its shutdown both failed; neither operation is replayed.", failure, cleanup);
                throw;
            }
            if (failure != null) ExceptionDispatchInfo.Capture(failure).Throw();
        }
    }
}
