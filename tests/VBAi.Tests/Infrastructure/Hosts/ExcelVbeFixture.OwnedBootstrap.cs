using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration
{
    internal sealed partial class ExcelVbeFixture
    {
        // Keep explicitly retained native objects/handles alive for diagnosis through this testhost's lifetime.
        private static readonly List<ExcelVbeFixture> retainedBootstraps = new List<ExcelVbeFixture>();
        // Explicit launches retain the original native handle; startup evidence must use the same reader.
        private Func<string> ownedImagePath;
        private IsolatedTestDesktop.NativeChild privateDesktopChild;

        /// <summary>Selects the inherited private-desktop trace or preserves the caller's main-desktop trace.</summary>
        internal static string SelectOwnedTracePath(string customTrace, string desktop, string inheritedTrace)
        {
            return string.IsNullOrWhiteSpace(desktop) ? customTrace :
                ExcelOwnedBootstrapPlan.RequireLocalAbsolutePath(inheritedTrace);
        }

        internal static string SelectOwnedTracePath(string customTrace)
        {
            return SelectOwnedTracePath(customTrace, Environment.GetEnvironmentVariable("VBAi_TEST_DESKTOP_NAME"),
                Environment.GetEnvironmentVariable(VbeInspectionTrace.EnvironmentName));
        }

        /// <summary>Refuses unsupported private-child settings before durable launch intent is recorded.</summary>
        internal static void RecordLaunchIntentAfterPrivateGuards(string desktop, string tracePath,
            string inheritedTrace, string pathVisibilityManifest, Action recordIntent)
        {
            if (recordIntent == null) throw new ArgumentNullException(nameof(recordIntent));
            if (!string.IsNullOrWhiteSpace(desktop))
            {
                if (pathVisibilityManifest != null)
                    throw new InvalidOperationException("Per-child path visibility is not supported by this private-desktop bootstrap.");
                if (string.IsNullOrWhiteSpace(inheritedTrace) || string.IsNullOrWhiteSpace(tracePath) ||
                    !string.Equals(inheritedTrace, tracePath, StringComparison.Ordinal))
                    throw new InvalidOperationException("Private Excel must inherit the exact frozen trace path.");
            }
            recordIntent();
        }

        /// <summary>Requires the owned Excel window on the observed launch thread's desktop.</summary>
        internal static void RequireOwnedWindowDesktop(string configuredPrivateDesktop,
            string launchThreadDesktop, string applicationWindowDesktop)
        {
            if (string.IsNullOrWhiteSpace(launchThreadDesktop) || string.IsNullOrWhiteSpace(applicationWindowDesktop) ||
                !string.Equals(launchThreadDesktop, applicationWindowDesktop, StringComparison.OrdinalIgnoreCase) ||
                (!string.IsNullOrWhiteSpace(configuredPrivateDesktop) &&
                 !string.Equals(configuredPrivateDesktop, applicationWindowDesktop, StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException("The owned Excel window desktop differs from its launch thread or configured private desktop.");
        }

        /// <summary>Explicit environment-controlled launch used only by the scalar diagnostic pages.</summary>
        internal static ExcelVbeFixture StartOwnedWithTrace(string tracePath, string pathVisibilityManifest = null)
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_EXCEL_TESTS") != "1")
                Assert.Inconclusive("Excel automation is opt-in. Set VBAi_RUN_EXCEL_TESTS=1.");
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                throw new InvalidOperationException("Owned Excel NativeOM attachment requires the test's STA thread.");
            int[] existingIds = ExistingExcelIds();
            ExcelOwnedBootstrapPlan.RequireFreshLaunch(existingIds);
            if (pathVisibilityManifest != null)
            {
                pathVisibilityManifest = ExcelOwnedBootstrapPlan.RequireLocalAbsolutePath(pathVisibilityManifest);
                if (!System.IO.File.Exists(pathVisibilityManifest)) throw new ArgumentException("The fixed path-visibility manifest must exist before launch.");
            }
            string executable = ExcelOwnedBootstrapPlan.ResolveExecutable();
            string desktop = Environment.GetEnvironmentVariable("VBAi_TEST_DESKTOP_NAME");
            if (!string.IsNullOrWhiteSpace(desktop)) IsolatedTestDesktop.RequireCurrent(desktop);
            uint launchThread = IsolatedTestDesktop.GetCurrentThreadId();
            string launchDesktop = IsolatedTestDesktop.DesktopName(launchThread);
            string inputDesktop = IsolatedTestDesktop.InputDesktopName();
            string output = Environment.GetEnvironmentVariable("VBAi_EXCEL_RESULTS");
            string parent = ExcelOwnedBootstrapPlan.RequireLocalAbsolutePath(string.IsNullOrWhiteSpace(output)
                ? Path.Combine(Path.GetTempPath(), "VBAi-VSTest") : output);
            var fixture = new ExcelVbeFixture { retainEvidence = true, Root = Path.Combine(parent, Guid.NewGuid().ToString("N")) };
            Directory.CreateDirectory(fixture.Root);
            string seed = fixture.File("OwnedScalarSeed.xlsx");
            var launch = new Dictionary<string, object> {
                ["Phase"] = "Preparing", ["StartedUtc"] = DateTime.UtcNow.ToString("o"), ["Executable"] = executable,
                ["ExecutableFileVersion"] = FileVersionInfo.GetVersionInfo(executable).FileVersion,
                ["Seed"] = seed, ["EnvironmentName"] = VbeInspectionTrace.EnvironmentName,
                ["TracePath"] = tracePath, ["StartAttempts"] = 0, ["ExistingExcelProcessIds"] = existingIds,
                ["Desktop"] = launchDesktop, ["LaunchThreadId"] = launchThread,
                ["LaunchThreadDesktop"] = launchDesktop, ["InputDesktopAtPreflight"] = inputDesktop,
                ["ExpectedAssemblyMvid"] = typeof(VbeSession).Module.ModuleVersionId.ToString("D"),
                ["BootstrapCloseAttempts"] = 0, ["BootstrapQuitAttempts"] = 0, ["ForceTerminationAttempts"] = 0,
                ["StartupFileLoadingSuppressed"] = true
            };
            Action record = () => fixture.WriteEvidence("owned-bootstrap.json", launch);
            try
            {
                ExcelOwnedBootstrapPlan.WriteSeed(seed);
                using (var sha = SHA256.Create())
                using (var bytes = System.IO.File.OpenRead(seed)) launch["SeedSha256"] = BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", "");
                var info = ExcelOwnedBootstrapPlan.CreateStartInfo(executable, seed, tracePath);
                // Set only this child's opt-in. The default scalar/bootstrap environment remains unchanged.
                if (pathVisibilityManifest != null)
                {
                    info.EnvironmentVariables[PathVisibilityDiagnostic.EnvironmentName] = pathVisibilityManifest;
                    launch["DiagnosticEnvironmentName"] = PathVisibilityDiagnostic.EnvironmentName;
                    launch["DiagnosticManifestPath"] = pathVisibilityManifest;
                }
                using (var trace = new FileStream(tracePath, FileMode.OpenOrCreate, FileAccess.Write, FileShare.Read))
                    if (trace.Length > VbeInspectionTrace.MaximumFileBytes - 65536)
                        throw new InvalidOperationException("Phase evidence has insufficient capacity; no Excel process was launched.");
                // Recheck after preparing the seed, before the only native launch.
                ExcelOwnedBootstrapPlan.RequireFreshLaunch(ExistingExcelIds());
                ExcelOwnedBootstrapPlan.ValidateExecutable(executable);
                Assert.AreEqual(launch["ExecutableFileVersion"], FileVersionInfo.GetVersionInfo(executable).FileVersion,
                    "Excel executable identity changed after preflight; no launch is permitted.");
                fixture.RecordStartup("ExplicitLaunchPrepared", existingIds);
                RecordLaunchIntentAfterPrivateGuards(desktop, tracePath,
                    Environment.GetEnvironmentVariable(VbeInspectionTrace.EnvironmentName), pathVisibilityManifest,
                    () => { launch["Phase"] = "LaunchIntent"; launch["StartAttempts"] = 1; record(); });
                if (string.IsNullOrWhiteSpace(desktop)) fixture.ownedProcess = Process.Start(info);
                else
                {
                    // The private child inherits the worker's exact trace opt-in. Unsupported
                    // per-child diagnostics were refused before launch intent.
                    fixture.privateDesktopChild = IsolatedTestDesktop.Launch(executable,
                        new[] { "/x", "/automation", seed }, Path.GetDirectoryName(seed), desktop);
                    if (fixture.privateDesktopChild.Wait(0))
                        throw new InvalidOperationException("Private Excel exited before identity capture; no relaunch.");
                    fixture.ownedProcess = Process.GetProcessById(fixture.privateDesktopChild.ProcessId);
                    launch["Desktop"] = desktop;
                    launch["OriginalLaunchHandle"] = fixture.privateDesktopChild.ProcessHandle.ToInt64();
                }
                if (fixture.ownedProcess == null) throw new InvalidOperationException("Process.Start returned no owned process; no COM activation fallback.");
                IntPtr retainedHandle = fixture.ownedProcess.Handle;
                if (string.IsNullOrWhiteSpace(desktop)) launch["OriginalLaunchHandle"] = retainedHandle.ToInt64();
                fixture.ownedImagePath = () => ExcelOwnedProcessImage.Read(retainedHandle);
                fixture.ProcessId = fixture.ownedProcess.Id;
                string processStart = fixture.ownedProcess.StartTime.ToUniversalTime().ToString("o");
                launch["ProcessId"] = fixture.ProcessId; launch["ProcessStartUtc"] = processStart;
                launch["ImageIdentityReader"] = "QueryFullProcessImageNameW";
                launch["Phase"] = "ProcessImageIdentityPending"; record();
                ExcelOwnedBootstrapPlan.VerifyAttachedIdentity(fixture.ProcessId, executable, processStart, fixture.ProcessId,
                    fixture.ownedImagePath(), fixture.ownedProcess.StartTime.ToUniversalTime().ToString("o"));
                launch["Phase"] = "NativeAttachmentPending"; record();
                var elapsed = Stopwatch.StartNew();
                while (elapsed.ElapsedMilliseconds < 20000)
                {
                    if (fixture.ownedProcess.HasExited) throw new InvalidOperationException("The explicitly launched Excel exited before attachment; no restart.");
                    try
                    {
                        fixture.application = ExcelOwnedApplication.Resolve(fixture.ProcessId,
                            () => throw new NativeDocumentNotReadyException());
                        break;
                    }
                    catch (NativeDocumentNotReadyException) { Thread.Sleep(50); } // Read-only discovery, never another launch.
                }
                if (fixture.application == null) throw new InvalidOperationException("The owned EXCEL7 document could not be attached; retain process and evidence.");
                dynamic excel = fixture.application;
                uint actualPid;
                uint windowThread = GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(excel.Hwnd)), out actualPid);
                string actualDesktop = IsolatedTestDesktop.DesktopName(windowThread);
                launch["ApplicationWindowThreadId"] = windowThread;
                launch["ApplicationWindowDesktop"] = actualDesktop;
                RequireOwnedWindowDesktop(desktop, launchDesktop, actualDesktop);
                if (!string.IsNullOrWhiteSpace(desktop))
                {
                    IsolatedTestDesktop.RequireCurrent(desktop);
                }
                ExcelOwnedBootstrapPlan.VerifyAttachedIdentity(fixture.ProcessId, executable, processStart, (int)actualPid,
                    fixture.ownedImagePath(), fixture.ownedProcess.StartTime.ToUniversalTime().ToString("o"));
                launch["ApplicationHwndProcessId"] = actualPid;
                fixture.workbooks = excel.Workbooks;
                Assert.AreEqual(1, Convert.ToInt32(((dynamic)fixture.workbooks).Count), "Only the command-line-owned macro-free seed workbook may be present.");
                fixture.workbook = ((dynamic)fixture.workbooks).Item(1);
                Assert.AreEqual(seed, Convert.ToString(((dynamic)fixture.workbook).FullName), true, "The initial workbook differs from the owned seed.");
                launch["InitialWorkbook"] = seed;
                launch["InitialWorkbookCount"] = 1;
                fixture.owned = true;
                launch["Phase"] = "ExactNativeApplicationAndSeedAttached"; record();
                fixture.RecordStartup("ExactNativeApplicationAndSeedAttached", existingIds);
                excel.Visible = true;
                excel.DisplayAlerts = false;
                object bars = null;
                try { bars = excel.CommandBars; ((dynamic)bars).ExecuteMso("VisualBasic"); }
                finally { Release(bars); }
                var status = fixture.Command("status");
                Assert.IsNotNull(status, "The owned VBE bridge is unavailable; retain process without retry.");
                Assert.AreEqual(true, status["Ok"]);
                var state = VbeBridgeClient.Object(status["Data"]);
                Assert.AreEqual(fixture.ProcessId, Convert.ToInt32(state["HostProcessId"]));
                Assert.AreEqual(typeof(VbeSession).Module.ModuleVersionId.ToString("D"), state["AssemblyModuleVersionId"]);
                fixture.RecordStartup("Ready", existingIds, status);
                launch["LoadedAssemblyMvid"] = state["AssemblyModuleVersionId"]; launch["Phase"] = "Ready"; record();
                return fixture;
            }
            catch (Exception primary)
            {
                if (fixture.ownedProcess != null) lock (retainedBootstraps) retainedBootstraps.Add(fixture);
                launch["FailedAtPhase"] = launch["Phase"];
                launch["Phase"] = "FailedPreserved"; launch["Error"] = primary.ToString();
                launch["Recovery"] = "No Close/Quit/termination was emitted. Inspect the recorded exact process, seed and evidence; never activate or relaunch as a fallback.";
                try { record(); }
                catch (Exception evidence) { throw new AggregateException("Owned bootstrap and evidence persistence failed; no host cleanup was attempted.", primary, evidence); }
                throw;
            }
        }

        private sealed class NativeDocumentNotReadyException : Exception { }

        private static int[] ExistingExcelIds()
        {
            var existing = Process.GetProcessesByName("EXCEL");
            try { return existing.Select(process => process.Id).ToArray(); }
            finally { foreach (var process in existing) process.Dispose(); }
        }
    }
}
