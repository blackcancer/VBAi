using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Web.Script.Serialization;

namespace VBAi.Tests.Integration.Hosts.Excel
{
    /// <summary>Full native Format qualification with terminal-only restoration and durable owned-host evidence.</summary>
    [TestClass, TestCategory("Excel")]
    public sealed partial class ExcelFormatOptionsTests
    {
        public TestContext TestContext { get; set; }
        private static readonly List<ExcelVbeFixture> retainedFormatHosts = new List<ExcelVbeFixture>();
        private int evidenceSequence;
        private string evidenceDirectory;
        private delegate bool WindowCallback(IntPtr handle, IntPtr parameter);
        [DllImport("user32.dll", SetLastError = true)] private static extern bool EnumWindows(WindowCallback callback, IntPtr parameter);
        [DllImport("user32.dll", SetLastError = true)] private static extern uint GetWindowThreadProcessId(IntPtr handle, out uint processId);
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr handle);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr handle, StringBuilder text, int capacity);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr handle, StringBuilder text, int capacity);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowTextLength(IntPtr handle);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr handle);

        [STATestMethod]
        public void NativeFormatChoicesRoundTripAndRestoreCompleteOptionsVersion()
        { RunQualification(false); }

        [STATestMethod]
        public void NativeMarginCheckboxRoundTripAndRestoreCompleteOptionsVersion()
        { RunQualification(true); }

        [STATestMethod]
        public void NativeHistoricalFontPalettePrefixAndRestoreCompleteOptionsVersion()
        { RunQualification(false, true); }

        [STATestMethod]
        public void NativeHistoricalCompleteFormatAndRestoreCompleteOptionsVersion()
        { RunQualification(false, true, historicalFullMatrix: true); }

        [STATestMethod]
        public void NativeFontSizeCatalogueAndRestoreCompleteOptionsVersion()
        { RunQualification(false, fontSizeOnly: true); }

        private void RunQualification(bool marginOnly, bool historicalPalettePrefix = false, bool fontSizeOnly = false, bool historicalFullMatrix = false)
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_EXCEL_TESTS") != "1")
                Assert.Inconclusive("Excel automation is opt-in. Set VBAi_RUN_EXCEL_TESTS=1.");
            Assert.AreEqual(9, ExcelFormatOptionsQualification.Scenarios.Length);
            string output = Environment.GetEnvironmentVariable("VBAi_TEST_FORMAT_OPTIONS_OUTPUT");
            if (!string.IsNullOrEmpty(output) && !Path.IsPathRooted(output))
                throw new ArgumentException("VBAi_TEST_FORMAT_OPTIONS_OUTPUT must be an absolute path.");
            evidenceDirectory = Path.Combine(string.IsNullOrEmpty(output) ? TestContext.TestRunDirectory : output,
                "options-evidence-" + Guid.NewGuid().ToString("N"));
            ExcelFormatOptionsQualification.RunOwned(Environment.GetEnvironmentVariable("VBAi_RUN_EXCEL_TESTS") == "1",
                Environment.GetEnvironmentVariable("VBAi_EXCEL_RESULTS"), evidenceDirectory,
                Environment.GetEnvironmentVariable(PathVisibilityDiagnostic.EnvironmentName), () =>
                {
                    Directory.CreateDirectory(evidenceDirectory);
                    TestContext.WriteLine("Retained options evidence: " + evidenceDirectory);
                }, trace =>
                {
                    // The private child inherits this testhost's exact per-case trace opt-in.
                    if (!string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("VBAi_TEST_DESKTOP_NAME")))
                        Environment.SetEnvironmentVariable(VbeInspectionTrace.EnvironmentName, trace);
                    return ExcelVbeFixture.StartOwnedWithTrace(trace);
                }, host => QualifyReadyHost(host, marginOnly, historicalPalettePrefix, fontSizeOnly, historicalFullMatrix));
        }

        private void QualifyReadyHost(ExcelVbeFixture host, bool marginOnly, bool historicalPalettePrefix, bool fontSizeOnly, bool historicalFullMatrix)
        {
            // No using/finally Dispose: uncertainty must retain this exact fixture and all owning COM references.
            DateTime startUtc = DateTime.MinValue;
            try { using (var process = Process.GetProcessById(host.ProcessId)) startUtc = process.StartTime.ToUniversalTime(); }
            catch (Exception identityFailure)
            {
                RetainHost(host);
                try { AttachEvidence(host, startUtc, "StartupIdentityUnavailable", new { Error = identityFailure.ToString(), CleanupAllowed = false }); }
                catch (Exception recording) { throw new AggregateException("Startup identity and its evidence both failed; host retained.", identityFailure, recording); }
                throw;
            }
            Q026OptionsGuardTrace trace = null;
            var lifecycle = new ExcelFormatOptionsQualification(host.ProcessId, host.Command,
                () => ObserveOptionsClosureSettled(host, startUtc), () => RetainHost(host),
                () => { trace?.Dispose(); host.Dispose(); AttachEvidence(host, startUtc, "ShutdownVerified", host.ShutdownDiagnostics); },
                (phase, data) => AttachEvidence(host, startUtc, phase, data), verifyReadStability: !historicalPalettePrefix, marginOnly: marginOnly,
                historicalPalettePrefix: historicalPalettePrefix, verifyExclusiveHost: () => VerifyExclusiveHost(host, startUtc), fontSizeOnly: fontSizeOnly,
                historicalFullMatrix: historicalFullMatrix);
            Exception primary = null, detach = null;
            bool runStarted = false;
            try
            {
                if (Q026OptionsGuardTrace.NeedsGuardWarmupIfRequested()) lifecycle.WarmGuardForBreakpoint();
                trace = Q026OptionsGuardTrace.StartIfRequested(host, startUtc, evidenceDirectory);
                runStarted = true;
                lifecycle.Run();
            }
            catch (Exception error) { primary = error; if (!runStarted) RetainHost(host); }
            try { trace?.Dispose(); } catch (Exception error) { detach = error; RetainHost(host); }
            if (primary != null && detach != null) throw new AggregateException("Format qualification and CLR detachment failures are both retained.", primary, detach);
            if (primary != null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(primary).Throw();
            if (detach != null) throw detach;
        }

        private static void RetainHost(ExcelVbeFixture host)
        {
            host.PreserveForDiagnosticRecovery = true;
            lock (retainedFormatHosts) if (!retainedFormatHosts.Contains(host)) retainedFormatHosts.Add(host);
        }

        /// <summary>Only native reads: no focus, input, accessibility action, COM call or bridge request.</summary>
        private IDictionary<string, object> ObserveOptionsClosureSettled(ExcelVbeFixture host, DateTime expectedStartUtc)
        {
            // Historical builds post Cancel without waiting for destruction. Observe its completion;
            // never send input or another bridge request while an owned dialog is still present.
            var timer = Stopwatch.StartNew();
            var first = ObserveOptionsClosure(host.ProcessId, expectedStartUtc);
            var last = first; int observations = 1;
            while (Equals(last["ProcessIdentityVerified"], true) && Equals(last["EnumerationSucceeded"], true) &&
                !Equals(last["OptionsDialogAbsent"], true) && timer.ElapsedMilliseconds < 2000)
            {
                System.Threading.Thread.Sleep(50);
                last = ObserveOptionsClosure(host.ProcessId, expectedStartUtc); observations++;
            }
            AttachEvidence(host, expectedStartUtc, "ClosureObservation", new
            {
                First = first,
                Last = last,
                Observations = observations,
                ElapsedMilliseconds = timer.ElapsedMilliseconds,
                NativeInput = 0,
                BridgeRequests = 0
            });
            return last;
        }

        private void VerifyExclusiveHost(ExcelVbeFixture host, DateTime startUtc)
        {
            var competitors = new List<object>();
            bool ownedAlive = false;
            foreach (var name in new[] { "EXCEL", "WINWORD", "POWERPNT", "MSACCESS", "MSPUB", "SLDWORKS" })
                foreach (var process in Process.GetProcessesByName(name))
                    using (process)
                    {
                        DateTime actualStart = process.StartTime.ToUniversalTime();
                        if (process.Id == host.ProcessId && name == "EXCEL" && actualStart == startUtc && !process.HasExited)
                            ownedAlive = true;
                        else competitors.Add(new { ProcessId = process.Id, Name = name, ProcessStartUtc = actualStart.ToString("o") });
                    }
            AttachEvidence(host, startUtc, "HostExclusivityObservation", new
            {
                OwnedAlive = ownedAlive,
                Competitors = competitors,
                BeforeDispatch = true,
                NativeInput = 0,
                BridgeRequests = 0
            });
            if (!ownedAlive || competitors.Count != 0)
                throw new InvalidOperationException("The owned Excel identity or exclusive VBE-host interval changed; retain without further native dispatch.");
        }

        private static IDictionary<string, object> ObserveOptionsClosure(int processId, DateTime expectedStartUtc)
        {
            var windows = new List<object>(); var confirmedGone = new List<long>(); bool optionsAbsent = true, identity = false, complete = false;
            string error = null;
            DateTime beforeStart = DateTime.MinValue, afterStart = DateTime.MinValue;
            try
            {
                using (var process = Process.GetProcessById(processId))
                {
                    beforeStart = process.StartTime.ToUniversalTime();
                    if (process.HasExited || beforeStart != expectedStartUtc)
                        throw new InvalidOperationException("The owned process identity changed before native window enumeration.");
                    complete = EnumWindows((handle, parameter) =>
                    {
                        try
                        {
                            uint? pid = NativeWindowEnumerationOwnership.ReadProcessId(handle,
                                GetWindowThreadProcessId, Marshal.GetLastWin32Error, IsWindow);
                            if (!pid.HasValue) { confirmedGone.Add(handle.ToInt64()); return true; }
                            if (pid.Value != processId) return true;
                            var className = new StringBuilder(256);
                            if (GetClassName(handle, className, className.Capacity) == 0)
                                throw new InvalidOperationException("An owned window class could not be read.");
                            int length = GetWindowTextLength(handle);
                            if (length < 0 || length > 8192) throw new InvalidOperationException("An owned window title cannot be bounded safely.");
                            var title = new StringBuilder(length + 1);
                            if (length > 0 && GetWindowText(handle, title, title.Capacity) != length)
                                throw new InvalidOperationException("An owned window title changed or could not be read.");
                            bool visible = IsWindowVisible(handle);
                            windows.Add(new { Handle = handle.ToInt64(), ProcessId = pid, Class = className.ToString(), Title = title.ToString(), Visible = visible });
                            // An owned dialog, including a hidden/unrecognized one, cannot prove safe Options closure.
                            if (className.ToString() == "#32770")
                                optionsAbsent = false;
                            return true;
                        }
                        catch (Exception failure) { error = failure.ToString(); return false; }
                    }, IntPtr.Zero);
                    process.Refresh(); afterStart = process.StartTime.ToUniversalTime();
                    identity = !process.HasExited && beforeStart == expectedStartUtc && afterStart == expectedStartUtc;
                }
            }
            catch (Exception failure) { error = failure.ToString(); }
            return new Dictionary<string, object>
            {
                ["ProcessId"] = processId,
                ["ExpectedStartUtc"] = expectedStartUtc.ToString("o"),
                ["BeforeStartUtc"] = beforeStart.ToString("o"),
                ["AfterStartUtc"] = afterStart.ToString("o"),
                ["ProcessIdentityVerified"] = identity,
                ["EnumerationSucceeded"] = complete && error == null,
                ["OptionsDialogAbsent"] = complete && identity && error == null && optionsAbsent,
                ["OwnedWindows"] = windows.ToArray(),
                ["ConfirmedGoneDuringEnumeration"] = confirmedGone.ToArray(),
                ["Error"] = error,
                ["ObservationOnly"] = true
            };
        }

        private void AttachEvidence(ExcelVbeFixture host, DateTime startUtc, string phase, object evidence)
        {
            int sequence = ++evidenceSequence;
            string path = Path.Combine(evidenceDirectory, "options-" + sequence.ToString("D4") + "-" + phase + "-" + host.ProcessId + ".json");
            var record = new
            {
                Phase = phase,
                Sequence = sequence,
                ObservedUtc = DateTime.UtcNow.ToString("o"),
                host.ProcessId,
                ProcessStartUtc = startUtc.ToString("o"),
                FixtureRoot = host.Root,
                StartupEvidence = host.File("startup.json"),
                OwnedBootstrapEvidence = host.File("owned-bootstrap.json"),
                PhaseTrace = Path.Combine(evidenceDirectory, "owned-bootstrap-phases.jsonl"),
                LaunchContext = "ExplicitXAutomation",
                PathVisibilityManifestSupplied = false,
                ProductMvid = typeof(VbeSession).Module.ModuleVersionId.ToString("D"),
                TestMvid = typeof(ExcelFormatOptionsTests).Module.ModuleVersionId.ToString("D"),
                EvidenceRoot = evidenceDirectory,
                Data = evidence
            };
            File.WriteAllText(path, new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 }.Serialize(record), new UTF8Encoding(false));
            try { TestContext.AddResultFile(path); }
            catch (Exception error) { TestContext.WriteLine("Evidence retained at " + path + "; result attachment failed: " + error); }
        }
    }
}
