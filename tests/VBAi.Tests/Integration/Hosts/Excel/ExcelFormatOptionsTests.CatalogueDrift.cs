using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration.Hosts.Excel
{
    public sealed partial class ExcelFormatOptionsTests
    {
        /// <summary>Controlled metadata drift; never presents an induced refusal as the original historical failure.</summary>
        [STATestMethod]
        public void NativeCatalogueFocusRevisionWithoutPreferenceWrites()
        {
            const string inertProperty = "__Q026_CATALOGUE_DRIFT_NO_WRITE__";
            if (Environment.GetEnvironmentVariable("VBAi_RUN_EXCEL_TESTS") != "1") Assert.Inconclusive("Excel automation is opt-in.");
            string desktop = Environment.GetEnvironmentVariable("VBAi_TEST_DESKTOP_NAME");
            IsolatedTestDesktop.RequireCurrent(desktop);
            string output = Environment.GetEnvironmentVariable("VBAi_TEST_FORMAT_OPTIONS_OUTPUT");
            Assert.IsTrue(Path.IsPathRooted(output));
            evidenceDirectory = Path.Combine(output, "options-evidence-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(evidenceDirectory);
            string bootstrap = Path.Combine(evidenceDirectory, "owned-bootstrap-phases.jsonl");
            Environment.SetEnvironmentVariable(VbeInspectionTrace.EnvironmentName, bootstrap);
            ExcelVbeFixture host = ExcelVbeFixture.StartOwnedWithTrace(bootstrap);
            DateTime start; using (var process = Process.GetProcessById(host.ProcessId)) start = process.StartTime.ToUniversalTime();
            object recording = new object();
            Action<string, object> record = (phase, value) => { lock (recording) AttachEvidence(host, start, phase, value); };
            Action guard = () => { lock (recording) VerifyExclusiveHost(host, start); };
            var intent = new ManualResetEventSlim(); var done = new ManualResetEventSlim();
            Q026OptionsGuardTrace trace = null;
            Exception actorError = null;
            object request = null;
            IDictionary<string, object> response = null;
            bool actorStarted = false;
            try
            {
                bool historical = Q026OptionsGuardTrace.NeedsGuardWarmupIfRequested();
                var warmup = new ExcelFormatOptionsQualification(host.ProcessId, host.Command,
                    () => ObserveOptionsClosureSettled(host, start), () => RetainHost(host), () => { }, record,
                    verifyExclusiveHost: guard);
                if (historical) warmup.WarmGuardForBreakpoint();
                trace = Q026OptionsGuardTrace.StartIfRequested(host, start, evidenceDirectory);
                guard(); var baseline = ReadSizeFocusSnapshot(host); record("BaselineComplete", baseline);
                var format = ((object[])baseline["Tabs"]).Select(VbeBridgeClient.Object)
                    .Single(tab => ((object[])tab["Controls"]).Select(VbeBridgeClient.Object)
                        .Any(c => Equals(c["Name"], "Size") || Equals(c["Name"], "Taille :")));
                var size = ((object[])format["Controls"]).Select(VbeBridgeClient.Object)
                    .Single(c => (Equals(c["Name"], "Size") || Equals(c["Name"], "Taille :")) && Equals(c["Type"], "ControlType.ComboBox"));
                Assert.IsFalse(((object[])baseline["Tabs"]).Select(VbeBridgeClient.Object)
                    .SelectMany(tab => ((object[])tab["Controls"]).Select(VbeBridgeClient.Object)).Any(c => Equals(c["Name"], inertProperty)));
                if (historical) Assert.AreEqual(0, ((object[])size["Choices"]).Length);
                else Assert.IsTrue(((object[])size["Choices"]).Length > 0, "Current candidate must prepare Size before hashing.");
                record("CatalogueDriftScenario", new { Historical = historical, ControlledFocus = true,
                    NaturalHistoricalRefusal = false, OriginalHistoricalCauseProven = false, PreferenceWriteTargetExists = false });
                var actor = new Thread(() => {
                    IntPtr lease = IntPtr.Zero;
                    try
                    {
                        lease = SizeFocusNative.OpenDesktopW(desktop, 0, false, 0xC7);
                        if (lease == IntPtr.Zero || !SizeFocusNative.SetThreadDesktop(lease)) throw new Win32Exception(System.Runtime.InteropServices.Marshal.GetLastWin32Error());
                        IsolatedTestDesktop.RequireCurrent(desktop);
                        if (!intent.Wait(TimeSpan.FromSeconds(15))) throw new TimeoutException("Inert command intent was not emitted.");
                        IntPtr dialog = SizeFocusNative.WaitDialog(host.ProcessId, guard);
                        var timer = Stopwatch.StartNew(); IntPtr target;
                        // Observe the product's own tab navigation. Do not select a tab or race its tab enumeration.
                        do { target = SizeFocusNative.TryFindVisibleSize(dialog, host.ProcessId); if (target != IntPtr.Zero) break; Thread.Sleep(2); }
                        while (timer.ElapsedMilliseconds < 45000);
                        Assert.AreNotEqual(IntPtr.Zero, target, "The actual visible Size control must be observed once.");
                        IntPtr parent = SizeFocusNative.GetParent(target);
                        uint owner = SizeFocusNative.RequireOwned(target, host.ProcessId);
                        Assert.AreEqual("#32770", SizeFocusNative.Class(parent)); SizeFocusNative.RequireOwned(parent, host.ProcessId);
                        string before = SizeFocusNative.Text(target); int countBefore = SizeFocusNative.Count(target);
                        guard(); record("CatalogueDriftFocusIntent", new { Dialog = dialog.ToInt64(), Target = target.ToInt64(),
                            Parent = parent.ToInt64(), CountBefore = countBefore, ValueBefore = before, Attempts = 1,
                            Message = "WM_NEXTDLGCTL", PreferenceWrites = 0, KeyboardInput = 0 });
                        SizeFocusNative.RequireOwned(parent, host.ProcessId); SizeFocusNative.RequireOwned(target, host.ProcessId);
                        Assert.AreEqual(parent, SizeFocusNative.GetParent(target));
                        if (!SizeFocusNative.PostMessage(parent, 0x28, target, new IntPtr(1))) throw new Win32Exception(System.Runtime.InteropServices.Marshal.GetLastWin32Error());
                        timer.Restart(); IntPtr focus;
                        do { focus = SizeFocusNative.Focus(owner); if (focus == target || SizeFocusNative.GetParent(focus) == target) break; Thread.Sleep(2); }
                        while (timer.ElapsedMilliseconds < 3000);
                        Assert.IsTrue(focus == target || SizeFocusNative.GetParent(focus) == target, "One posted focus must be independently observed; no repost.");
                        int countAfter = SizeFocusNative.Count(target); string after = SizeFocusNative.Text(target);
                        Assert.AreEqual(before, after); Assert.AreEqual(size["Value"], after); Assert.IsTrue(countAfter > 0);
                        record("CatalogueDriftFocusVerified", new { CountBefore = countBefore, CountAfter = countAfter,
                            ValueBefore = before, ValueAfter = after, FocusHwnd = focus.ToInt64(), Attempts = 1, PreferenceWrites = 0, KeyboardInput = 0 });
                    }
                    catch (Exception error) { actorError = error; record("CatalogueDriftActorFailed", new { Error = error.ToString(), ReplayAllowed = false }); }
                    finally { if (lease != IntPtr.Zero) SizeFocusNative.CloseDesktop(lease); done.Set(); }
                }) { IsBackground = true };
                actor.SetApartmentState(ApartmentState.MTA); actor.Start(); actorStarted = true;
                request = new { Command = "set_vbe_option", Pane = format["Tab"], Property = inertProperty,
                    Value = "No preference mutation", ExpectedOptionsVersion = baseline["OptionsVersion"] };
                guard(); record("CatalogueDriftIntent", request); intent.Set(); response = host.Command(request);
                record("CatalogueDriftReply", response);
                Assert.IsTrue(done.Wait(TimeSpan.FromSeconds(10)), "Actor must be terminal before another command or cleanup.");
                if (actorError != null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(actorError).Throw();
                Assert.IsNotNull(response); Assert.AreEqual(false, response["Ok"]); Assert.IsNull(response["Data"]);
                Assert.AreEqual(historical ? "VBE options changed since inspection; read them again." :
                    "The exact option is absent, ambiguous or unreadable.", response["Error"]);
                var closure = ObserveOptionsClosureSettled(host, start);
                Assert.AreEqual(true, closure["ProcessIdentityVerified"]); Assert.AreEqual(true, closure["EnumerationSucceeded"]);
                Assert.AreEqual(true, closure["OptionsDialogAbsent"]); record("CatalogueDriftClosureVerified", closure);
                guard(); var restored = ReadSizeFocusSnapshot(host);
                var json = new JavaScriptSerializer();
                Assert.AreEqual(baseline["OptionsVersion"], restored["OptionsVersion"]);
                Assert.AreEqual(json.Serialize(baseline["Tabs"]), json.Serialize(restored["Tabs"]));
                record("CatalogueDriftBaselineUnchanged", new { Baseline = baseline, Readback = restored,
                    PreferenceWrites = 0, ControlledGuardCaptureRequiresOfflineVerification = historical });
                trace?.Dispose(); host.Dispose(); record("ShutdownVerified", host.ShutdownDiagnostics);
                Assert.AreEqual(true, host.ShutdownDiagnostics["Exited"]); Assert.AreEqual(0, Convert.ToInt32(host.ShutdownDiagnostics["ExitCode"]));
                Assert.AreEqual(false, host.ShutdownDiagnostics["ForcedTermination"]);
                record("CatalogueDriftTerminal", new { Verified = true, ControlledFocus = true,
                    NativeReplayAllowed = false, OriginalHistoricalCauseProven = false });
            }
            catch (Exception primary)
            {
                RetainHost(host);
                record("HostRetained", new { Error = primary.ToString(), Request = request, Response = response,
                    CommittedRestoreEntries = new object[0], ActorStarted = actorStarted, ActorTerminal = done.IsSet,
                    NativeReplayAllowed = false, CleanupAllowed = false, PreferenceWriteTargetExists = false });
                try { trace?.Dispose(); } catch (Exception detach) { throw new AggregateException("Controlled catalogue diagnosis and detachment both failed; retained.", primary, detach); }
                throw;
            }
            finally { if (!actorStarted || done.IsSet) { intent.Dispose(); done.Dispose(); } }
        }
    }
}
