using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Automation;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Integration.Hosts.Excel
{
    public sealed partial class ExcelFormatOptionsTests
    {
        [STATestMethod]
        public void NativeSizeFocusCatalogueWithoutPreferenceWrites()
        {
            if (Environment.GetEnvironmentVariable("VBAi_RUN_EXCEL_TESTS") != "1")
                Assert.Inconclusive("Excel automation is opt-in.");
            string desktop = Environment.GetEnvironmentVariable("VBAi_TEST_DESKTOP_NAME");
            IsolatedTestDesktop.RequireCurrent(desktop);
            string output = Environment.GetEnvironmentVariable("VBAi_TEST_FORMAT_OPTIONS_OUTPUT");
            Assert.IsTrue(System.IO.Path.IsPathRooted(output));
            evidenceDirectory = System.IO.Path.Combine(output, "options-evidence-" + Guid.NewGuid().ToString("N"));
            System.IO.Directory.CreateDirectory(evidenceDirectory);
            string trace = System.IO.Path.Combine(evidenceDirectory, "owned-bootstrap-phases.jsonl");
            Environment.SetEnvironmentVariable(VbeInspectionTrace.EnvironmentName, trace);
            var host = ExcelVbeFixture.StartOwnedWithTrace(trace);
            DateTime start; using (var process = Process.GetProcessById(host.ProcessId)) start = process.StartTime.ToUniversalTime();
            object recording = new object();
            Action<string, object> record = (phase, value) => { lock (recording) AttachEvidence(host, start, phase, value); };
            Action guard = () => { lock (recording) VerifyExclusiveHost(host, start); };
            var intent = new ManualResetEventSlim(); var done = new ManualResetEventSlim();
            Exception actorError = null; bool closed = false; bool menuEmitted = false;
            try
            {
                guard();
                var baseline = ReadSizeFocusSnapshot(host);
                record("SizeFocusBaseline", baseline);
                var format = ((object[])baseline["Tabs"]).Select(VbeBridgeClient.Object)
                    .Single(x => ((object[])x["Controls"]).Select(VbeBridgeClient.Object)
                        .Any(c => Equals(c["Name"], "Size") || Equals(c["Name"], "Taille :")));
                string tab = (string)format["Tab"];
                var actor = new Thread(() => {
                    IntPtr lease = IntPtr.Zero;
                    try
                    {
                        lease = SizeFocusNative.OpenDesktopW(desktop, 0, false, 0xC7);
                        if (lease == IntPtr.Zero || !SizeFocusNative.SetThreadDesktop(lease)) throw new Win32Exception(Marshal.GetLastWin32Error());
                        IsolatedTestDesktop.RequireCurrent(desktop);
                        if (!intent.Wait(TimeSpan.FromSeconds(15))) throw new TimeoutException("Options command was not emitted.");
                        IntPtr dialog = SizeFocusNative.WaitDialog(host.ProcessId, guard);
                        record("SizeFocusDialog", new { Hwnd = dialog.ToInt64(), PreferenceWrites = 0 });
                        var tabs = AutomationElement.FromHandle(dialog).FindAll(TreeScope.Descendants,
                            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem));
                        var matching = tabs.Cast<AutomationElement>().Where(x => x.Current.Name == tab).ToArray();
                        Assert.AreEqual(1, matching.Length, "Select only the observed Format tab.");
                        guard(); ((SelectionItemPattern)matching[0].GetCurrentPattern(SelectionItemPattern.Pattern)).Select();
                        IntPtr size = SizeFocusNative.FindSize(dialog, host.ProcessId);
                        IntPtr parent = SizeFocusNative.GetParent(size);
                        uint owner = SizeFocusNative.RequireOwned(size, host.ProcessId);
                        Assert.AreEqual("#32770", SizeFocusNative.Class(parent));
                        SizeFocusNative.RequireOwned(parent, host.ProcessId);
                        string before = SizeFocusNative.Text(size);
                        int countBefore = SizeFocusNative.Count(size);
                        record("SizeFocusBefore", new { SizeHwnd = size.ToInt64(), ParentHwnd = parent.ToInt64(), OwnerThreadId = owner,
                            Value = before, Count = countBefore, FocusHwnd = SizeFocusNative.Focus(owner).ToInt64() });
                        guard();
                        record("SizeFocusIntent", new { Message = "WM_NEXTDLGCTL", Target = size.ToInt64(), Attempts = 1 });
                        if (!SizeFocusNative.PostMessage(parent, 0x28, size, new IntPtr(1))) throw new Win32Exception(Marshal.GetLastWin32Error());
                        var timer = Stopwatch.StartNew(); IntPtr focus;
                        do { focus = SizeFocusNative.Focus(owner); if (focus == size || SizeFocusNative.GetParent(focus) == size) break; Thread.Sleep(25); }
                        while (timer.ElapsedMilliseconds < 3000);
                        Assert.IsTrue(focus == size || SizeFocusNative.GetParent(focus) == size, "Posted focus must be independently observed; no repost.");
                        int countFocused = SizeFocusNative.Count(size);
                        bool expanded = false;
                        guard();
                        if (SizeFocusNative.Send(size, 0x157, IntPtr.Zero) == IntPtr.Zero)
                        { record("SizeFocusExpansionIntent", new { Attempts = 1 }); SizeFocusNative.Send(size, 0x14F, new IntPtr(1)); expanded = true; }
                        var choices = SizeFocusNative.Choices(size);
                        string after = SizeFocusNative.Text(size);
                        if (expanded) { guard(); SizeFocusNative.Send(size, 0x14F, IntPtr.Zero); }
                        Assert.AreEqual(before, after, "Focus and expansion must preserve the exact edit value.");
                        record("SizeFocusCatalogue", new { CountBefore = countBefore, CountFocused = countFocused,
                            CountExpanded = choices.Length, Choices = choices, ValueBefore = before, ValueAfter = after,
                            FocusHwnd = focus.ToInt64(), PreferenceWrites = 0, KeyboardInput = 0, ExpansionAttempted = expanded });
                        IntPtr cancel = SizeFocusNative.GetDlgItem(dialog, 2);
                        Assert.AreNotEqual(IntPtr.Zero, cancel); Assert.AreEqual("Button", SizeFocusNative.Class(cancel));
                        SizeFocusNative.RequireOwned(cancel, host.ProcessId);
                        Assert.IsTrue(SizeFocusNative.IsWindowEnabled(cancel));
                        guard(); record("SizeFocusCancelIntent", new { Dialog = dialog.ToInt64(), Cancel = cancel.ToInt64(), Attempts = 1 });
                        if (!SizeFocusNative.PostMessage(dialog, 0x111, new IntPtr(2), cancel)) throw new Win32Exception(Marshal.GetLastWin32Error());
                        timer.Restart(); while (SizeFocusNative.IsWindow(dialog) && timer.ElapsedMilliseconds < 5000) Thread.Sleep(25);
                        Assert.IsFalse(SizeFocusNative.IsWindow(dialog), "The exact captured dialog must be destroyed.");
                        closed = true; record("SizeFocusDialogDestroyed", new { Dialog = dialog.ToInt64(), CancelAttempts = 1 });
                    }
                    catch (Exception error) { actorError = error; record("SizeFocusActorFailed", new { Error = error.ToString(), CleanupAllowed = false }); }
                    finally { done.Set(); if (lease != IntPtr.Zero) SizeFocusNative.CloseDesktop(lease); }
                }) { IsBackground = true };
                actor.SetApartmentState(ApartmentState.MTA); actor.Start();
                guard();
                host.ExecuteOwnedOptionsMenu(value => {
                    record("SizeFocusMenu", value);
                    if (value.GetType().GetProperty("Phase").GetValue(value).ToString() == "OptionsMenuIntent")
                    { menuEmitted = true; intent.Set(); }
                });
                Assert.IsTrue(done.Wait(TimeSpan.FromSeconds(10)), "UI observation must finish before further COM dispatch.");
                if (actorError != null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(actorError).Throw();
                Assert.IsTrue(closed && menuEmitted);
                guard(); var restored = ReadSizeFocusSnapshot(host);
                var json = new System.Web.Script.Serialization.JavaScriptSerializer();
                Assert.AreEqual(baseline["OptionsVersion"], restored["OptionsVersion"]);
                Assert.AreEqual(json.Serialize(baseline["Tabs"]), json.Serialize(restored["Tabs"]));
                record("SizeFocusBaselineUnchanged", new { Baseline = baseline, Readback = restored, PreferenceWrites = 0 });
                host.Dispose(); record("ShutdownVerified", host.ShutdownDiagnostics);
                Assert.AreEqual(true, host.ShutdownDiagnostics["Exited"]);
                Assert.AreEqual(0, Convert.ToInt32(host.ShutdownDiagnostics["ExitCode"]));
                Assert.AreEqual(false, host.ShutdownDiagnostics["ForcedTermination"]);
                record("SizeFocusTerminal", new { Verified = true, SizeMutationAccepted = false, NativeReplayAllowed = false });
            }
            catch (Exception error)
            {
                RetainHost(host);
                record("SizeFocusRetained", new { Error = error.ToString(), menuEmitted, closed, ActorTerminal = done.IsSet,
                    ActorError = actorError?.ToString(), CleanupReplayed = false });
                throw;
            }
        }

        private static IDictionary<string, object> ReadSizeFocusSnapshot(ExcelVbeFixture host)
        {
            var reply = host.Command("vbe_options"); Assert.IsNotNull(reply); Assert.AreEqual(true, reply["Ok"]);
            var data = VbeBridgeClient.Object(reply["Data"]); Assert.AreEqual(true, data["DialogClosed"]);
            Assert.IsTrue(((object[])data["Tabs"]).Length > 0); Assert.AreEqual(64, ((string)data["OptionsVersion"]).Length);
            return data;
        }

        private static class SizeFocusNative
        {
            [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)] internal static extern IntPtr OpenDesktopW(string name, uint flags, bool inherit, uint access);
            [DllImport("user32.dll", SetLastError = true)] internal static extern bool SetThreadDesktop(IntPtr desktop);
            [DllImport("user32.dll")] internal static extern bool CloseDesktop(IntPtr desktop);
            [DllImport("user32.dll")] internal static extern IntPtr GetParent(IntPtr window);
            [DllImport("user32.dll")] internal static extern IntPtr GetDlgItem(IntPtr window, int id);
            [DllImport("user32.dll")] internal static extern bool IsWindow(IntPtr window);
            [DllImport("user32.dll")] internal static extern bool IsWindowEnabled(IntPtr window);
            [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
            [DllImport("user32.dll")] private static extern int GetDlgCtrlID(IntPtr window);
            [DllImport("user32.dll", SetLastError = true)] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
            private delegate bool Visitor(IntPtr window, IntPtr state);
            [DllImport("user32.dll", SetLastError = true)] private static extern bool EnumWindows(Visitor visitor, IntPtr state);
            [DllImport("user32.dll", SetLastError = true)] private static extern bool EnumChildWindows(IntPtr parent, Visitor visitor, IntPtr state);
            [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder text, int length);
            [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern IntPtr SendMessageTimeout(IntPtr window, uint msg, IntPtr wp, IntPtr lp, uint flags, uint timeout, out IntPtr result);
            [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "SendMessageTimeoutW", SetLastError = true)] private static extern IntPtr SendText(IntPtr window, uint msg, IntPtr wp, StringBuilder text, uint flags, uint timeout, out IntPtr result);
            [DllImport("user32.dll", SetLastError = true)] internal static extern bool PostMessage(IntPtr window, uint msg, IntPtr wp, IntPtr lp);
            [DllImport("user32.dll", SetLastError = true)] private static extern bool GetGUIThreadInfo(uint thread, ref GuiInfo info);
            [StructLayout(LayoutKind.Sequential)] private struct GuiInfo
            { internal uint Size, Flags; internal IntPtr Active, Focus, Capture, MenuOwner, MoveSize, Caret; internal int Left, Top, Right, Bottom; }
            internal static IntPtr Focus(uint owner) { var info = new GuiInfo { Size = (uint)Marshal.SizeOf(typeof(GuiInfo)) }; if (!GetGUIThreadInfo(owner, ref info)) throw new Win32Exception(Marshal.GetLastWin32Error()); return info.Focus; }
            internal static uint RequireOwned(IntPtr window, int expected) { uint pid; uint tid = GetWindowThreadProcessId(window, out pid); Assert.IsTrue(tid != 0 && pid == expected && IsWindow(window)); return tid; }
            internal static string Class(IntPtr window) { var text = new StringBuilder(128); Assert.IsTrue(GetClassName(window, text, text.Capacity) > 0); return text.ToString(); }
            internal static IntPtr Send(IntPtr window, uint msg, IntPtr wp) { IntPtr result; if (SendMessageTimeout(window, msg, wp, IntPtr.Zero, 2, 2000, out result) == IntPtr.Zero) throw new Win32Exception(Marshal.GetLastWin32Error()); return result; }
            internal static string Text(IntPtr window) { int length = Send(window, 0xE, IntPtr.Zero).ToInt32(); Assert.IsTrue(length >= 0 && length < 4096); var text = new StringBuilder(length + 1); IntPtr result; if (SendText(window, 0xD, new IntPtr(text.Capacity), text, 2, 2000, out result) == IntPtr.Zero) throw new Win32Exception(Marshal.GetLastWin32Error()); Assert.AreEqual(length, result.ToInt32()); return text.ToString(); }
            internal static int Count(IntPtr window) { int count = Send(window, 0x146, IntPtr.Zero).ToInt32(); Assert.IsTrue(count >= 0 && count <= 2000); return count; }
            internal static string[] Choices(IntPtr window)
            {
                int count = Count(window); var choices = new List<string>();
                for (int i = 0; i < count; i++) { int length = Send(window, 0x149, new IntPtr(i)).ToInt32(); Assert.IsTrue(length >= 0 && length < 4096); var text = new StringBuilder(length + 1); IntPtr result; if (SendText(window, 0x148, new IntPtr(i), text, 2, 2000, out result) == IntPtr.Zero) throw new Win32Exception(Marshal.GetLastWin32Error()); Assert.AreEqual(length, result.ToInt32()); choices.Add(text.ToString()); }
                Assert.AreEqual(count, Count(window)); return choices.ToArray();
            }
            internal static IntPtr WaitDialog(int pid, Action guard)
            {
                var timer = Stopwatch.StartNew();
                do {
                    guard(); var found = new List<IntPtr>(); Exception failure = null;
                    bool ok = EnumWindows((window, _) => { try { var owner = NativeWindowEnumerationOwnership.ReadProcessId(window, GetWindowThreadProcessId, Marshal.GetLastWin32Error, IsWindow); if (owner == pid && IsWindowVisible(window) && Class(window) == "#32770" && Text(window) == "Options") found.Add(window); return true; } catch (Exception error) { failure = error; return false; } }, IntPtr.Zero);
                    if (failure != null) throw failure; Assert.IsTrue(ok); Assert.IsTrue(found.Count <= 1);
                    if (found.Count == 1) return found[0]; Thread.Sleep(50);
                } while (timer.ElapsedMilliseconds < 15000);
                throw new TimeoutException("Exactly one owned Options dialog is required.");
            }
            internal static IntPtr FindSize(IntPtr dialog, int pid)
            {
                var found = new List<IntPtr>(); Exception failure = null;
                EnumChildWindows(dialog, (window, _) => { try { if (GetDlgCtrlID(window) == 4911 && Class(window) == "ComboBox" && IsWindowVisible(window) && IsWindowEnabled(window)) { RequireOwned(window, pid); found.Add(window); } return true; } catch (Exception error) { failure = error; return false; } }, IntPtr.Zero);
                if (failure != null) throw failure; Assert.AreEqual(1, found.Count); return found[0];
            }
        }
    }
}
