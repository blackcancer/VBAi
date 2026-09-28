namespace CodexVBE.Tests.Unit
{
    using System;
    using System.IO;
    using System.Text;
    using System.Threading;
    using System.Web.Script.Serialization;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using CodexVBE;

    // Matrix before execution: exact active project, design/unprotected/version/command caption guards;
    // callback mode/version/identity/command drift; native unavailable/ambiguous controls via probe;
    // read Cancel and sanitized fields; stale options; valid lock/clear; strict UTF8 file limits;
    // native write/readback/accept/identity failures => Cancel and no secret or retry;
    // closed/pending dialog is not persistence proof and never triggers Save.
    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class VbeProjectProtectionTests
    {
        [TestMethod]
        public void ReadProtectionReturnsOnlyPresenceAndVersionThenCancels()
        {
            var native = new Probe { Locked = true, PasswordLength = 8, ConfirmationLength = 8, Secret = "TOPSECRET" };
            dynamic result = VbeDebugWindows.ReadProjectProtection(new Request { Project = "ExactProject", Caption = "ExactProject" }, native);
            Assert.IsTrue((bool)result.Available); Assert.IsTrue((bool)result.PasswordPresent);
            Assert.AreEqual(64, ((string)result.OptionsVersion).Length);
            Assert.IsTrue((bool)result.DialogClosed); Assert.AreEqual(1, native.Cancels);
            Assert.IsFalse(new JavaScriptSerializer().Serialize((object)result).Contains(native.Secret));
            native = new Probe { FailCapture = true, Secret = "TOPSECRET" };
            result = VbeDebugWindows.ReadProjectProtection(new Request { Project = "ExactProject", Caption = "ExactProject" }, native);
            Assert.IsFalse((bool)result.Available); Assert.AreEqual(1, native.Cancels);
            Assert.IsFalse(new JavaScriptSerializer().Serialize((object)result).Contains(native.Secret));
            native = new Probe { Open = false };
            result = VbeDebugWindows.ReadProjectProtection(new Request { Project = "ExactProject", Caption = "ExactProject" }, native);
            Assert.IsFalse((bool)result.Available);
        }

        [TestMethod]
        public void ClearAndLockVerifyNativeControlStateWithoutClaimingPersistence()
        {
            var native = new Probe { Locked = true, PasswordLength = 4, ConfirmationLength = 4 };
            var request = Request(native);
            dynamic clear = VbeDebugWindows.SetProjectProtection(request, native);
            Assert.IsTrue((bool)clear.CommittedRequested); Assert.IsTrue((bool)clear.Closed);
            Assert.IsFalse((bool)clear.PersistenceVerified); Assert.IsFalse(native.Locked);
            Assert.AreEqual(0, native.PasswordLength); Assert.AreEqual(1, native.Accepts);
            string path = Path.Combine(Path.GetTempPath(), "CodexVBE-protection-" + Guid.NewGuid().ToString("N") + ".txt");
            try
            {
                File.WriteAllText(path, "TOPSECRET", new UTF8Encoding(false));
                native = new Probe(); request = Request(native); request.Action = "lock"; request.Path = path;
                dynamic locked = VbeDebugWindows.SetProjectProtection(request, native);
                Assert.IsTrue(native.Locked); Assert.AreEqual(9, native.PasswordLength);
                Assert.IsTrue((bool)locked.ControlValueVerified);
                Assert.IsFalse(new JavaScriptSerializer().Serialize((object)locked).Contains("TOPSECRET"));
                Assert.AreEqual(0, native.Cancels);
            }
            finally { File.Delete(path); }
        }

        [TestMethod]
        public void SecretFileRejectsInvalidUtf8MultilineEmptyOversizeAndMissingPathsBeforeWrites()
        {
            string path = Path.Combine(Path.GetTempPath(), "CodexVBE-secret-" + Guid.NewGuid().ToString("N") + ".txt");
            try
            {
                foreach (byte[] bytes in new[] { new byte[0], new byte[] { 0xff }, Encoding.UTF8.GetBytes("one\ntwo"),
                    Encoding.UTF8.GetBytes(new string('x', 129)), new byte[513], Encoding.UTF8.GetBytes("zero\0value") })
                {
                    File.WriteAllBytes(path, bytes);
                    var native = new Probe(); var request = Request(native); request.Action = "lock"; request.Path = path;
                    dynamic result = VbeDebugWindows.SetProjectProtection(request, native);
                    Assert.IsFalse((bool)result.Available); Assert.AreEqual(0, native.Writes);
                    Assert.AreEqual(1, native.Cancels); Assert.AreEqual(0, native.Accepts);
                }
                foreach (string invalid in new[] { null, "relative.txt", path + ".absent" })
                {
                    var native = new Probe(); var request = Request(native); request.Action = "lock"; request.Path = invalid;
                    dynamic result = VbeDebugWindows.SetProjectProtection(request, native);
                    Assert.IsFalse((bool)result.Available); Assert.AreEqual(0, native.Writes);
                    Assert.AreEqual(1, native.Cancels);
                }
            }
            finally { File.Delete(path); }
        }

        [TestMethod]
        public void ChangedOptionsAndNativeFailuresCancelWithoutSecretsOrAutomaticRetry()
        {
            for (int fault = 0; fault < 6; fault++)
            {
                var native = new Probe { Locked = true, PasswordLength = 5, ConfirmationLength = 5 };
                var request = Request(native);
                if (fault == 0) request.ExpectedOptionsVersion = "stale";
                if (fault == 1) native.FailCapture = true;
                if (fault == 2) native.FailWrite = true;
                if (fault == 3) native.IgnoreWrite = true;
                if (fault == 4) native.ChangeIdentity = true;
                if (fault == 5) native.FailAccept = true;
                dynamic result = VbeDebugWindows.SetProjectProtection(request, native);
                Assert.IsFalse((bool)result.Available);
                Assert.AreEqual(fault < 2 ? 0 : 1, native.Writes);
                Assert.AreEqual(1, native.Cancels);
                Assert.AreEqual(fault == 5 ? 1 : 0, native.Accepts);
                Assert.IsFalse((bool)result.RetryAllowed);
            }
            var pending = new Probe { KeepOpen = true }; var pendingRequest = Request(pending);
            dynamic pendingResult = VbeDebugWindows.SetProjectProtection(pendingRequest, pending);
            Assert.IsFalse((bool)pendingResult.Closed);
            Assert.IsFalse((bool)pendingResult.PersistenceVerified);
            Assert.AreEqual(1, pending.Accepts); Assert.AreEqual(1, pending.Cancels);
        }

        [TestMethod]
        public void NativeProtectionDialogRequiresExactProjectTitleClassProcessAndUniqueness()
        {
            var enumBefore = VbeDebugWindows.EnumWindows;
            var pidBefore = VbeDebugWindows.GetWindowThreadProcessId;
            var classBefore = VbeDebugWindows.GetClassName;
            var titleBefore = VbeDebugWindows.GetWindowText;
            var visibleBefore = VbeDebugWindows.IsWindowVisible;
            try
            {
                var type = typeof(VbeDebugWindows).GetNestedType("NativeProjectProtectionProbe", System.Reflection.BindingFlags.NonPublic);
                var native = (VbeDebugWindows.IProjectProtectionProbe)Activator.CreateInstance(type, true);
                for (int fault = 0; fault < 7; fault++)
                {
                    uint current = (uint)System.Diagnostics.Process.GetCurrentProcess().Id;
                    VbeDebugWindows.EnumWindows = (callback, data) =>
                    {
                        callback(new IntPtr(71), data);
                        if (fault == 6) callback(new IntPtr(72), data);
                        return true;
                    };
                    VbeDebugWindows.GetWindowThreadProcessId = (IntPtr window, out uint pid) =>
                    { pid = fault == 2 ? current + 1 : current; return 1; };
                    VbeDebugWindows.GetClassName = (window, buffer, length) =>
                    { buffer.Append(fault == 3 ? "Other" : "#32770"); return buffer.Length; };
                    VbeDebugWindows.GetWindowText = (window, buffer, length) =>
                    {
                        buffer.Append(fault == 1 ? "ExactProject - Propriétés du projet" :
                            fault == 4 ? "OtherProject - Project Properties" : "ExactProject - Project Properties");
                        return buffer.Length;
                    };
                    VbeDebugWindows.IsWindowVisible = window => fault != 5;
                    if (fault == 6) Assert.ThrowsException<InvalidOperationException>(() => native.Dialog("ExactProject"));
                    else Assert.AreEqual(fault <= 1 ? new IntPtr(71) : IntPtr.Zero, native.Dialog("ExactProject"));
                }
            }
            finally
            {
                VbeDebugWindows.EnumWindows = enumBefore;
                VbeDebugWindows.GetWindowThreadProcessId = pidBefore;
                VbeDebugWindows.GetClassName = classBefore;
                VbeDebugWindows.GetWindowText = titleBefore;
                VbeDebugWindows.IsWindowVisible = visibleBefore;
            }
        }

        // Additional native matrix: same-process TCITEMW buffer; exact/duplicate/missing tab title;
        // PSM_SETCURSEL and current-index readback; PID/style/ID/control-count guards;
        // checkbox and password length readback; ignored native click/text replacement.
        [TestMethod]
        public void Win32ProtectionReadsExactTabAndNativeControlsWithoutUiAutomation()
        {
            using (var f = new Win32ProtectionScope())
            {
                var state = f.Native.Capture(new IntPtr(71), "ExactProject");
                Assert.AreEqual(1, f.Selected); Assert.AreEqual(2, f.TabReads);
                Assert.IsFalse(state.Locked); Assert.AreEqual(0, state.PasswordLength);
                f.Native.Write(new IntPtr(71), "ExactProject", true, "fixture");
                state = f.Native.Capture(new IntPtr(71), "ExactProject");
                Assert.IsTrue(state.Locked); Assert.AreEqual(7, state.PasswordLength);
                Assert.AreEqual(7, state.ConfirmationLength); Assert.AreEqual(2, f.Writes);
                f.Native.Write(new IntPtr(71), "ExactProject", false, "");
                state = f.Native.Capture(new IntPtr(71), "ExactProject");
                Assert.IsFalse(state.Locked); Assert.AreEqual(0, state.PasswordLength);
                f.Native.Accept(new IntPtr(71)); f.Native.Cancel(new IntPtr(71));
                Assert.AreEqual(2, f.Posts);
            }
        }

        [TestMethod]
        public void Win32ProtectionRejectsTabOwnerStyleAndControlIdentityFaultsBeforeSecrets()
        {
            for (int fault = 0; fault < 9; fault++)
                using (var f = new Win32ProtectionScope())
                {
                    f.DuplicateTab = fault == 0; f.FailTabRead = fault == 1; f.FailSelect = fault == 2;
                    f.WrongOwner = fault == 3; f.WrongCheckbox = fault == 4;
                    f.MissingPassword = fault == 5; f.WrongStyle = fault == 6;
                    if (fault == 7) f.Locked = 2;
                    if (fault == 8) f.Length1 = 129;
                    Assert.ThrowsException<InvalidOperationException>(() => f.Native.Capture(new IntPtr(71), "ExactProject"));
                    Assert.AreEqual(0, f.Writes);
                }
        }

        [TestMethod]
        public void Win32ProtectionRejectsIgnoredNativeCheckboxAndPasswordWrites()
        {
            using (var f = new Win32ProtectionScope())
            {
                f.IgnoreClick = true;
                Assert.ThrowsException<InvalidOperationException>(() => f.Native.Write(new IntPtr(71), "ExactProject", true, "fixture"));
                Assert.AreEqual(0, f.Writes);
            }
            using (var f = new Win32ProtectionScope())
            {
                f.IgnorePassword = true;
                Assert.ThrowsException<InvalidOperationException>(() => f.Native.Write(new IntPtr(71), "ExactProject", true, "fixture"));
                Assert.AreEqual(1, f.Writes);
            }
        }
    }
}
