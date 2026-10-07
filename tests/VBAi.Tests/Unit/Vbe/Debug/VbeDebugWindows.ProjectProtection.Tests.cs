namespace VBAi.Tests.Unit
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.IO;
    using System.Text;
    using System.Web.Script.Serialization;
    using VBAi;

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
        public void ProtectionRequestDialogAndReadbackTransitionMatrixIsComplete()
        {
            foreach (Request invalid in new[] { null, new Request { Caption = "ExactProject" }, new Request { Project = "P" } })
                Assert.ThrowsException<ArgumentException>(() => VbeDebugWindows.ReadProjectProtection(invalid, new Probe()));
            var readRequest = new Request { Project = "P", Caption = "ExactProject" };
            var pending = new Probe { KeepCancelOpen = true }; dynamic read = VbeDebugWindows.ReadProjectProtection(readRequest, pending);
            Assert.IsTrue((bool)read.Available); Assert.IsFalse((bool)read.DialogClosed);
            Assert.IsFalse((bool)((dynamic)VbeDebugWindows.ReadProjectProtection(readRequest, new Probe { FailDialog = true })).Available);
            Assert.IsFalse((bool)((dynamic)VbeDebugWindows.SetProjectProtection(readRequest, new Probe { Open = false })).Available);
            foreach (string fault in new[] { "action", "version", "native identity", "confirmation" })
            {
                var native = new Probe(); var request = Request(native);
                if (fault == "action") request.Action = "unknown";
                if (fault == "version") request.ExpectedOptionsVersion = null;
                if (fault == "native identity") native.ChangeNativeIdentity = true;
                if (fault == "confirmation") native.MismatchConfirmation = true;
                dynamic result = VbeDebugWindows.SetProjectProtection(request, native);
                Assert.IsFalse((bool)result.Available, fault); Assert.AreEqual(0, native.Accepts, fault);
            }
        }

        private sealed class TruncatedSecretStream : MemoryStream
        {
            internal TruncatedSecretStream() : base(new byte[] { 65 }) { }
            public override long Length => 2;
        }

        [TestMethod]
        public void ProtectionSecretReadRefusesTruncatedStreamAndBomOnlyContent()
        {
            var open = VbeDebugWindows.OpenProtectionSecret;
            try
            {
                foreach (bool truncated in new[] { true, false })
                {
                    VbeDebugWindows.OpenProtectionSecret = p => truncated ? (Stream)new TruncatedSecretStream() : new MemoryStream(new byte[] { 0xef, 0xbb, 0xbf });
                    var probe = new Probe(); var request = Request(probe); request.Action = "lock"; request.Path = @"C:\fixture\secret.txt";
                    dynamic result = VbeDebugWindows.SetProjectProtection(request, probe);
                    Assert.IsFalse((bool)result.Available); Assert.AreEqual(0, probe.Writes);
                }
            }
            finally { VbeDebugWindows.OpenProtectionSecret = open; }
        }

        [TestMethod]
        public void NativeProtectionRejectsEveryControlPageOwnerAndBoundFault()
        {
            foreach (string fault in new[] { "tab hidden", "tab disabled", "tab parent", "tab absent", "tab duplicate", "tab count zero", "tab count large", "no protection tab", "selection readback", "check hidden", "check disabled", "check style", "page class", "page hidden", "page parent", "password parent", "password duplicate", "password ids", "password owner", "negative length", "children bound" })
                using (var f = new Win32ProtectionScope())
                {
                    var visible = VbeDebugWindows.IsWindowVisible; var enabled = VbeDebugWindows.OptionsWindowEnabled; var parent = VbeDebugWindows.ProtectionWindowParent;
                    var message = VbeDebugWindows.SendMessageInt; var enumerate = VbeDebugWindows.EnumChildWindows; var ids = VbeDebugWindows.GetDlgCtrlID; var styles = VbeDebugWindows.ProtectionWindowStyle; var classes = VbeDebugWindows.GetClassName; var owners = VbeDebugWindows.GetWindowThreadProcessId;
                    if (fault == "tab hidden" || fault == "check hidden" || fault == "page hidden") VbeDebugWindows.IsWindowVisible = h => h.ToInt32() != (fault == "tab hidden" ? 73 : fault == "check hidden" ? 74 : 72) && visible(h);
                    if (fault == "tab disabled" || fault == "check disabled") VbeDebugWindows.OptionsWindowEnabled = h => h.ToInt32() != (fault == "tab disabled" ? 73 : 74) && enabled(h);
                    if (fault == "tab parent" || fault == "page parent" || fault == "password parent") VbeDebugWindows.ProtectionWindowParent = h => h.ToInt32() == (fault == "tab parent" ? 73 : fault == "page parent" ? 72 : 75) ? IntPtr.Zero : parent(h);
                    if (fault == "tab absent" || fault == "tab duplicate" || fault == "password duplicate" || fault == "children bound") VbeDebugWindows.EnumChildWindows = (h, callback, p) =>
                    {
                        if (fault == "children bound") { for (int i = 0; i < 257; i++) callback(new IntPtr(1000 + i), p); return true; }
                        bool result = enumerate(h, (child, arg) => fault == "tab absent" && child.ToInt32() == 73 ? true : callback(fault == "password duplicate" && child.ToInt32() == 76 ? new IntPtr(75) : child, arg), p);
                        if (fault == "tab duplicate") callback(new IntPtr(73), p); return result;
                    };
                    if (fault == "tab count zero" || fault == "tab count large") VbeDebugWindows.SendMessageInt = (h, m, w, l) => m == 0x1304 ? new IntPtr(fault == "tab count zero" ? 0 : 17) : message(h, m, w, l);
                    if (fault == "no protection tab") VbeDebugWindows.ProtectionTabText = (h, i) => "General";
                    if (fault == "selection readback") VbeDebugWindows.ProtectionSelectTab = (d, t, i) => true;
                    if (fault == "check style") VbeDebugWindows.ProtectionWindowStyle = h => h.ToInt32() == 74 ? 4 : styles(h);
                    if (fault == "page class") VbeDebugWindows.GetClassName = (h, b, c) => { if (h.ToInt32() != 72) return classes(h, b, c); b.Append("wrong page"); return b.Length; };
                    if (fault == "password ids") VbeDebugWindows.GetDlgCtrlID = h => h.ToInt32() == 75 || h.ToInt32() == 76 ? 5469 : ids(h);
                    if (fault == "password owner") VbeDebugWindows.GetWindowThreadProcessId = (IntPtr h, out uint pid) => { uint result = owners(h, out pid); if (h.ToInt32() == 75) pid++; return result; };
                    if (fault == "negative length") f.Length1 = -1;
                    Assert.ThrowsException<InvalidOperationException>(() => f.Native.Capture(new IntPtr(71), "ExactProject"), fault); Assert.AreEqual(0, f.Writes, fault);
                }
        }

        [TestMethod]
        public void NativeProtectionButtonsAndDisabledSecretsKeepEveryDeliveryGuard()
        {
            foreach (string fault in new[] { "not validated", "dialog replaced", "missing button", "button class", "button parent", "button id", "button hidden", "button disabled", "post refused" })
                using (var f = new Win32ProtectionScope())
                {
                    if (fault != "not validated") f.Native.Dialog("ExactProject");
                    if (fault == "dialog replaced") VbeDebugWindows.EnumWindows = (c, p) => true;
                    if (fault == "missing button") VbeDebugWindows.GetDlgItem = (h, i) => IntPtr.Zero;
                    if (fault == "button class") { var read = VbeDebugWindows.GetClassName; VbeDebugWindows.GetClassName = (h, b, c) => { if (h.ToInt32() != 77) return read(h, b, c); b.Append("Edit"); return b.Length; }; }
                    if (fault == "button parent") VbeDebugWindows.ProtectionWindowParent = h => IntPtr.Zero;
                    if (fault == "button id") VbeDebugWindows.GetDlgCtrlID = h => 999;
                    if (fault == "button hidden") VbeDebugWindows.IsWindowVisible = h => h.ToInt32() != 77;
                    if (fault == "button disabled") VbeDebugWindows.OptionsWindowEnabled = h => h.ToInt32() != 77;
                    if (fault == "post refused") VbeDebugWindows.PostMessage = (h, m, w, l) => false;
                    Assert.ThrowsException<InvalidOperationException>(() => f.Native.Accept(new IntPtr(71)), fault);
                }
            foreach (string fault in new[] { "class", "style", "disabled nonempty", "disabled old secret", "disabled empty", "owner reader" })
                using (var f = new Win32ProtectionScope())
                {
                    var type = typeof(VbeDebugWindows).GetNestedType("NativeProjectProtectionProbe", System.Reflection.BindingFlags.NonPublic);
                    var set = type.GetMethod("SetPassword", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
                    if (fault == "class") VbeDebugWindows.GetClassName = (h, b, c) => { b.Append("Button"); return b.Length; };
                    if (fault == "style") VbeDebugWindows.ProtectionWindowStyle = h => 0;
                    if (fault.StartsWith("disabled", StringComparison.Ordinal)) VbeDebugWindows.OptionsWindowEnabled = h => false;
                    if (fault == "disabled old secret") f.Length1 = 1;
                    if (fault == "owner reader") VbeDebugWindows.GetWindowThreadProcessId = (IntPtr h, out uint pid) => { pid = 0; return 0; };
                    string password = fault == "disabled nonempty" ? "value" : "";
                    if (fault == "disabled empty") { set.Invoke(null, new object[] { new IntPtr(75), password }); Assert.AreEqual(0, f.Writes); }
                    else { var error = Assert.ThrowsException<System.Reflection.TargetInvocationException>(() => set.Invoke(null, new object[] { new IntPtr(75), password })); Assert.IsInstanceOfType(error.InnerException, typeof(InvalidOperationException)); }
                }
        }

        [TestMethod]
        public void ProtectionBufferFailuresReleaseEarlierAllocationAndDefaultWrappersRetainTheirNativePaths()
        {
            foreach (int failedAllocation in new[] { 1, 2 }) using (var f = new Win32ProtectionScope())
            {
                int calls = 0; VbeDebugWindows.AllocateProtectionBuffer = size => { if (++calls == failedAllocation) throw new OutOfMemoryException("owned allocation fault"); return System.Runtime.InteropServices.Marshal.AllocHGlobal(size); };
                Assert.ThrowsException<OutOfMemoryException>(() => VbeDebugWindows.ProtectionTabText(new IntPtr(73), 0));
                Assert.AreEqual(failedAllocation, calls);
            }
            using (var f = new Win32ProtectionScope())
            {
                VbeDebugWindows.PauseNative = milliseconds => { };
                var request = new Request { Project = "P", Caption = "ExactProject" };
                dynamic read = VbeDebugWindows.ReadProjectProtection(request); Assert.IsTrue((bool)read.Available); Assert.IsFalse((bool)read.DialogClosed);
                request.Action = "clear"; request.ExpectedOptionsVersion = read.OptionsVersion;
                dynamic changed = VbeDebugWindows.SetProjectProtection(request); Assert.IsTrue((bool)changed.ControlValueVerified);
                f.Selected = 1; f.Locked = 1; f.Native.Write(new IntPtr(71), "ExactProject", true, ""); Assert.AreEqual(1, f.Locked);
            }
        }

        [TestMethod]
        public void ProtectionNativeReadbackRejectsIgnoredClearWrongDialogAndOversizedTabCaption()
        {
            using (var f = new Win32ProtectionScope())
            {
                f.Locked = 1; f.IgnoreClick = true;
                Assert.ThrowsException<InvalidOperationException>(() => f.Native.Write(new IntPtr(71), "ExactProject", false, ""));
                Assert.ThrowsException<InvalidOperationException>(() => f.Native.Capture(new IntPtr(72), "ExactProject"));
            }
            foreach (bool oversized in new[] { false, true }) using (var f = new Win32ProtectionScope())
            {
                IntPtr replacement = oversized ? System.Runtime.InteropServices.Marshal.StringToHGlobalUni(new string('x', 256)) : IntPtr.Zero;
                try
                {
                    VbeDebugWindows.SendMessageInt = (h, m, w, l) =>
                    {
                        Assert.AreEqual(0x133c, m);
                        var type = typeof(VbeDebugWindows).GetNestedType("ProtectionTabItem", System.Reflection.BindingFlags.NonPublic);
                        object descriptor = System.Runtime.InteropServices.Marshal.PtrToStructure(l, type);
                        type.GetField("Text").SetValue(descriptor, replacement);
                        System.Runtime.InteropServices.Marshal.StructureToPtr(descriptor, l, false); return new IntPtr(1);
                    };
                    if (oversized) Assert.ThrowsException<InvalidOperationException>(() => VbeDebugWindows.ProtectionTabText(new IntPtr(73), 0));
                    else Assert.AreEqual("", VbeDebugWindows.ProtectionTabText(new IntPtr(73), 0));
                }
                finally { if (replacement != IntPtr.Zero) System.Runtime.InteropServices.Marshal.FreeHGlobal(replacement); }
            }
        }

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
            string path = Path.Combine(Path.GetTempPath(), "VBAi-protection-" + Guid.NewGuid().ToString("N") + ".txt");
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
            string path = Path.Combine(Path.GetTempPath(), "VBAi-secret-" + Guid.NewGuid().ToString("N") + ".txt");
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
