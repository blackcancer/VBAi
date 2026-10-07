namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Threading;
    using VBAi;

    public sealed partial class VbeProjectProtectionTests
    {
        private sealed class Probe : VbeDebugWindows.IProjectProtectionProbe
        {
            internal bool Open = true, Locked, FailCapture, FailWrite, FailAccept, IgnoreWrite, ChangeIdentity, KeepOpen;
            internal int PasswordLength, ConfirmationLength, Writes, Cancels, Accepts;
            internal string Secret;
            internal bool KeepCancelOpen, FailDialog, ChangeNativeIdentity, MismatchConfirmation;
            public IntPtr Dialog(string projectName) { if (FailDialog) throw new InvalidOperationException("dialog unavailable"); return Open ? new IntPtr(71) : IntPtr.Zero; }
            public VbeDebugWindows.ProjectProtectionState Capture(IntPtr dialog, string projectName)
            {
                if (FailCapture) throw new InvalidOperationException(Secret);
                return new VbeDebugWindows.ProjectProtectionState
                {
                    Identity = ChangeIdentity && Writes > 0 ? "other" : "exact",
                    NativeIdentity = ChangeNativeIdentity && Writes > 0 ? "changed" : "native",
                    Locked = Locked,
                    PasswordLength = PasswordLength,
                    ConfirmationLength = MismatchConfirmation && Writes > 0 ? ConfirmationLength + 1 : ConfirmationLength
                };
            }
            public void Write(IntPtr dialog, string projectName, bool locked, string password)
            {
                Writes++; Secret = password;
                if (FailWrite) throw new InvalidOperationException(password);
                if (IgnoreWrite) return;
                Locked = locked; PasswordLength = password.Length; ConfirmationLength = password.Length;
            }
            public void Accept(IntPtr dialog)
            {
                Accepts++;
                if (FailAccept) throw new InvalidOperationException(Secret);
                if (!KeepOpen) Open = false;
            }
            public void Cancel(IntPtr dialog) { Cancels++; if (!KeepCancelOpen) Open = false; }
            public void Pause(int milliseconds) { }
        }
        private static Request Request(Probe native)
        {
            dynamic read = VbeDebugWindows.ReadProjectProtection(new Request { Project = @"C:\fixture\macro.swp", Caption = "ExactProject" }, native);
            native.Open = true; native.Cancels = 0;
            return new Request
            {
                Project = @"C:\fixture\macro.swp",
                Caption = "ExactProject",
                Action = "clear",
                ExpectedOptionsVersion = read.OptionsVersion
            };
        }

        public sealed class Project
        {
            public string Name { get; set; } = "ExactProject";
            public string FileName { get; set; } = @"C:\fixture\macro.swp";
            public int Mode { get; set; } = 2;
            public int Protection { get; set; }
        }
        public sealed class Vbe
        {
            public List<Project> VBProjects { get; } = new List<Project>();
            public Project ActiveVBProject { get; set; }
            public List<VbeDebugTests.FakeBar> CommandBars { get; } = new List<VbeDebugTests.FakeBar>();
        }
        private sealed class QueueContext : SynchronizationContext
        {
            internal SendOrPostCallback Callback;
            public override void Post(SendOrPostCallback callback, object state) { Callback = callback; }
            internal void Run() { Callback(null); }
        }
    }
}
