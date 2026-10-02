using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit")]
    public sealed class VbaTestPowerPointValuesHostTests
    {
        [TestMethod]
        public void RunPassesOneExactArgumentArrayAndReturnsTheNativeVariant()
        {
            using (var fixture = new Fixture())
            {
                object[] arguments = { "Tests", "Alpha" }; object returned = new object();
                fixture.Host.RunProcedure = (application, macro, array) => {
                    Assert.AreSame(fixture.Application, application);
                    Assert.AreEqual("Original.pptm!Support.Run", macro);
                    CollectionAssert.AreEqual(arguments, array);
                    Assert.AreNotSame(arguments, array); fixture.Invocations++; return returned;
                };
                object target = fixture.Host.ResolveTarget(fixture.Source.VBProject, fixture.Source.FullName);
                Assert.AreSame(returned, fixture.Host.Invoke(target, "Support", "Run", arguments));
                Assert.AreEqual(1, fixture.Invocations);
            }
        }

        [TestMethod]
        public void OwnerProcessAndWindowPidAreRequiredWithoutApplicationFallback()
        {
            using (var fixture = new Fixture())
            {
                fixture.Host.ReadProcessName = () => "EXCEL";
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Resolve());
                Assert.AreEqual(0, fixture.ApplicationReads);
                fixture.Host.ReadProcessName = () => "POWERPNT";
                fixture.Host.ReadWindowOwner = window => 777;
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Resolve());
                Assert.AreEqual(1, fixture.ApplicationReads);
                Assert.AreEqual(0, fixture.Invocations);
            }
        }

        [TestMethod]
        public void ForceDisableIsAConfirmedRefusalBeforeTargetResolutionOrNativeInvocation()
        {
            using (var fixture = new Fixture())
            {
                fixture.Application.AutomationSecurity = 3;
                var refusal = Assert.ThrowsException<InvalidOperationException>(() => fixture.Resolve());
                StringAssert.Contains(refusal.Message, "ForceDisable");
                Assert.AreEqual(3, fixture.Application.AutomationSecurity);
                Assert.AreEqual(0, fixture.Invocations);
                fixture.Application.AutomationSecurity = 2;
                var target = fixture.Resolve();
                fixture.Application.AutomationSecurity = 3;
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Host.Invoke(target, "Support", "Run", null));
                Assert.AreEqual(3, fixture.Application.AutomationSecurity);
                Assert.AreEqual(0, fixture.Invocations);
            }
        }

        [TestMethod]
        public void ApplicationWindowReaderUsesTheExactApplicationAndRefusesZeroOrFailedProof()
        {
            using (var fixture = new Fixture())
            {
                int ownerReads = 0;
                fixture.Host.ReadWindowOwner = window => { ownerReads++; return 123; };
                fixture.Host.ReadApplicationWindow = application => {
                    Assert.AreSame(fixture.Application, application); return IntPtr.Zero;
                };
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Resolve());
                Assert.AreEqual(0, ownerReads);
                fixture.Host.ReadApplicationWindow = application => {
                    Assert.AreSame(fixture.Application, application); throw new InvalidOperationException("Window proof unavailable");
                };
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Resolve());
                Assert.AreEqual(0, ownerReads);
                Assert.AreEqual(2, fixture.ApplicationReads);
                Assert.AreEqual(0, fixture.Invocations);
            }
        }

        [TestMethod]
        public void ExactProjectAndSavedPathAreRecheckedBeforeEachInvocation()
        {
            using (var fixture = new Fixture())
            {
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Host.ResolveTarget(new object(), fixture.Source.FullName));
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Host.ResolveTarget(fixture.Source.VBProject, "Original.pptm"));
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Host.ResolveTarget(fixture.Source.VBProject, @"C:Original.pptm"));
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Host.ResolveTarget(fixture.Source.VBProject, @"\Original.pptm"));
                object target = fixture.Resolve();
                fixture.Source.FullName = Path.Combine(fixture.Folder, "Renamed.pptm");
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Host.Invoke(target, "Support", "Run", new object[0]));
                fixture.Source.FullName = Path.Combine(fixture.Folder, "Original.pptm");
                fixture.Application.Presentations.Clear();
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Host.Invoke(target, "Support", "Run", new object[0]));
                Assert.AreEqual(0, fixture.Invocations);
            }
        }

        [TestMethod]
        public void DuplicateLoadedFilenamesAndForeignTargetsAreRefused()
        {
            using (var fixture = new Fixture())
            {
                object target = fixture.Resolve();
                fixture.Application.Presentations.Add(new Presentation { FullName = Path.Combine(fixture.Folder, "Other", "Original.pptm") });
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Host.Invoke(target, "Support", "Run", new object[0]));
                var other = new VbaTestPowerPointValuesHost();
                Assert.ThrowsException<InvalidOperationException>(() => other.Invoke(target, "Support", "Run", new object[0]));
                Assert.AreEqual(0, fixture.Invocations);
            }
        }

        [TestMethod]
        public void WrongThreadCannotReadComAndNativeFailureIsNeverRetried()
        {
            using (var fixture = new Fixture())
            {
                Exception refused = null;
                var thread = new Thread(() => { try { fixture.Resolve(); } catch (Exception error) { refused = error; } });
                thread.Start(); thread.Join();
                Assert.IsInstanceOfType(refused, typeof(InvalidOperationException));
                Assert.AreEqual(0, fixture.ApplicationReads);
                fixture.Host.RunProcedure = (application, macro, array) => { fixture.Invocations++; throw new InvalidOperationException("Unverified native completion"); };
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Host.Invoke(fixture.Resolve(), "Support", "Run", new object[0]));
                Assert.AreEqual(1, fixture.Invocations);
            }
        }

        [TestMethod]
        public void InvalidIdentityCollectionAndIdentifiersNeverDispatch()
        {
            using (var f = new Fixture())
            {
                Assert.ThrowsException<InvalidOperationException>(() => f.Host.ResolveTarget(null, f.Source.FullName));
                Assert.ThrowsException<InvalidOperationException>(() => f.Host.Invoke(null, "Support", "Run", null));
                var target = f.Resolve();
                foreach (var pair in new[] { new[] { (string)null, "Run" }, new[] { "bad!", "Run" }, new[] { "Support", (string)null }, new[] { "Support", "bad!" } })
                    Assert.ThrowsException<InvalidOperationException>(() => f.Host.Invoke(target, pair[0], pair[1], null));
                f.Source.NameOverride = "Other.pptm";
                Assert.ThrowsException<InvalidOperationException>(() => f.Host.Invoke(target, "Support", "Run", null));
                f.Source.NameOverride = null;
                f.Host.Invoke(target, "Support", "Run", null);
                f.Host.ReadActiveApplication = _ => null;
                Assert.ThrowsException<InvalidOperationException>(() => f.Host.ValidateTarget(target));
                f.Host.ReadApplicationWindow = _ => new IntPtr(99);
                f.Host.ReadActiveApplication = _ => new Application();
                Assert.ThrowsException<InvalidOperationException>(() => f.Host.ValidateTarget(target));
                f.Host.ReadActiveApplication = _ => f.Application;
                var replacement = new Presentation { FullName = f.Source.FullName, VBProject = f.Source.VBProject };
                f.Application.Presentations[0] = replacement;
                Assert.ThrowsException<InvalidOperationException>(() => f.Host.ValidateTarget(target));
                f.Application.Presentations.Add(f.Source);
                Assert.ThrowsException<InvalidOperationException>(() => f.Resolve());
                f.Application.Presentations.Clear(); f.Application.Presentations.Add(f.Source);
                f.Source.PathOverride = " ";
                Assert.ThrowsException<InvalidOperationException>(() => f.Resolve());
                f.Source.PathOverride = null;
                for (int i = 0; i < 1000; i++) f.Application.Presentations.Add(new Presentation());
                Assert.ThrowsException<InvalidOperationException>(() => f.Resolve());
                Assert.AreEqual(1, f.Invocations);
            }
            Assert.IsFalse(VbaTestPowerPointValuesHost.SamePath(null, @"C:\A.pptm"));
            Assert.IsFalse(VbaTestPowerPointValuesHost.SamePath(@"C:\A.pptm", "A.pptm"));
            Assert.IsFalse(VbaTestPowerPointValuesHost.SamePath(" ", null));
        }

        [TestMethod]
        public void DefaultProcessAndWindowReadersObserveOnlyTheTestProcess()
        {
            var host = new VbaTestPowerPointValuesHost();
            using (var process = System.Diagnostics.Process.GetCurrentProcess())
            {
                Assert.AreEqual(process.ProcessName, host.ReadProcessName());
                Assert.AreEqual(process.Id, host.ReadProcessId());
            }
            Assert.AreEqual(0u, host.ReadWindowOwner(IntPtr.Zero));
            Assert.IsFalse(host.SameIdentity(new object(), new object()));
        }

        private sealed class ApplicationProxy : System.Runtime.Remoting.Proxies.RealProxy
        {
            internal string Macro;
            internal object[] Arguments;
            internal readonly object Result = new object();
            internal ApplicationProxy() : base(typeof(VbaTestPowerPointValuesHost).Assembly.GetType("Microsoft.Office.Interop.PowerPoint._Application", true)) { }
            public override System.Runtime.Remoting.Messaging.IMessage Invoke(System.Runtime.Remoting.Messaging.IMessage message)
            {
                var call = (System.Runtime.Remoting.Messaging.IMethodCallMessage)message;
                Assert.AreEqual("Run", call.MethodName);
                Macro = (string)call.Args[0]; Arguments = (object[])call.Args[1];
                return new System.Runtime.Remoting.Messaging.ReturnMessage(Result, call.Args, call.ArgCount, call.LogicalCallContext, call);
            }
        }

        [TestMethod]
        public void NativeRunUsesThePowerPointPiaReferenceArraySignature()
        {
            var host = new VbaTestPowerPointValuesHost();
            var proxy = new ApplicationProxy(); object[] arguments = { "Tests", "Alpha" };
            Assert.AreSame(proxy.Result, host.RunProcedure(proxy.GetTransparentProxy(), "Owned.pptm!Support.Run", arguments));
            Assert.AreEqual("Owned.pptm!Support.Run", proxy.Macro);
            Assert.AreSame(arguments, proxy.Arguments);
            Assert.ThrowsException<InvalidCastException>(() => host.RunProcedure(new object(), "Owned.pptm!Support.Run", arguments));
        }

        [TestMethod]
        public void AbsolutePathGuardsRejectRootRelativeDriveRelativeAndUnsafeFilename()
        {
            foreach (string path in new[] { (string)null, " ", "relative.pptm", @"\relative.pptm", @"C:relative.pptm" })
                Assert.ThrowsException<InvalidOperationException>(() => VbaTestPowerPointValuesHost.RequireAbsolutePath(path));
            using (var f = new Fixture())
            {
                f.Source.FullName = Path.Combine(f.Folder, "Unsafe!name.pptm");
                Assert.ThrowsException<InvalidOperationException>(() => f.Host.Invoke(f.Resolve(), "Support", "Run", null));
                Assert.AreEqual(0, f.Invocations);
            }
        }

        internal sealed class Fixture : IDisposable
        {
            internal readonly string Folder = Path.Combine(Path.GetTempPath(), "VBAi-PowerPoint-" + Guid.NewGuid().ToString("N"));
            internal readonly Application Application = new Application();
            internal readonly Presentation Source;
            internal readonly VbaTestPowerPointValuesHost Host = new VbaTestPowerPointValuesHost();
            internal int ApplicationReads, Invocations;
            internal Fixture()
            {
                Source = new Presentation { FullName = Path.Combine(Folder, "Original.pptm") };
                Application.Presentations.Add(Source);
                Host.ReadProcessName = () => "POWERPNT";
                Host.ReadProcessId = () => 123;
                Host.ReadApplicationWindow = application => { Assert.AreSame(Application, application); return new IntPtr(99); };
                Host.ReadWindowOwner = hwnd => { Assert.AreEqual(new IntPtr(99), hwnd); return 123; };
                Host.ReadActiveApplication = progId => { ApplicationReads++; Assert.AreEqual("PowerPoint.Application", progId); return Application; };
                Host.RunProcedure = (application, macro, array) => { Invocations++; return new object[] { "Passed", "", "0" }; };
            }
            internal object Resolve() => Host.ResolveTarget(Source.VBProject, Source.FullName);
            public void Dispose() { if (Directory.Exists(Folder)) Directory.Delete(Folder, true); }
        }

        public sealed class Application
        {
            public int AutomationSecurity = 2; // msoAutomationSecurityByUI; the transport must not change it.
            public Presentations Presentations { get; } = new Presentations();
        }
        public sealed class Presentations : List<Presentation>
        {
            public int OpenCalls, ReadOnly, Untitled, WithWindow;
            public Func<string, Presentation> OnOpen;
            public Presentation Open(string path, int readOnly, int untitled, int withWindow)
            {
                OpenCalls++; ReadOnly = readOnly; Untitled = untitled; WithWindow = withWindow;
                var copy = OnOpen == null ? new Presentation { FullName = path } : OnOpen(path);
                if (copy != null && !Contains(copy)) Add(copy);
                if (copy != null) copy.OnClose = () => Remove(copy);
                return copy;
            }
        }
        public sealed class Presentation
        {
            public object VBProject { get; set; } = new object();
            public string FullName { get; set; }
            public string NameOverride, PathOverride;
            public string Name => NameOverride ?? System.IO.Path.GetFileName(FullName);
            public string Path => PathOverride ?? System.IO.Path.GetDirectoryName(FullName);
            public int SaveCalls, SaveAsCalls, SaveCopyCalls, CloseCalls, LastFormat, EmbedFonts;
            public string LastCopyPath;
            public Action OnSaveCopy;
            public Action OnClose;
            public void SaveCopyAs(string path, int format, int embedFonts)
            { SaveCopyCalls++; LastCopyPath = path; LastFormat = format; EmbedFonts = embedFonts; OnSaveCopy?.Invoke(); }
            public void Close() { CloseCalls++; OnClose?.Invoke(); }
        }
    }
}
