using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    [TestClass, TestCategory("Unit"), DoNotParallelize]
    public sealed class VbaTestWordValuesHostTests
    {
        [TestMethod]
        public void InvocationUsesDocumentQualifiedNameAndZeroOrTwoPositionalArguments()
        {
            using (var fixture = new Fixture())
            {
                object native = new object(); fixture.Application.Result = native;
                var target = fixture.Resolve();
                Assert.AreSame(native, fixture.Host.Invoke(target, "Support", "Reset", new object[0]));
                Assert.AreEqual("Support.Reset", fixture.Application.LastMacro);
                Assert.IsNull(fixture.Application.First);
                Assert.AreSame(native, fixture.Host.Invoke(target, "Support", "Run", new object[] { "Tests", "Alpha" }));
                Assert.AreEqual("Tests", fixture.Application.First);
                Assert.AreEqual("Alpha", fixture.Application.Second);
                Assert.AreEqual(2, fixture.Application.RunCalls);
                Assert.AreEqual(2, fixture.Source.ActivateCalls);
            }
        }

        [TestMethod]
        public void WrongProcessPidProjectAndStalePathNeverActivateOrRun()
        {
            using (var fixture = new Fixture())
            {
                fixture.Host.ReadProcessName = () => "EXCEL";
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Resolve());
                Assert.AreEqual(0, fixture.Reads);
                fixture.Host.ReadProcessName = () => "WINWORD";
                fixture.Host.ReadWindowOwner = hwnd => 999;
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Resolve());
                fixture.Host.ReadWindowOwner = hwnd => 123;
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Host.ResolveTarget(new object(), fixture.Source.FullName));
                var target = fixture.Resolve();
                fixture.Source.FullName = Path.Combine(fixture.Folder, "Renamed.docm");
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Host.Invoke(target, "Support", "Run", new object[0]));
                Assert.AreEqual(0, fixture.Source.ActivateCalls + fixture.Application.RunCalls);
            }
        }

        [TestMethod]
        public void DuplicateFilenamesUnsafeQuotingAndUnsupportedArgumentsRefuseBeforeActivation()
        {
            using (var fixture = new Fixture())
            {
                var target = fixture.Resolve();
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Host.Invoke(target, "Support", "Run", new object[] { 1 }));
                fixture.Application.Documents.Add(new Document { FullName = Path.Combine(fixture.Folder, "Other", "Original.docm"), Application = fixture.Application });
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Host.Invoke(target, "Support", "Run", new object[0]));
                fixture.Application.Documents.RemoveAt(1);
                fixture.Source.FullName = Path.Combine(fixture.Folder, "User's.docm");
                target = fixture.Resolve();
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Host.Invoke(target, "Support", "Run", new object[0]));
                Assert.AreEqual(0, fixture.Source.ActivateCalls + fixture.Application.RunCalls);
            }
        }

        [TestMethod]
        public void ActivationContextFailureIsUncertainWithoutFallbackOrRetry()
        {
            foreach (string fault in new[] { "path", "context", "pid", "name" })
            using (var fixture = new Fixture())
            {
                var target = fixture.Resolve();
                fixture.Source.OnActivate = () => {
                    if (fault == "path") fixture.Source.FullName = Path.Combine(fixture.Folder, "Changed.docm");
                    if (fault == "context") fixture.Application.ActiveDocument = new Document();
                    if (fault == "pid") fixture.Host.ReadWindowOwner = hwnd => 999;
                    if (fault == "name") fixture.Application.Documents.Add(new Document { FullName = Path.Combine(fixture.Folder, "Other", "Original.docm") });
                };
                var error = Assert.ThrowsException<VbaTestInvocationException>(() => fixture.Host.Invoke(target, "Support", "Run", new object[0]));
                Assert.IsTrue(error.Uncertain, fault);
                Assert.AreEqual(1, fixture.Source.ActivateCalls, fault);
                Assert.AreEqual(0, fixture.Application.RunCalls, fault);
            }
        }

        [TestMethod]
        public void NativeMacroFailureIsUncertainAndAttemptedExactlyOnce()
        {
            using (var fixture = new Fixture())
            {
                int attempts = 0;
                fixture.Host.RunProcedure = (application, macro, arguments) => { attempts++; throw new InvalidOperationException("Native completion unavailable"); };
                var error = Assert.ThrowsException<VbaTestInvocationException>(() => fixture.Host.Invoke(fixture.Resolve(), "Support", "Run", new object[] { "Tests", "Alpha" }));
                Assert.IsTrue(error.Uncertain);
                Assert.AreEqual(1, attempts);
            }
        }

        [TestMethod]
        public void WrongThreadAndDriveRelativePathsCannotReadWordCom()
        {
            using (var fixture = new Fixture())
            {
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Host.ResolveTarget(fixture.Source.VBProject, @"C:Original.docm"));
                Exception error = null;
                var thread = new Thread(() => { try { fixture.Resolve(); } catch (Exception caught) { error = caught; } });
                thread.Start(); thread.Join();
                Assert.IsInstanceOfType(error, typeof(InvalidOperationException));
                Assert.AreEqual(0, fixture.Reads);
            }
        }

        [TestMethod]
        public void MissingOrForeignTargetsAndInvalidIdentifiersRefuseBeforeActivation()
        {
            using (var f = new Fixture())
            {
                Assert.ThrowsException<InvalidOperationException>(() => f.Host.ResolveTarget(null, f.Source.FullName));
                Assert.ThrowsException<InvalidOperationException>(() => f.Host.Invoke(null, "Support", "Run", null));
                Assert.ThrowsException<InvalidOperationException>(() => f.Host.Invoke(new VbaTestWordValuesHost.OwnedTarget { Owner = new VbaTestWordValuesHost() }, "Support", "Run", null));
                var target = f.Resolve();
                foreach (var pair in new[] { new[] { (string)null, "Run" }, new[] { "bad!", "Run" }, new[] { "Support", (string)null }, new[] { "Support", "bad!" } })
                    Assert.ThrowsException<InvalidOperationException>(() => f.Host.Invoke(target, pair[0], pair[1], null));
                f.Host.Invoke(target, "Support", "Run", null);
                Assert.AreEqual(1, f.Application.RunCalls);
            }
        }

        [TestMethod]
        public void ChangedApplicationOrDocumentIdentityAndUnboundedCollectionAreRejected()
        {
            using (var f = new Fixture())
            {
                var target = f.Resolve();
                f.Host.ReadActiveApplication = _ => null;
                Assert.ThrowsException<InvalidOperationException>(() => f.Host.ValidateTarget(target));
                f.Host.ReadActiveApplication = _ => new Application();
                Assert.ThrowsException<InvalidOperationException>(() => f.Host.ValidateTarget(target));
                f.Host.ReadActiveApplication = _ => f.Application;
                var replacement = new Document { FullName = f.Source.FullName, VBProject = f.Source.VBProject, Application = f.Application };
                f.Application.Documents[0] = replacement;
                Assert.ThrowsException<InvalidOperationException>(() => f.Host.ValidateTarget(target));
                f.Application.Documents.Add(f.Source);
                Assert.ThrowsException<InvalidOperationException>(() => f.Resolve());
                f.Application.Documents.Clear(); f.Application.Documents.Add(f.Source);
                for (int i = 0; i < 1000; i++) f.Application.Documents.Add(new Document());
                Assert.ThrowsException<InvalidOperationException>(() => f.Resolve());
                Assert.AreEqual(0, f.Application.RunCalls);
            }
        }

        [TestMethod]
        public void NameAndPathValidationRejectUnsafeSavedMacroContexts()
        {
            using (var f = new Fixture())
            {
                f.Source.PathOverride = " ";
                Assert.ThrowsException<InvalidOperationException>(() => f.Resolve());
                f.Source.PathOverride = null;
                var target = f.Resolve();
                f.Source.NameOverride = "Other.docm";
                Assert.ThrowsException<InvalidOperationException>(() => f.Host.Invoke(target, "Support", "Run", null));
                Assert.AreEqual(0, f.Source.ActivateCalls);
            }
            Assert.IsFalse(VbaTestWordValuesHost.SamePath(null, @"C:\A.docm"));
            Assert.IsFalse(VbaTestWordValuesHost.SamePath(@"C:\A.docm", "A.docm"));
            Assert.IsFalse(VbaTestWordValuesHost.SamePath(@"\A.docm", @"C:\A.docm"));
            Assert.IsFalse(VbaTestWordValuesHost.SamePath(@"C:A.docm", @"C:\A.docm"));
            Assert.IsFalse(VbaTestWordValuesHost.SamePath(" ", null));
        }

        [TestMethod]
        public void DefaultProcessAndWindowReadersObserveOnlyTheTestProcess()
        {
            var host = new VbaTestWordValuesHost();
            using (var process = System.Diagnostics.Process.GetCurrentProcess())
            {
                Assert.AreEqual(process.ProcessName, host.ReadProcessName());
                Assert.AreEqual(process.Id, host.ReadProcessId());
            }
            Assert.AreEqual(0u, host.ReadWindowOwner(IntPtr.Zero));
            var indexed = new OneBasedCollection<Document> { new Document() };
            Assert.AreSame(indexed[1], host.ReadDocumentItem(indexed, 1));
            Assert.IsFalse(host.SameIdentity(new object(), new object()));
        }

        [TestMethod]
        public void ActiveWordWindowHandleIsRequiredBeforeDocumentsCanBeInspected()
        {
            using (var f = new Fixture())
            {
                Assert.IsNull(typeof(Application).GetProperty("Hwnd"), "The Word application model must not expose the nonexistent application HWND.");
                Assert.AreEqual(new IntPtr(99), VbaTestWordValuesHost.ReadApplicationWindow(f.Application));
                Assert.ThrowsException<InvalidOperationException>(() => VbaTestWordValuesHost.ReadApplicationWindow(null));
                f.Application.ActiveWindow = null;
                Assert.ThrowsException<InvalidOperationException>(() => f.Resolve());
                f.Application.ActiveWindow = new Window();
                Assert.ThrowsException<InvalidOperationException>(() => f.Resolve());
                f.Application.ActiveWindow.Hwnd = 99;
                f.Host.ReadWindowOwner = _ => 0;
                Assert.ThrowsException<InvalidOperationException>(() => f.Resolve());
                f.Host.ReadWindowOwner = _ => 999;
                Assert.ThrowsException<InvalidOperationException>(() => f.Resolve());
                Assert.AreEqual(0, f.Source.ActivateCalls + f.Application.RunCalls);
            }
        }

        [TestMethod]
        public void ActiveWordWindowOwnershipIsRecheckedAfterDocumentActivation()
        {
            foreach (string fault in new[] { "missing", "zero", "foreign" })
            using (var f = new Fixture())
            {
                var target = f.Resolve();
                f.Source.OnActivate = () =>
                {
                    if (fault == "missing") f.Application.ActiveWindow = null;
                    if (fault == "zero") f.Application.ActiveWindow.Hwnd = 0;
                    if (fault == "foreign") f.Host.ReadWindowOwner = _ => 999;
                };
                var error = Assert.ThrowsException<VbaTestInvocationException>(() => f.Host.Invoke(target, "Support", "Run", null));
                Assert.IsTrue(error.Uncertain, fault);
                Assert.AreEqual(1, f.Source.ActivateCalls, fault);
                Assert.AreEqual(0, f.Application.RunCalls, fault);
            }
        }
        [TestMethod]
        public void ModuleQualifierMustHaveTheExactSingleOwnerIncludingGlobalTemplates()
        {
            foreach (string fault in new[] { "foreign", "missing", "duplicate", "unreadable", "projects", "components" })
            using (var f = new Fixture())
            {
                var target = f.Resolve();
                var project = (Project)f.Source.VBProject;
                if (fault == "foreign") f.Application.VBE.VBProjects.Add(new Project());
                if (fault == "missing") project.Components.Clear();
                if (fault == "duplicate") project.Components.Add(new Component { Name = "Support" });
                if (fault == "unreadable") f.Application.VBE.VBProjects.Add(new Project { ComponentsError = new InvalidOperationException("Protected template components unavailable") });
                if (fault == "projects") f.Application.VBE.VBProjects.CountOverride = 1001;
                if (fault == "components") project.Components.CountOverride = 1001;
                Assert.ThrowsException<InvalidOperationException>(() => f.Host.Invoke(target, "Support", "Run", new object[0]), fault);
                Assert.AreEqual(0, f.Source.ActivateCalls + f.Application.RunCalls, fault);
            }
            using (var f = new Fixture())
            {
                var template = new Project(); template.Components[1].Name = "OtherModule";
                f.Application.VBE.VBProjects.Add(template);
                f.Host.Invoke(f.Resolve(), "Support", "Run", new object[] { "Tests", "Alpha" });
                Assert.AreEqual("Support.Run", f.Application.LastMacro);
                Assert.AreEqual(1, f.Application.RunCalls);
            }
        }

        [TestMethod]
        public void ModuleOwnershipIsRecheckedAfterExactDocumentActivation()
        {
            foreach (string fault in new[] { "collision", "missing" })
            using (var f = new Fixture())
            {
                var target = f.Resolve();
                f.Source.OnActivate = () => {
                    if (fault == "collision") f.Application.VBE.VBProjects.Add(new Project());
                    else ((Project)f.Source.VBProject).Components.Clear();
                };
                Assert.IsTrue(Assert.ThrowsException<VbaTestInvocationException>(() => f.Host.Invoke(target, "Support", "Run", null)).Uncertain);
                Assert.AreEqual(1, f.Source.ActivateCalls);
                Assert.AreEqual(0, f.Application.RunCalls);
            }
        }

        [TestMethod]
        public void OnlyGeneratedCoverageZeroArgumentFunctionsReceiveExplicitSentinels()
        {
            foreach (string procedure in new[] { VbaCoverageInstrumentation.ResetProcedure, VbaCoverageInstrumentation.SnapshotProcedure })
            using (var f = new Fixture())
            {
                ((Project)f.Source.VBProject).Components.Add(new Component { Name = VbaCoverageInstrumentation.ModuleName });
                object[] received = null;
                f.Host.RunProcedure = (application, macro, arguments) => {
                    Assert.AreEqual(VbaCoverageInstrumentation.ModuleName.ToLowerInvariant() + "." + procedure.ToLowerInvariant(), macro);
                    received = arguments; return f.Application.Result;
                };
                f.Host.Invoke(f.Resolve(), VbaCoverageInstrumentation.ModuleName.ToLowerInvariant(), procedure.ToLowerInvariant(), new object[0]);
                CollectionAssert.AreEqual(new object[] { false, false }, received);
            }
            using (var f = new Fixture())
            {
                ((Project)f.Source.VBProject).Components.Add(new Component { Name = VbaCoverageInstrumentation.ModuleName });
                object[] received = null;
                f.Host.RunProcedure = (application, macro, arguments) => { received = arguments; return null; };
                var target = f.Resolve();
                f.Host.Invoke(target, "Support", VbaCoverageInstrumentation.ResetProcedure, new object[0]);
                Assert.AreEqual(0, received.Length);
                f.Host.Invoke(target, VbaCoverageInstrumentation.ModuleName, "OtherFunction", new object[0]);
                Assert.AreEqual(0, received.Length);
                f.Host.Invoke(target, VbaCoverageInstrumentation.ModuleName, VbaCoverageInstrumentation.ResetProcedure, new object[] { "Tests", "Alpha" });
                CollectionAssert.AreEqual(new object[] { "Tests", "Alpha" }, received);
            }
        }

        [TestMethod]
        public void AcquiredComAliasesAreBalancedWithoutReleasingBorrowedTargetReferences()
        {
            var previousCheck = VbaTestWordValuesHost.IsComReference;
            var previousRelease = VbaTestWordValuesHost.ReleaseComReference;
            var released = new List<object>();
            try
            {
                VbaTestWordValuesHost.IsComReference = value => true;
                VbaTestWordValuesHost.ReleaseComReference = value => { released.Add(value); return 1; };
                using (var f = new Fixture())
                {
                    var target = f.Resolve();
                    Assert.AreEqual(0, released.FindAll(value => ReferenceEquals(value, f.Source)).Count, "The acquired matched document is transferred to its target.");
                    released.Clear();
                    f.Host.Invoke(target, "Support", "Run", new object[0]);
                    Assert.AreEqual(5, released.FindAll(value => ReferenceEquals(value, f.Source)).Count, "Two validation matches, two filename items and one ActiveDocument getter are independent acquisitions.");
                    Assert.AreEqual(0, released.FindAll(value => ReferenceEquals(value, f.Application)).Count, "ReadActiveApplication is borrowed by this seam.");
                }
            }
            finally { VbaTestWordValuesHost.IsComReference = previousCheck; VbaTestWordValuesHost.ReleaseComReference = previousRelease; }
            using (var f = new Fixture())
            {
                f.Application.Documents.CountOverride = -1;
                Assert.ThrowsException<InvalidOperationException>(() => f.Resolve());
            }
        }

        [TestMethod]
        public void NativeTargetOwnsEachAcquisitionAndDisposesOnlyOnce()
        {
            WithTrackedComReleases(released =>
            {
                using (var f = new Fixture(ownsApplication: true))
                {
                    var target = (VbaTestWordValuesHost.OwnedTarget)f.Resolve();
                    released.Clear();
                    f.Host.ValidateTarget(target);
                    Assert.AreEqual(1, CountIdentity(released, f.Application), "Validation balances its temporary ROT acquisition.");
                    Assert.AreEqual(1, CountIdentity(released, f.Source), "Validation balances its document index acquisition.");
                    released.Clear();
                    target.Dispose(); target.Dispose(); target.RetainOnUncertain();
                    Assert.AreEqual(1, CountIdentity(released, f.Source));
                    Assert.AreEqual(1, CountIdentity(released, f.Application));
                    Assert.AreEqual(0, CountIdentity(released, f.Source.VBProject), "The supplied project is borrowed.");
                    Assert.IsNull(target.Application); Assert.IsNull(target.Document); Assert.IsNull(target.Project);
                    Assert.IsFalse(target.IsRetained);
                    int reads = f.Reads;
                    Assert.ThrowsException<InvalidOperationException>(() => f.Host.ValidateTarget(target));
                    Assert.AreEqual(reads, f.Reads, "A released target cannot acquire more COM references.");
                    var empty = new VbaTestWordValuesHost.OwnedTarget { Owner = f.Host };
                    empty.Dispose(); empty.Dispose();
                }
                using (var f = new Fixture())
                {
                    var target = (VbaTestWordValuesHost.OwnedTarget)f.Resolve();
                    released.Clear(); target.Dispose();
                    Assert.AreEqual(1, CountIdentity(released, f.Source));
                    Assert.AreEqual(0, CountIdentity(released, f.Application), "An injected application remains borrowed.");
                }
            });
        }

        [TestMethod]
        public void NativeApplicationLeaseBalancesRefusalsAndBorrowedOverrides()
        {
            WithTrackedComReleases(released =>
            {
                using (var f = new Fixture(ownsApplication: true))
                {
                    f.Application.AutomationSecurity = 3;
                    Assert.ThrowsException<InvalidOperationException>(() => f.Resolve());
                    Assert.AreEqual(1, CountIdentity(released, f.Application));
                    f.Application.AutomationSecurity = 2;
                    released.Clear();
                    Assert.ThrowsException<InvalidOperationException>(() => f.Host.ResolveTarget(new object(), f.Source.FullName));
                    Assert.AreEqual(1, CountIdentity(released, f.Application));
                    var target = (VbaTestWordValuesHost.OwnedTarget)f.Resolve();
                    released.Clear();
                    f.Host.SameIdentity = (left, right) => false;
                    Assert.ThrowsException<InvalidOperationException>(() => f.Host.ValidateTarget(target));
                    Assert.AreEqual(1, CountIdentity(released, f.Application));
                    f.Host.SameIdentity = ReferenceEquals;
                    f.Application.Documents.Clear();
                    released.Clear();
                    Assert.ThrowsException<InvalidOperationException>(() => f.Host.ValidateTarget(target));
                    Assert.AreEqual(1, CountIdentity(released, f.Application));
                    target.Dispose();
                    f.Host.ReadActiveApplication = _ => f.Application;
                    using (var lease = f.Host.ResolveApplicationLease())
                    {
                        Assert.AreSame(f.Application, lease.Application);
                        released.Clear(); lease.Dispose(); lease.Dispose(); lease.RetainOnUncertain();
                        Assert.AreEqual(0, CountIdentity(released, f.Application));
                        Assert.IsNull(lease.Application);
                    }
                }
            });
        }

        [TestMethod]
        public void TargetAndApplicationDisposalRefuseForeignThreadsWithoutReleasing()
        {
            WithTrackedComReleases(released =>
            {
                using (var f = new Fixture(ownsApplication: true))
                {
                    var target = (VbaTestWordValuesHost.OwnedTarget)f.Resolve();
                    var lease = f.Host.ResolveApplicationLease();
                    released.Clear();
                    Exception targetError = null, leaseError = null;
                    var thread = new Thread(() =>
                    {
                        try { target.Dispose(); } catch (Exception error) { targetError = error; }
                        try { lease.Dispose(); } catch (Exception error) { leaseError = error; }
                    });
                    thread.Start(); thread.Join();
                    Assert.IsInstanceOfType(targetError, typeof(InvalidOperationException));
                    Assert.IsInstanceOfType(leaseError, typeof(InvalidOperationException));
                    Assert.AreEqual(0, released.Count);
                    target.Dispose(); lease.Dispose();
                    Assert.AreEqual(2, CountIdentity(released, f.Application));
                    Assert.AreEqual(1, CountIdentity(released, f.Source));
                }
            });
        }

        [TestMethod]
        public void UncertainInvocationRetainsTargetAndPreventsFurtherNativeDispatch()
        {
            WithTrackedComReleases(released =>
            {
                using (var f = new Fixture(ownsApplication: true))
                {
                    var target = (VbaTestWordValuesHost.OwnedTarget)f.Resolve();
                    int attempts = 0;
                    f.Host.RunProcedure = (application, macro, arguments) => { attempts++; throw new InvalidOperationException("Unknown native outcome"); };
                    Assert.IsTrue(Assert.ThrowsException<VbaTestInvocationException>(() => f.Host.Invoke(target, "Support", "Run", null)).Uncertain);
                    Assert.IsTrue(target.IsRetained);
                    released.Clear();
                    target.Dispose(); target.Dispose(); target.RetainOnUncertain();
                    Assert.AreEqual(0, released.Count, "Unknown native targets must retain their application and document acquisitions.");
                    Assert.AreSame(f.Application, target.Application); Assert.AreSame(f.Source, target.Document); Assert.AreSame(f.Source.VBProject, target.Project);
                    Assert.ThrowsException<InvalidOperationException>(() => f.Host.Invoke(target, "Support", "Run", null));
                    Assert.AreEqual(1, attempts);
                    using (var lease = f.Host.ResolveApplicationLease())
                    {
                        released.Clear(); lease.RetainOnUncertain(); lease.RetainOnUncertain(); lease.Dispose();
                        Assert.AreEqual(0, released.Count, "Unknown close verification retains its temporary application lease.");
                        Assert.AreSame(f.Application, lease.Application);
                    }
                }
                VbaTestWordValuesHost.RetainAcquired(null);
            });
        }

        [TestMethod]
        public void DocumentReleaseFailureStillBalancesApplicationWithoutRepeatingEitherRelease()
        {
            WithTrackedComReleases(released =>
            {
                using (var f = new Fixture(ownsApplication: true))
                {
                    var target = (VbaTestWordValuesHost.OwnedTarget)f.Resolve();
                    released.Clear();
                    VbaTestWordValuesHost.ReleaseComReference = value =>
                    {
                        released.Add(value);
                        if (ReferenceEquals(value, f.Source)) throw new InvalidOperationException("Document release failed");
                        return 0;
                    };
                    Assert.ThrowsException<InvalidOperationException>(() => target.Dispose());
                    target.Dispose();
                    Assert.AreEqual(1, CountIdentity(released, f.Source));
                    Assert.AreEqual(1, CountIdentity(released, f.Application));
                }
            });
        }

        private static int CountIdentity(List<object> values, object expected)
        { return values.FindAll(value => ReferenceEquals(value, expected)).Count; }

        private static void WithTrackedComReleases(Action<List<object>> test)
        {
            var previousCheck = VbaTestWordValuesHost.IsComReference;
            var previousRelease = VbaTestWordValuesHost.ReleaseComReference;
            var released = new List<object>();
            try
            {
                VbaTestWordValuesHost.IsComReference = value => true;
                VbaTestWordValuesHost.ReleaseComReference = value => { released.Add(value); return 0; };
                test(released);
            }
            finally { VbaTestWordValuesHost.IsComReference = previousCheck; VbaTestWordValuesHost.ReleaseComReference = previousRelease; }
        }

        private sealed class ApplicationProxy : System.Runtime.Remoting.Proxies.RealProxy
        {
            internal string Macro;
            internal object[] Arguments;
            internal int RunCalls;
            internal Exception Failure;
            internal readonly object Result = new object();
            internal ApplicationProxy() : base(typeof(VbaTestWordValuesHost).Assembly.GetType("Microsoft.Office.Interop.Word._Application", true)) { }
            public override System.Runtime.Remoting.Messaging.IMessage Invoke(System.Runtime.Remoting.Messaging.IMessage message)
            {
                var call = (System.Runtime.Remoting.Messaging.IMethodCallMessage)message;
                Assert.AreEqual("Run", call.MethodName);
                Assert.AreEqual(31, call.ArgCount);
                RunCalls++;
                Macro = (string)call.Args[0];
                Arguments = new object[30];
                Array.Copy(call.Args, 1, Arguments, 0, Arguments.Length);
                if (Failure != null) return new System.Runtime.Remoting.Messaging.ReturnMessage(Failure, call);
                return new System.Runtime.Remoting.Messaging.ReturnMessage(Result, call.Args, call.ArgCount, call.LogicalCallContext, call);
            }
        }

        [DataTestMethod]
        [DataRow(0)]
        [DataRow(2)]
        public void NativeRunUsesThirtyWordPiaByReferenceSlotsAndMissingOptionalValues(int argumentCount)
        {
            var host = new VbaTestWordValuesHost();
            var proxy = new ApplicationProxy();
            object[] arguments = argumentCount == 0 ? new object[0] : new object[] { "Tests", "Alpha" };
            var before = (object[])arguments.Clone();
            Assert.AreSame(proxy.Result, host.RunProcedure(proxy.GetTransparentProxy(), "'Owned.docm'!Support.Run", arguments));
            Assert.AreEqual("'Owned.docm'!Support.Run", proxy.Macro);
            Assert.AreEqual(1, proxy.RunCalls);
            for (int i = 0; i < proxy.Arguments.Length; i++)
                if (i < argumentCount) Assert.AreSame(arguments[i], proxy.Arguments[i]);
                else Assert.AreSame(Type.Missing, proxy.Arguments[i], "Optional slot " + (i + 1));
            CollectionAssert.AreEqual(before, arguments);
            var parameters = typeof(VbaTestWordValuesHost).Assembly.GetType("Microsoft.Office.Interop.Word._Application", true).GetMethod("Run").GetParameters();
            Assert.AreEqual(31, parameters.Length);
            for (int i = 1; i < parameters.Length; i++)
            {
                Assert.IsTrue(parameters[i].ParameterType.IsByRef, "Slot " + i);
                Assert.IsTrue(parameters[i].IsOptional, "Slot " + i);
            }
        }

        [TestMethod]
        public void NativeRunRejectsForeignObjectsAndPropagatesWordFailureWithoutRetry()
        {
            var host = new VbaTestWordValuesHost();
            Assert.ThrowsException<InvalidCastException>(() => host.RunProcedure(new object(), "'Owned.docm'!Support.Run", new object[0]));
            var failure = new System.Runtime.InteropServices.COMException("Word Run failed", unchecked((int)0x80020003));
            var proxy = new ApplicationProxy { Failure = failure };
            var error = Assert.ThrowsException<System.Runtime.InteropServices.COMException>(() => host.RunProcedure(proxy.GetTransparentProxy(), "'Owned.docm'!Support.Run", new object[0]));
            Assert.AreSame(failure, error);
            Assert.AreEqual(1, proxy.RunCalls);
            using (var f = new Fixture())
            {
                var guardedProxy = new ApplicationProxy { Failure = failure };
                f.Host.RunProcedure = (application, macro, arguments) => host.RunProcedure(guardedProxy.GetTransparentProxy(), macro, arguments);
                var guardedError = Assert.ThrowsException<VbaTestInvocationException>(() => f.Host.Invoke(f.Resolve(), "Support", "Run", new object[] { "Tests", "Alpha" }));
                Assert.IsTrue(guardedError.Uncertain);
                Assert.AreSame(failure, guardedError.InnerException);
                Assert.AreEqual(1, guardedProxy.RunCalls);
                Assert.AreEqual(1, f.Source.ActivateCalls);
            }
        }

        [TestMethod]
        public void ForceDisableRefusesInitialAndChangedPolicyBeforeActivationOrInvocation()
        {
            using (var f = new Fixture())
            {
                var target = f.Resolve();
                f.Application.AutomationSecurity = 3;
                Assert.ThrowsException<InvalidOperationException>(() => f.Resolve());
                Assert.ThrowsException<InvalidOperationException>(() => f.Host.Invoke(target, "Support", "Run", new object[0]));
                Assert.AreEqual(3, f.Application.AutomationSecurity);
                Assert.AreEqual(0, f.Source.ActivateCalls + f.Application.RunCalls);
            }
        }

        internal sealed class Fixture : IDisposable
        {
            internal readonly string Folder = Path.Combine(Path.GetTempPath(), "VBAi-Word-" + Guid.NewGuid().ToString("N"));
            internal readonly Application Application = new Application();
            internal readonly Document Source;
            internal readonly VbaTestWordValuesHost Host;
            internal int Reads;
            internal Fixture(string extension = ".docm", bool ownsApplication = false)
            {
                Func<string, object> reader = progId => { Reads++; Assert.AreEqual("Word.Application", progId); return Application; };
                Host = ownsApplication ? new VbaTestWordValuesHost(reader) : new VbaTestWordValuesHost();
                if (!ownsApplication) Host.ReadActiveApplication = reader;
                Directory.CreateDirectory(Folder);
                Source = new Document { FullName = Path.Combine(Folder, "Original" + extension), Application = Application };
                File.WriteAllText(Source.FullName, "Saved Word fixture bytes");
                Application.Documents.Add(Source);
                Application.VBE.VBProjects.Add(Source.VBProject);
                Host.ReadDocumentItem = (documents, index) => ((Documents)documents)[index - 1];
                Host.ReadProcessName = () => "WINWORD"; Host.ReadProcessId = () => 123;
                Host.ReadWindowOwner = hwnd => { Assert.AreEqual(new IntPtr(99), hwnd); return 123; };
                Host.RunProcedure = (application, macro, arguments) =>
                {
                    Assert.AreSame(Application, application);
                    return Application.Run(macro, arguments.Length == 0 ? null : arguments[0], arguments.Length == 0 ? null : arguments[1]);
                };
            }
            internal object Resolve() => Host.ResolveTarget(Source.VBProject, Source.FullName);
            public void Dispose() { if (Directory.Exists(Folder)) Directory.Delete(Folder, true); }
        }

        public sealed class Application
        {
            public int AutomationSecurity = 2; // msoAutomationSecurityByUI; never changed by the production transport.
            public Vbe VBE { get; } = new Vbe();
            public Window ActiveWindow { get; set; } = new Window { Hwnd = 99 };
            public Documents Documents { get; } = new Documents();
            public Document ActiveDocument { get; set; }
            public object Result = new object();
            public string LastMacro;
            public object First, Second;
            public int RunCalls;
            public object Run(string macro, object first = null, object second = null)
            { RunCalls++; LastMacro = macro; First = first; Second = second; return Result; }
        }
        public sealed class Window
        {
            public int Hwnd { get; set; }
        }
        public class OneBasedCollection<T> : List<T>
        {
            public int? CountOverride;
            public new int Count => CountOverride ?? base.Count;
            public new T this[int index] { get => base[index - 1]; set => base[index - 1] = value; }
        }
        public sealed class Vbe
        {
            public OneBasedCollection<object> VBProjects { get; } = new OneBasedCollection<object>();
        }
        public sealed class Project
        {
            public Exception ComponentsError;
            public readonly OneBasedCollection<Component> Components = new OneBasedCollection<Component> { new Component { Name = "Support" } };
            public OneBasedCollection<Component> VBComponents => ComponentsError == null ? Components : throw ComponentsError;
        }
        public sealed class Component { public string Name { get; set; } }
        public sealed class Documents : List<Document>
        {
            public int? CountOverride;
            public new int Count => CountOverride ?? base.Count;
            public int OpenCalls;
            public bool LastReadOnly, LastRecent, LastVisible, LastConversions, LastRevert, LastRepair;
            public Func<string, Document> OnOpen;
            public Document Open(string FileName, bool ConfirmConversions, bool ReadOnly, bool AddToRecentFiles, bool Revert, bool Visible, bool OpenAndRepair)
            {
                OpenCalls++; LastReadOnly = ReadOnly; LastRecent = AddToRecentFiles; LastVisible = Visible;
                LastConversions = ConfirmConversions; LastRevert = Revert; LastRepair = OpenAndRepair;
                var document = OnOpen(FileName);
                if (document != null && !Contains(document)) Add(document);
                return document;
            }
        }
        public sealed class Document
        {
            public object VBProject { get; set; } = new Project();
            public string FullName { get; set; }
            public object Saved { get; set; } = true;
            public string NameOverride, PathOverride;
            public string Name => NameOverride ?? System.IO.Path.GetFileName(FullName);
            public string Path => PathOverride ?? System.IO.Path.GetDirectoryName(FullName);
            public Application Application;
            public Action OnActivate;
            public bool CancelClose;
            public int ActivateCalls, CloseCalls, LastSaveChanges, SaveCalls, SaveAsCalls;
            public void Activate() { ActivateCalls++; Application.ActiveDocument = this; OnActivate?.Invoke(); }
            public void Close(int SaveChanges) { CloseCalls++; LastSaveChanges = SaveChanges; if (!CancelClose) Application.Documents.Remove(this); }
        }
    }
}
