using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi;
using TYPEKIND = System.Runtime.InteropServices.ComTypes.TYPEKIND;

namespace VBAi.Tests.Unit
{
    [TestClass]
    public sealed class AccessHelpContextDispatchTests
    {
        internal sealed class Calls : VbeProjectComponents.IAccessHelpContextCalls
        {
            internal int Preparations, Invocations, Owners, Disposals, Member = 117, Result;
            internal Action OnPrepare, OnInvoke;
            internal Exception PrepareError, InvokeError, DisposeError;
            internal string Corruption;
            internal int Value;
            internal readonly List<IntPtr> Freed = new List<IntPtr>();
            internal VbeProjectComponents.AccessHelpContextExceptionInfo ExceptionInfo;
            public int Prepare() { Preparations++; OnPrepare?.Invoke(); if (PrepareError != null) throw PrepareError; return Member; }
            public void RequireOwner() { Owners++; if (Corruption == "owner" && Owners == 2) throw new InvalidOperationException("Owner changed"); }
            public int Invoke(int member, ushort flags, IntPtr result, ref VbeProjectComponents.AccessHelpContextParameters parameters, IntPtr exception, out uint argumentError)
            {
                Invocations++; argumentError = 0;
                Assert.AreEqual(117, member); Assert.AreEqual((ushort)4, flags, "Only scalar PROPERTYPUT is permitted.");
                Assert.AreEqual(IntPtr.Zero, result, "A PROPERTYPUT supplies no result VARIANT.");
                Assert.AreEqual((uint)1, parameters.ArgumentCount); Assert.AreEqual((uint)1, parameters.NamedArgumentCount);
                Assert.AreEqual(-3, Marshal.ReadInt32(parameters.NamedArguments));
                Assert.AreEqual((short)3, Marshal.ReadInt16(parameters.Arguments));
                Assert.AreEqual((short)0, Marshal.ReadInt16(parameters.Arguments, 2));
                Value = Marshal.ReadInt32(parameters.Arguments, 8);
                Marshal.StructureToPtr(ExceptionInfo, exception, false);
                if (Corruption == "argument") Marshal.WriteInt32(parameters.Arguments, 8, Value + 1);
                if (Corruption == "variantBoundary") Marshal.WriteInt64(parameters.Arguments, 24, 0);
                if (Corruption == "exceptionBoundary") Marshal.WriteInt64(exception, 64, 0);
                if (Corruption == "named") Marshal.WriteInt32(parameters.NamedArguments, -4);
                if (Corruption == "parameters") parameters.ArgumentCount = 2;
                OnInvoke?.Invoke();
                if (InvokeError != null) throw InvokeError;
                return Result;
            }
            public void FreeExceptionString(IntPtr text) { Freed.Add(text); }
            public void Dispose() { Disposals++; if (DisposeError != null) throw DisposeError; }
        }

        [TestMethod]
        public void ExactX64LayoutsMatchTheNativeScalarContract()
        {
            VbeProjectComponents.RequireAccessHelpContextAbi();
            Assert.AreEqual(24, Marshal.SizeOf(typeof(VbeProjectComponents.AccessHelpContextParameters)));
            Assert.AreEqual(64, Marshal.SizeOf(typeof(VbeProjectComponents.AccessHelpContextExceptionInfo)));
            var names = new[] { "Code", "Reserved", "Source", "Description", "HelpFile", "HelpContext", "ReservedPointer", "DeferredCallback", "Scode" };
            var offsets = new[] { 0, 2, 8, 16, 24, 32, 40, 48, 56 };
            for (int i = 0; i < names.Length; i++) Assert.AreEqual(offsets[i], Marshal.OffsetOf(typeof(VbeProjectComponents.AccessHelpContextExceptionInfo), names[i]).ToInt32());
            Assert.AreEqual(0, Marshal.OffsetOf(typeof(VbeProjectComponents.AccessHelpContextParameters), "Arguments").ToInt32());
            Assert.AreEqual(8, Marshal.OffsetOf(typeof(VbeProjectComponents.AccessHelpContextParameters), "NamedArguments").ToInt32());
            Assert.AreEqual(16, Marshal.OffsetOf(typeof(VbeProjectComponents.AccessHelpContextParameters), "ArgumentCount").ToInt32());
            Assert.AreEqual(20, Marshal.OffsetOf(typeof(VbeProjectComponents.AccessHelpContextParameters), "NamedArgumentCount").ToInt32());
        }

        [DataTestMethod]
        [DataRow(0)][DataRow(321)][DataRow(int.MinValue)][DataRow(int.MaxValue)]
        public void Int32PutUsesExactPackingOneEntryAndFinalGuard(int value)
        {
            var calls = new Calls(); var dispatch = new VbeProjectComponents.AccessHelpContextDispatch(calls); int guards = 0;
            dispatch.Put(value, () => { guards++; Assert.AreEqual(1, calls.Preparations); Assert.AreEqual(0, calls.Invocations); });
            Assert.AreEqual(value, calls.Value); Assert.AreEqual(1, guards); Assert.AreEqual(1, dispatch.InvokeEntries);
            Assert.AreEqual(1, calls.Invocations); Assert.AreEqual(1, calls.Disposals);
            Assert.ThrowsException<InvalidOperationException>(() => dispatch.Put(value, () => { }));
            Assert.AreEqual(1, calls.Invocations); Assert.AreEqual(1, calls.Disposals);
        }

        [DataTestMethod]
        [DataRow("binding")][DataRow("preparation")][DataRow("owner")][DataRow("authorization")]
        public void ReadonlyPreparationAndFinalRefusalsNeverInvokeOrReplay(string refusal)
        {
            var calls = new Calls(); if (refusal == "binding") calls.Member = 118;
            if (refusal == "preparation") calls.PrepareError = new InvalidOperationException("No type information");
            if (refusal == "owner") calls.Corruption = "owner";
            var dispatch = new VbeProjectComponents.AccessHelpContextDispatch(calls);
            Assert.ThrowsException<InvalidOperationException>(() => dispatch.Put(321, () => { if (refusal == "authorization") throw new InvalidOperationException("Revoked"); }));
            Assert.AreEqual(0, dispatch.InvokeEntries); Assert.AreEqual(0, calls.Invocations); Assert.AreEqual(1, calls.Disposals);
            Assert.ThrowsException<InvalidOperationException>(() => dispatch.Put(321, () => { })); Assert.AreEqual(1, calls.Preparations);
        }

        [DataTestMethod]
        [DataRow("argument")][DataRow("variantBoundary")][DataRow("exceptionBoundary")][DataRow("named")][DataRow("parameters")]
        public void CorruptedNativeBoundaryNeverFreesOutputPointersOrReplays(string corruption)
        {
            var calls = new Calls { Corruption = corruption, ExceptionInfo = new VbeProjectComponents.AccessHelpContextExceptionInfo { Description = new IntPtr(123) } };
            var dispatch = new VbeProjectComponents.AccessHelpContextDispatch(calls);
            var error = Assert.ThrowsException<InvalidOperationException>(() => dispatch.Put(321, () => { }));
            Assert.IsTrue(error.Data.Contains("AccessHelpContextOutputCleanup")); Assert.AreEqual(0, calls.Freed.Count);
            Assert.AreEqual(1, dispatch.InvokeEntries); Assert.ThrowsException<InvalidOperationException>(() => dispatch.Put(321, () => { })); Assert.AreEqual(1, calls.Invocations);
        }

        [DataTestMethod]
        [DataRow(false)][DataRow(true)]
        public void FailedHresultStaysOriginalEvenWithUntrustedOutputAndCleanupFailure(bool corrupt)
        {
            const int hr = unchecked((int)0x80020009);
            var calls = new Calls { Result = hr, Corruption = corrupt ? "argument" : null, DisposeError = new InvalidOperationException("cleanup"),
                ExceptionInfo = new VbeProjectComponents.AccessHelpContextExceptionInfo { Source = new IntPtr(123), Description = new IntPtr(123), HelpFile = new IntPtr(456), Scode = hr } };
            var dispatch = new VbeProjectComponents.AccessHelpContextDispatch(calls);
            var error = Assert.ThrowsException<COMException>(() => dispatch.Put(321, () => { }));
            Assert.AreEqual(hr, error.HResult); Assert.AreEqual(corrupt ? 0 : 2, calls.Freed.Count, "Aliased BSTRs are freed once only when outputs remain trusted.");
            Assert.IsTrue(error.Data.Contains("AccessHelpContextCleanupFailure"));
            Assert.ThrowsException<InvalidOperationException>(() => dispatch.Put(321, () => { })); Assert.AreEqual(1, calls.Invocations);
        }

        [TestMethod]
        public void ThrownInvokePreservesOriginalErrorAndDoesNotTrustOrFreeItsOutput()
        {
            var original = new COMException("Original unknown native return", unchecked((int)0xE19D7318));
            var calls = new Calls { InvokeError = original, ExceptionInfo = new VbeProjectComponents.AccessHelpContextExceptionInfo { Description = new IntPtr(123) } };
            var dispatch = new VbeProjectComponents.AccessHelpContextDispatch(calls);
            Assert.AreSame(original, Assert.ThrowsException<COMException>(() => dispatch.Put(321, () => { })));
            Assert.AreEqual(0, calls.Freed.Count); Assert.IsTrue(original.Data.Contains("AccessHelpContextOutputCleanup"));
            Assert.AreEqual(1, calls.Disposals); Assert.ThrowsException<InvalidOperationException>(() => dispatch.Put(321, () => { })); Assert.AreEqual(1, calls.Invocations);
        }

        [DataTestMethod]
        [DataRow("guid")][DataRow("kind")][DataRow("getter")][DataRow("setter")][DataRow("putrefOrDuplicate")]
        public void OnlyObservedVBProjectInt32TypeContractIsAccepted(string changed)
        {
            Guid actual = new Guid("EEE00915-E393-11D1-BB03-00C04FB6C4A6");
            VbeProjectComponents.RequireAccessHelpContextTypeContract(actual, TYPEKIND.TKIND_DISPATCH, true, true, false);
            Assert.ThrowsException<InvalidOperationException>(() => VbeProjectComponents.RequireAccessHelpContextTypeContract(
                changed == "guid" ? Guid.Empty : actual, changed == "kind" ? TYPEKIND.TKIND_INTERFACE : TYPEKIND.TKIND_DISPATCH,
                changed != "getter", changed != "setter", changed == "putrefOrDuplicate"));
        }
    }

    [TestClass]
    public sealed class AccessHelpContextProjectTests
    {
        public sealed class Project : VbeProjectComponentsTests.FakeProject
        {
            public int Protection { get; set; }
            private int context;
            internal Action OnReadContext;
            public int HelpContextID { get { OnReadContext?.Invoke(); return context; } set { context = value; } }
        }
        public sealed class Window { public long HWnd { get; set; } = 42; }
        public sealed class Vbe
        {
            public List<VbeProjectComponentsTests.FakeProject> VBProjects { get; } = new List<VbeProjectComponentsTests.FakeProject>();
            public Window MainWindow { get; } = new Window();
        }
        private sealed class Fixture
        {
            internal readonly Project Project = new Project { Name = "P" };
            internal readonly Vbe Vbe = new Vbe();
            internal readonly VbeProjectComponents Service;
            internal readonly AccessHelpContextDispatchTests.Calls Calls = new AccessHelpContextDispatchTests.Calls();
            internal Fixture()
            {
                Vbe.VBProjects.Add(Project); Service = new VbeProjectComponents(Vbe, new VbeForms(Vbe));
                Service.AccessHelpContextHost = () => true;
                Service.AccessHelpContextNativeProject = target => true;
                Service.AccessHelpContextIdentity = ReferenceEquals;
                Service.AccessHelpContextFactory = (target, window) => { Assert.AreSame(Project, target); Assert.AreEqual(new IntPtr(42), window); return new VbeProjectComponents.AccessHelpContextDispatch(Calls); };
                Calls.OnInvoke = () => Project.HelpContextID = Calls.Value;
            }
            internal Request Request()
            {
                dynamic state = Service.ProjectProperties("P");
                return new Request { Project = "P", Property = "HelpContextID", Value = 321, ExpectedProjectVersion = state.Version };
            }
        }

        [TestMethod]
        public void OriginalAccessProjectUsesRawWriteAndExactRetentionReadback()
        {
            var fixture = new Fixture(); var request = fixture.Request(); int checks = 0;
            request.RevalidateProjectPropertyAuthorization = validateScope => checks++;
            fixture.Service.SetProjectProperty(request);
            Assert.AreEqual(321, fixture.Project.HelpContextID); Assert.AreEqual(1, fixture.Calls.Invocations);
            Assert.IsTrue(checks >= 2, "Runtime authorization surrounds the final project/version reads.");
        }

        [DataTestMethod]
        [DataRow("revision")][DataRow("identityWithSameVersion")][DataRow("mode")][DataRow("protection")][DataRow("window")][DataRow("host")][DataRow("nativeTarget")]
        public void ChangedContextAfterReadonlyPreparationRefusesBeforeNativeWrite(string changed)
        {
            var fixture = new Fixture(); var request = fixture.Request();
            fixture.Calls.OnPrepare = () => {
                if (changed == "revision") fixture.Project.Description = "Changed";
                if (changed == "identityWithSameVersion") fixture.Vbe.VBProjects[0] = new Project { Name = "P" };
                if (changed == "mode") fixture.Project.Mode = 1;
                if (changed == "protection") fixture.Project.Protection = 1;
                if (changed == "window") fixture.Vbe.MainWindow.HWnd = 43;
                if (changed == "host") fixture.Service.AccessHelpContextHost = () => false;
                if (changed == "nativeTarget") fixture.Service.AccessHelpContextNativeProject = target => false;
            };
            var error = Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.SetProjectProperty(request));
            Assert.AreEqual(0, fixture.Calls.Invocations); Assert.AreEqual(0, error.Data["AccessHelpContextInvokeEntries"]);
            Assert.IsFalse(VbeScalarProperty.FormatFailure(error).Contains("SetterInvocation"));
        }

        [TestMethod]
        public void AuthorizationRevalidatedAfterSecondProjectVersionReadCanRevokeTheWrite()
        {
            var fixture = new Fixture(); var request = fixture.Request(); int checks = 0;
            request.RevalidateProjectPropertyAuthorization = validateScope => { checks++; if (!validateScope) throw new InvalidOperationException("Policy revoked after final reads"); };
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.SetProjectProperty(request));
            Assert.AreEqual(2, checks); Assert.AreEqual(0, fixture.Calls.Invocations);
        }

        [TestMethod]
        public void TargetChangeDuringFirstAuthorizationIsCaughtBySecondIdentityRead()
        {
            var fixture = new Fixture(); var request = fixture.Request();
            request.RevalidateProjectPropertyAuthorization = validateScope => { if (validateScope) fixture.Vbe.VBProjects[0] = new Project { Name = "P" }; };
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.SetProjectProperty(request)); Assert.AreEqual(0, fixture.Calls.Invocations);
        }

        [TestMethod]
        public void RevisionChangedByHostScopeReadIsRecheckedBeforeCachedFinalAuthorization()
        {
            var fixture = new Fixture(); var request = fixture.Request(); var stages = new List<bool>();
            request.RevalidateProjectPropertyAuthorization = validateScope => {
                stages.Add(validateScope);
                if (validateScope) fixture.Project.Description = "Changed while reading host scope";
            };
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.SetProjectProperty(request));
            CollectionAssert.AreEqual(new[] { true }, stages.ToArray()); Assert.AreEqual(0, fixture.Calls.Invocations);
        }

        [TestMethod]
        public void SameVersionReplacementDuringLastRevisionReadIsRefusedByFinalIdentity()
        {
            var fixture = new Fixture(); var request = fixture.Request(); int reads = 0;
            fixture.Calls.OnPrepare = () => fixture.Project.OnReadContext = () => {
                if (++reads == 2) fixture.Vbe.VBProjects[0] = new Project { Name = "P" };
            };
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.SetProjectProperty(request));
            Assert.AreEqual(2, reads); Assert.AreEqual(0, fixture.Calls.Invocations);
        }

        [TestMethod]
        public void SuccessfulNativeReturnWithFailedRetentionDoesNotTryAnotherSetter()
        {
            var fixture = new Fixture(); var request = fixture.Request(); fixture.Calls.OnInvoke = () => { };
            var error = Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.SetProjectProperty(request));
            StringAssert.Contains(VbeScalarProperty.FormatFailure(error), "RetentionReadback"); Assert.AreEqual(1, fixture.Calls.Invocations);
        }

        [DataTestMethod]
        [DataRow(false)][DataRow(true)]
        public void ManagedAndOtherHostProjectsKeepExistingDescriptorPath(bool otherHost)
        {
            var fixture = new Fixture(); var request = fixture.Request();
            if (otherHost) fixture.Service.AccessHelpContextHost = () => false;
            else fixture.Service.AccessHelpContextNativeProject = target => false;
            fixture.Service.SetProjectProperty(request); Assert.AreEqual(321, fixture.Project.HelpContextID); Assert.AreEqual(0, fixture.Calls.Preparations);
        }

        [TestMethod]
        public void OtherProjectScalarsAreNotSentToTheAccessInt32Route()
        {
            var fixture = new Fixture(); var request = fixture.Request(); request.Property = "Description"; request.Value = "New description";
            fixture.Service.SetProjectProperty(request); Assert.AreEqual("New description", fixture.Project.Description); Assert.AreEqual(0, fixture.Calls.Preparations);
        }

        [TestMethod]
        public void Int32OverflowIsRefusedBeforeNativePreparation()
        {
            var fixture = new Fixture(); var request = fixture.Request(); request.Value = long.MaxValue;
            Assert.ThrowsException<OverflowException>(() => fixture.Service.SetProjectProperty(request)); Assert.AreEqual(0, fixture.Calls.Preparations);
        }
    }
}
