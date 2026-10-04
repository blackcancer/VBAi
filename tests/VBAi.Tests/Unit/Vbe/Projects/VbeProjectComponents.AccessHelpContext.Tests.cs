using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi;

namespace VBAi.Tests.Unit
{
    [TestClass]
    public sealed class AccessHelpContextDispatchTests
    {
        internal sealed class Calls : VbeProjectComponents.IAccessHelpContextCalls
        {
            internal int Preparations, Invocations, Owners, Disposals, Value;
            internal int RefuseOwnerAt;
            internal Action OnPrepare, OnInvoke;
            internal Exception PrepareError, InvokeError, DisposeError;
            internal readonly List<string> Events = new List<string>();
            public void Prepare()
            {
                Preparations++; Events.Add("prepare"); OnPrepare?.Invoke();
                if (PrepareError != null) throw PrepareError;
            }
            public void RequireOwner()
            {
                Owners++; Events.Add("owner");
                if (Owners == RefuseOwnerAt) throw new InvalidOperationException("Original native owner or canonical identity changed");
            }
            public void Set(int value)
            {
                Invocations++; Value = value; Events.Add("set"); OnInvoke?.Invoke();
                if (InvokeError != null) throw InvokeError;
            }
            public void Dispose() { Disposals++; Events.Add("dispose"); if (DisposeError != null) throw DisposeError; }
        }

        [TestMethod]
        public void OfficialImportedProjectDeclaresInt32SetterWithHresultTranslation()
        {
            Type contract = VbeProjectComponents.AccessHelpContextInterfaceType;
            Assert.AreEqual(new Guid("EEE00915-E393-11D1-BB03-00C04FB6C4A6"), contract.GUID);
            Assert.IsTrue(contract.IsImport); Assert.IsTrue(contract.IsInterface);
            MethodInfo setter = contract.GetMethod("set_HelpContextID");
            Assert.IsNotNull(setter); Assert.AreEqual(typeof(void), setter.ReturnType);
            Assert.AreEqual(1, setter.GetParameters().Length);
            Assert.AreEqual(typeof(int), setter.GetParameters()[0].ParameterType);
            Assert.IsTrue(setter.GetParameters()[0].IsIn);
            Assert.AreEqual(117, ((DispIdAttribute)Attribute.GetCustomAttribute(setter, typeof(DispIdAttribute))).Value);
            Assert.AreEqual((MethodImplAttributes)0, setter.GetMethodImplementationFlags() & MethodImplAttributes.PreserveSig,
                "The imported Void setter must translate a failing native HRESULT into an exception.");
        }

        [DataTestMethod]
        [DataRow(0)][DataRow(321)][DataRow(int.MinValue)][DataRow(int.MaxValue)]
        public void TypedInt32SetterReceivesExactValueAfterPreparationAndFinalGuard(int value)
        {
            var calls = new Calls(); var dispatch = new VbeProjectComponents.AccessHelpContextDispatch(calls); int guards = 0;
            dispatch.Put(value, () => { guards++; calls.Events.Add("authorize"); Assert.AreEqual(1, calls.Preparations); Assert.AreEqual(0, calls.Invocations); });
            CollectionAssert.AreEqual(new[] { "owner", "prepare", "authorize", "owner", "set", "dispose" }, calls.Events.ToArray());
            Assert.AreEqual(value, calls.Value); Assert.AreEqual(1, guards); Assert.AreEqual(1, dispatch.InvokeEntries);
            Assert.AreEqual(1, calls.Invocations); Assert.AreEqual(1, calls.Disposals);
            Assert.ThrowsException<InvalidOperationException>(() => dispatch.Put(value, () => { }));
            Assert.AreEqual(1, calls.Invocations); Assert.AreEqual(1, calls.Disposals);
        }

        [DataTestMethod]
        [DataRow("cast")][DataRow("preparation")][DataRow("initialOwner")][DataRow("finalOwner")][DataRow("authorization")]
        public void ReadonlyPreparationAndFinalRefusalsNeverSetOrReplay(string refusal)
        {
            var calls = new Calls();
            if (refusal == "cast") calls.PrepareError = new InvalidCastException("The original project does not expose the imported interface");
            if (refusal == "preparation") calls.PrepareError = new COMException("Canonical QI preparation failed", unchecked((int)0x80004002));
            if (refusal == "initialOwner") calls.RefuseOwnerAt = 1;
            if (refusal == "finalOwner") calls.RefuseOwnerAt = 2;
            var dispatch = new VbeProjectComponents.AccessHelpContextDispatch(calls); Exception observed = null;
            try { dispatch.Put(321, () => { if (refusal == "authorization") throw new InvalidOperationException("Revoked"); }); }
            catch (Exception error) { observed = error; }
            Assert.IsNotNull(observed);
            if (calls.PrepareError != null) Assert.AreSame(calls.PrepareError, observed);
            Assert.AreEqual(0, dispatch.InvokeEntries); Assert.AreEqual(0, calls.Invocations); Assert.AreEqual(1, calls.Disposals);
            int preparations = calls.Preparations;
            Assert.ThrowsException<InvalidOperationException>(() => dispatch.Put(321, () => { }));
            Assert.AreEqual(preparations, calls.Preparations); Assert.AreEqual(1, calls.Disposals);
        }

        [TestMethod]
        public void MissingFinalAuthorizationConsumesTheWriteWithoutPreparationOrSetter()
        {
            var calls = new Calls(); var dispatch = new VbeProjectComponents.AccessHelpContextDispatch(calls);
            Assert.ThrowsException<ArgumentNullException>(() => dispatch.Put(321, null));
            Assert.AreEqual(0, calls.Owners); Assert.AreEqual(0, calls.Preparations); Assert.AreEqual(0, calls.Invocations);
            Assert.AreEqual(0, dispatch.InvokeEntries); Assert.AreEqual(1, calls.Disposals);
            Assert.ThrowsException<InvalidOperationException>(() => dispatch.Put(321, () => { }));
        }

        [DataTestMethod]
        [DataRow(false)][DataRow(true)]
        public void FailedTypedSetterKeepsOriginalHresultEvenAfterPartialMutationOrCleanupFailure(bool partialMutation)
        {
            const int hr = unchecked((int)0x80020009);
            var original = new COMException("Original setter failure", hr); bool changed = false;
            var calls = new Calls { InvokeError = original, DisposeError = new InvalidOperationException("cleanup") };
            calls.OnInvoke = () => changed = partialMutation;
            var dispatch = new VbeProjectComponents.AccessHelpContextDispatch(calls);
            Assert.AreSame(original, Assert.ThrowsException<COMException>(() => dispatch.Put(321, () => { })));
            Assert.AreEqual(hr, original.HResult); Assert.AreEqual(partialMutation, changed);
            Assert.IsTrue(original.Data.Contains("AccessHelpContextCleanupFailure")); Assert.AreEqual(1, dispatch.InvokeEntries);
            Assert.ThrowsException<InvalidOperationException>(() => dispatch.Put(321, () => { })); Assert.AreEqual(1, calls.Invocations);
        }

        [TestMethod]
        public void CleanupFailureAfterSetterReturnRemainsUncertainAndCannotReplay()
        {
            var original = new InvalidOperationException("Identity reference cleanup failed");
            var calls = new Calls { DisposeError = original }; var dispatch = new VbeProjectComponents.AccessHelpContextDispatch(calls);
            Assert.AreSame(original, Assert.ThrowsException<InvalidOperationException>(() => dispatch.Put(321, () => { })));
            Assert.AreEqual(1, dispatch.InvokeEntries); Assert.AreEqual(1, calls.Invocations); Assert.AreEqual(1, calls.Disposals);
            Assert.ThrowsException<InvalidOperationException>(() => dispatch.Put(321, () => { })); Assert.AreEqual(1, calls.Disposals);
        }

        [TestMethod]
        public void PreparationFailureKeepsItsOriginalErrorWhenCleanupAlsoFails()
        {
            var original = new InvalidCastException("Unsupported original project interface");
            var calls = new Calls { PrepareError = original, DisposeError = new COMException("cleanup") };
            var dispatch = new VbeProjectComponents.AccessHelpContextDispatch(calls);
            Assert.AreSame(original, Assert.ThrowsException<InvalidCastException>(() => dispatch.Put(321, () => { })));
            Assert.IsTrue(original.Data.Contains("AccessHelpContextCleanupFailure"));
            Assert.AreEqual(0, dispatch.InvokeEntries); Assert.AreEqual(0, calls.Invocations); Assert.AreEqual(1, calls.Disposals);
        }
    }

    [TestClass]
    public sealed class AccessHelpContextProjectTests
    {
        [DataTestMethod]
        [DataRow("MSACCESS", true)][DataRow("msaccess", true)]
        [DataRow("MSPUB", true)][DataRow("mspub", true)]
        [DataRow("EXCEL", false)][DataRow("WINWORD", false)]
        [DataRow("POWERPNT", false)][DataRow("OUTLOOK", false)]
        [DataRow("SLDWORKS", false)][DataRow("MSPUB.exe", false)]
        [DataRow("", false)][DataRow(null, false)]
        public void OnlyAccessAndPublisherNativeProjectsUseTheTypedMetadataRoute(string processName, bool expected)
        {
            Assert.AreEqual(expected, VbeProjectComponents.IsAccessHelpContextHost(processName));
        }

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
        public void OriginalAccessProjectUsesTypedWriteAndExactRetentionReadback()
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

        [DataTestMethod]
        [DataRow("321")][DataRow(321L)][DataRow(321.0)]
        public void ExistingScalarConversionFeedsTheTypedInt32Setter(object value)
        {
            var fixture = new Fixture(); var request = fixture.Request(); request.Value = value;
            fixture.Service.SetProjectProperty(request);
            Assert.AreEqual(321, fixture.Calls.Value); Assert.AreEqual(1, fixture.Calls.Invocations);
            Assert.AreEqual(321, fixture.Project.HelpContextID);
        }

        [TestMethod]
        public void FailedTypedProjectSetterRetainsPartialMutationAsFailureWithoutReadbackOrFallback()
        {
            var fixture = new Fixture(); var request = fixture.Request();
            var original = new COMException("Failed declared setter after mutation", unchecked((int)0x80020009));
            fixture.Calls.InvokeError = original;
            int postMutationReads = 0;
            fixture.Calls.OnInvoke = () => {
                fixture.Project.HelpContextID = fixture.Calls.Value;
                fixture.Project.OnReadContext = () => postMutationReads++;
            };
            Assert.AreSame(original, Assert.ThrowsException<COMException>(() => fixture.Service.SetProjectProperty(request)));
            Assert.AreEqual(1, original.Data["AccessHelpContextInvokeEntries"]);
            StringAssert.Contains(VbeScalarProperty.FormatFailure(original), "SetterInvocation");
            Assert.AreEqual(0, postMutationReads); Assert.AreEqual(1, fixture.Calls.Invocations);
            fixture.Project.OnReadContext = null;
            Assert.AreEqual(321, fixture.Project.HelpContextID, "A changed property does not override the failed native outcome.");
        }
    }
}
