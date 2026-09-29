using System;
using System.Threading;
using VBAi;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VBAi.Tests.Unit
{
    public sealed partial class VbeDebugTests
    {
        /// <summary>Toute signature non déterministe de la matrice refuse la mise en file native.</summary>
        [TestMethod]
        public void ValuesSignatureMatrixRefusesByRefObjectsTypedArraysConditionalAndUnknownBinding()
        {
            var previous = SynchronizationContext.Current; var context = new NativeNavigationContext(); SynchronizationContext.SetSynchronizationContext(context);
            try
            {
                foreach (string scenario in ValueSignatureFailures)
                {
                    var fixture = Create(2); var host = new ValuesHost(); fixture.Service.ProcedureValuesHost = host;
                    string signature = "Public Function TryMe(ByVal values As Variant) As Variant";
                    if (scenario == "ByRef") signature = signature.Replace("ByVal", "ByRef");
                    if (scenario == "implicit ByRef") signature = signature.Replace("ByVal ", "");
                    if (scenario == "typed array") signature = signature.Replace("values As Variant", "values() As Variant");
                    if (scenario == "object") signature = signature.Replace("values As Variant", "values As Object");
                    if (scenario == "invalid ParamArray") signature = signature.Replace("ByVal values As Variant", "ParamArray values() As Long");
                    if (scenario == "private") signature = signature.Replace("Public", "Private");
                    if (scenario == "property") signature = signature.Replace("Function", "Property Get");
                    if (scenario == "return object") signature = signature.Replace(") As Variant", ") As Object");
                    var request = ValuesRequest(fixture, signature);
                    if (scenario == "conditional") { fixture.Module.Code += "\r\n#Const Flag=True"; request.ExpectedSha256 = Sha(fixture.Module.Code); }
                    if (scenario == "missing required") request.Arguments = new object[0];
                    if (scenario == "unknown named") request.ArgumentNames = new[] { "missing" };
                    if (scenario == "duplicate named") { request.Arguments = new object[] { 1, 2 }; request.ArgumentNames = new[] { "values", "VALUES" }; }
                    if (scenario == "missing required" || scenario == "unknown named" || scenario == "duplicate named")
                        Assert.ThrowsException<ArgumentException>(() => fixture.Service.RunProcedureValues(request), scenario);
                    else Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.RunProcedureValues(request), scenario);
                    Assert.AreEqual(0, host.Invocations);
                }
            }
            finally { SynchronizationContext.SetSynchronizationContext(previous); }
        }

        /// <summary>Une invocation reçoit une copie profonde et le statut relit le retour sans réexécuter de macro.</summary>
        [TestMethod]
        public void ValuesQueueCapturesArraysReturnsDataAndSharesTheExistingPendingGuard()
        {
            var previous = SynchronizationContext.Current; var context = new NativeNavigationContext(); SynchronizationContext.SetSynchronizationContext(context);
            try
            {
                var fixture = Create(2); var request = ValuesRequest(fixture); var host = new ValuesHost(); fixture.Service.ProcedureValuesHost = host;
                dynamic queued = fixture.Service.RunProcedureValues(request);
                Assert.IsTrue((bool)queued.Pending); Assert.IsFalse((bool)queued.InvocationInvoked);
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.RunProcedureValues(request));
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.RunProcedure(request));
                ((object[])request.Arguments[0])[0] = 99; request.Procedure = "changed"; context.RunAll();
                dynamic status = fixture.Service.ProcedureValuesStatus(new Request { Project = fixture.Project.Name, Query = queued.Query });
                Assert.AreEqual("Returned", (string)status.State); Assert.IsTrue((bool)status.ReturnValueVerified); Assert.IsFalse((bool)status.RuntimeSuccessVerified);
                Assert.AreEqual(1, ((object[])host.Received[0])[0]); Assert.AreEqual(1, host.Invocations);
                for (int i = 0; i < 3; i++) fixture.Service.ProcedureValuesStatus(new Request { Project = fixture.Project.Name, Query = queued.Query });
                Assert.AreEqual(1, host.Invocations);
                Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.ProcedureValuesStatus(new Request { Project = "Other", Query = queued.Query }));
            }
            finally { SynchronizationContext.SetSynchronizationContext(previous); }
        }

        /// <summary>Versions, chemin, mode, classe et transport d'un autre PID/identité refusent avant Invoke.</summary>
        [TestMethod]
        public void ValuesQueueRechecksEveryIdentitySourcePathAndModeBoundaryBeforeInvocation()
        {
            var previous = SynchronizationContext.Current; var context = new NativeNavigationContext(); SynchronizationContext.SetSynchronizationContext(context);
            try
            {
                foreach (string scenario in ValueQueueFailures)
                {
                    var fixture = Create(2); var request = ValuesRequest(fixture); var host = new ValuesHost(); fixture.Service.ProcedureValuesHost = host;
                    if (scenario == "stale source") request.ExpectedSha256 = "stale";
                    if (scenario == "wrong path") request.ExpectedHostPath = @"C:\Temp\Other.xlsm";
                    if (scenario == "missing path") request.ExpectedHostPath = null;
                    if (scenario == "runtime mode") fixture.Project.Mode = 1;
                    if (scenario == "class target") fixture.Project.VBComponents[0].Type = 2;
                    if (scenario == "host identity") host.RejectIdentity = true;
                    if (scenario.StartsWith("changed", StringComparison.Ordinal))
                    {
                        dynamic queued = fixture.Service.RunProcedureValues(request);
                        if (scenario == "changed source") fixture.Module.Code += "\r\n' changed";
                        if (scenario == "changed path") fixture.Project.FileName = @"C:\Temp\Other.xlsm";
                        if (scenario == "changed mode") fixture.Project.Mode = 1;
                        context.RunAll();
                        dynamic status = fixture.Service.ProcedureValuesStatus(new Request { Project = request.Project, Query = queued.Query });
                        Assert.AreEqual("Failed", (string)status.State); Assert.IsFalse((bool)status.InvocationInvoked);
                    }
                    else if (scenario == "missing path") Assert.ThrowsException<ArgumentException>(() => fixture.Service.RunProcedureValues(request));
                    else Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.RunProcedureValues(request));
                    Assert.AreEqual(0, host.Invocations, scenario);
                }
            }
            finally { SynchronizationContext.SetSynchronizationContext(previous); }
        }

        /// <summary>Un retour objet est rejeté après l'unique invocation; une erreur native garde l'incertitude des effets de bord.</summary>
        [TestMethod]
        public void UnsupportedReturnsAndNativeFailuresNeverTriggerAnotherInvocation()
        {
            var previous = SynchronizationContext.Current; var context = new NativeNavigationContext(); SynchronizationContext.SetSynchronizationContext(context);
            try
            {
                foreach (bool error in new[] { false, true })
                {
                    var fixture = Create(2); var request = ValuesRequest(fixture); var host = new ValuesHost { Return = new object() }; fixture.Service.ProcedureValuesHost = host;
                    if (error) host.OnInvoke = () => { throw new InvalidOperationException("native execution failed after entry"); };
                    dynamic queued = fixture.Service.RunProcedureValues(request); context.RunAll();
                    dynamic status = fixture.Service.ProcedureValuesStatus(new Request { Project = request.Project, Query = queued.Query });
                    Assert.AreEqual(error ? "Failed" : "ReturnRejected", (string)status.State); Assert.IsTrue((bool)status.InvocationInvoked);
                    Assert.IsFalse((bool)status.ReturnValueVerified); Assert.AreEqual(error, (bool)status.Uncertain); Assert.AreEqual(1, host.Invocations);
                }
            }
            finally { SynchronizationContext.SetSynchronizationContext(previous); }
        }

        [STATestMethod]
        public void NativeValuesTransportVerifiesActualWindowAndComIdentityBeforeManagedDispatch()
        {
            var native = new VbeDebug.NativeProcedureValuesHost();
            Assert.AreEqual(System.Diagnostics.Process.GetCurrentProcess().ProcessName, native.ReadProcessName());
            Assert.ThrowsException<InvalidOperationException>(() => native.ResolveTarget(new object(), @"C:\Temp\fixture.xlsm"));
            foreach (string failure in new[] { "owner", "absent", "duplicate", "path", "unsaved", "relative", "null-path", "limit", "valid" })
            using (var f = new VBAi.Tests.Infrastructure.NativeProcedureValuesFixture())
            {
                if (failure == "owner") f.Application.Hwnd = 0;
                if (failure == "absent") f.Workbook.VBProject = f.OtherProject;
                if (failure == "duplicate") f.Application.Workbooks.Add(f.Workbook);
                if (failure == "path") f.Workbook.FullName = @"C:\Temp\Other.xlsm";
                if (failure == "unsaved") f.Workbook.Path = " ";
                if (failure == "relative") f.Workbook.FullName = "fixture.xlsm";
                if (failure == "null-path") f.Workbook.FullName = null;
                if (failure == "limit") { f.Application.Workbooks.Clear(); for (int i = 0; i < 1001; i++) f.Application.Workbooks.Add(new VBAi.Tests.Infrastructure.NativeProcedureValuesFixture.WorkbookContract { VBProject = f.OtherProject }); }
                if (failure != "valid") { Assert.ThrowsException<InvalidOperationException>(() => f.Resolve(), failure); Assert.AreEqual(0, f.Application.Invocations); }
                else
                {
                    f.Application.Workbooks.Insert(0, new VBAi.Tests.Infrastructure.NativeProcedureValuesFixture.WorkbookContract { VBProject = f.OtherProject });
                    object target = f.Resolve(); object[] array = { 3, "quoted\" value" };
                    Assert.AreEqual("Native dispatch contract return", f.Host.Invoke(target, "Module1", "TryMe", new object[] { array, 7 }));
                    Assert.AreEqual("'C:\\Temp\\Owner''s.xlsm'!Module1.TryMe", f.Application.Macro); Assert.AreSame(array, f.Application.Arguments[0]); Assert.AreEqual(7, f.Application.Arguments[1]);
                    Assert.AreEqual(1, f.Application.Invocations);
                    f.Workbook.FullName = @"C:\Temp\Changed.xlsm";
                    Assert.ThrowsException<InvalidOperationException>(() => f.Host.Invoke(target, "Module1", "TryMe", new object[0]));
                    Assert.AreEqual(1, f.Application.Invocations);
                }
                Assert.ThrowsException<InvalidOperationException>(() => f.Host.Invoke(new object(), "Module1", "TryMe", new object[0]));
            }
        }

        [STATestMethod]
        public void ProcedureProjectIdentityReleasesEveryAcquiredIUnknownOnSuccessAndDisposedWrappers()
        {
            object first = Activator.CreateInstance(Type.GetTypeFromProgID("StdFont", true)), second = Activator.CreateInstance(Type.GetTypeFromProgID("StdFont", true));
            object released = Activator.CreateInstance(Type.GetTypeFromProgID("StdFont", true)); System.Runtime.InteropServices.Marshal.FinalReleaseComObject(released);
            try
            {
                Assert.IsFalse(VbeDebug.NativeProcedureValuesHost.SameComIdentity(null, first));
                Assert.IsFalse(VbeDebug.NativeProcedureValuesHost.SameComIdentity(first, null));
                Assert.IsFalse(VbeDebug.NativeProcedureValuesHost.SameComIdentity(new object(), first));
                Assert.IsFalse(VbeDebug.NativeProcedureValuesHost.SameComIdentity(first, new object()));
                Assert.IsTrue(VbeDebug.NativeProcedureValuesHost.SameComIdentity(first, first));
                Assert.IsFalse(VbeDebug.NativeProcedureValuesHost.SameComIdentity(first, second));
                Assert.ThrowsException<System.Runtime.InteropServices.InvalidComObjectException>(() => VbeDebug.NativeProcedureValuesHost.SameComIdentity(released, first));
                Assert.ThrowsException<System.Runtime.InteropServices.InvalidComObjectException>(() => VbeDebug.NativeProcedureValuesHost.SameComIdentity(first, released));
            }
            finally { System.Runtime.InteropServices.Marshal.FinalReleaseComObject(first); System.Runtime.InteropServices.Marshal.FinalReleaseComObject(second); }
        }

        [TestMethod]
        public void ValuesPreparationCoversAllPathBodyModuleAndLateMutationGuards()
        {
            var previous = SynchronizationContext.Current; var context = new NativeNavigationContext(); SynchronizationContext.SetSynchronizationContext(context);
            try
            {
                foreach (string fault in new[] { "mode", "module-absent", "module-duplicate", "sha-missing", "body-zero", "body-after", "source-empty", "relative-project", "relative-path", "unsafe-path", "extension", "late-mode", "late-path", "late-source", "late-empty" })
                {
                    var f = Create(2); var request = ValuesRequest(f); var host = new ValuesHost(); f.Service.ProcedureValuesHost = host;
                    if (fault == "mode") request.ExpectedMode = 1;
                    if (fault == "module-absent") request.Module = "Missing";
                    if (fault == "module-duplicate") f.Project.VBComponents.Add(f.Project.VBComponents[0]);
                    if (fault == "sha-missing") request.ExpectedSha256 = null;
                    if (fault == "body-zero") f.Module.ProcBodyLine.Body = 0;
                    if (fault == "body-after") f.Module.ProcBodyLine.Body = 100;
                    if (fault == "source-empty") { f.Module.Code = ""; request.ExpectedSha256 = Sha(""); }
                    if (fault == "relative-project") f.Project.FileName = "fixture.xlsm";
                    if (fault == "relative-path") request.ExpectedHostPath = "fixture.xlsm";
                    if (fault == "unsafe-path") request.ExpectedHostPath = @"C:\Temp\[fixture].xlsm";
                    if (fault == "extension") request.ExpectedHostPath = @"C:\Temp\fixture.xlsx";
                    host.OnResolve = () => { if (fault == "late-mode") f.Project.Mode = 0; if (fault == "late-path") f.Project.FileName = @"C:\Temp\Other.xlsm"; if (fault == "late-source") f.Module.Code += "\r\n' late"; if (fault == "late-empty") f.Module.Code = ""; };
                    if (fault == "mode" || fault == "relative-path" || fault == "unsafe-path" || fault == "extension") Assert.ThrowsException<ArgumentException>(() => f.Service.RunProcedureValues(request), fault);
                    else Assert.ThrowsException<InvalidOperationException>(() => f.Service.RunProcedureValues(request), fault);
                    Assert.AreEqual(0, host.Invocations);
                }
                Assert.ThrowsException<ArgumentNullException>(() => Create().Service.RunProcedureValues(null));
                Assert.ThrowsException<ArgumentNullException>(() => Create().Service.ProcedureValuesStatus(null));
                SynchronizationContext.SetSynchronizationContext(null);
                Assert.ThrowsException<InvalidOperationException>(() => Create().Service.RunProcedureValues(new Request()));
            }
            finally { SynchronizationContext.SetSynchronizationContext(previous); }
        }

        [TestMethod]
        public void ValuesOperationsCoverNamedCaptureDeliveringPostFailureIdentityChangeAndMixedExpiry()
        {
            var previous = SynchronizationContext.Current; var context = new NativeNavigationContext(); SynchronizationContext.SetSynchronizationContext(context);
            try
            {
                var f = Create(2); var request = ValuesRequest(f); var host = new ValuesHost(); f.Service.ProcedureValuesHost = host;
                request.ArgumentNames = new[] { "values" };
                host.OnInvoke = () => Assert.ThrowsException<InvalidOperationException>(() => f.Service.RunProcedureValues(request));
                dynamic first = f.Service.RunProcedureValues(request); string firstId = first.Query;
                request.ArgumentNames[0] = "changed"; context.RunAll(); Assert.AreEqual(1, host.Invocations);
                var immediate = ProcedureRequest(f); f.Service.ShowProcedureImmediate = () => { }; f.Service.ExecuteProcedureCall = command => System.Threading.Tasks.Task.FromResult<object>(null);
                for (int i = 0; i < 20; i++) { f.Service.RunProcedure(immediate); context.RunAll(); }
                Assert.ThrowsException<InvalidOperationException>(() => f.Service.ProcedureValuesStatus(new Request { Project = f.Project.Name, Query = firstId }));
                request = ValuesRequest(f); host.OnInvoke = null;
                for (int i = 0; i < 22; i++) { f.Service.RunProcedureValues(request); context.RunAll(); }
                context.RejectPost = context.Posts + 1;
                dynamic rejected = f.Service.RunProcedureValues(request); Assert.AreEqual("Failed", (string)rejected.State); Assert.IsFalse((bool)rejected.InvocationInvoked);
                context.RejectPost = 0;
                dynamic changed = f.Service.RunProcedureValues(request); f.Project.VBComponents[0].Name = "MODULE1"; context.RunAll();
                dynamic status = f.Service.ProcedureValuesStatus(new Request { Project = f.Project.Name, Query = changed.Query });
                Assert.AreEqual("Failed", (string)status.State); Assert.IsFalse((bool)status.InvocationInvoked);
            }
            finally { SynchronizationContext.SetSynchronizationContext(previous); }
        }
    }
}
