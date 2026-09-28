using System;
using System.Threading;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
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
                    if (scenario == "ParamArray") signature = signature.Replace("ByVal values As Variant", "ParamArray values() As Variant");
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
    }
}
