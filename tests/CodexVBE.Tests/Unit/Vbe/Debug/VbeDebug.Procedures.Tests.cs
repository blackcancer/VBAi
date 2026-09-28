namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Globalization;
    using System.Threading;
    using System.Threading.Tasks;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class VbeDebugTests
    {
        [TestMethod]
        public void ProcedureCallEncodesScalarsAndRefusesStaleOrUnqualifiedTargets()
        {
            var f = Create(2); f.Project.VBComponents[0].Type = 1;
            f.Module.Code = "Public Sub TryMe(ByVal text As String, ByVal count As Double)\r\nEnd Sub";
            var request = Location(f); request.ExpectedSha256 = Sha(f.Module.Code); request.Procedure = "TryMe";
            request.Arguments = new object[] { "x\": Stop: '\"", 1.5, true, null };
            string command = f.Service.PrepareProcedureCall(request);
            Assert.AreEqual("Call VBAProject.Module1.TryMe(\"x\"\": Stop: '\"\"\", 1.5, True, Null)", command);
            request.ExpectedSha256 = "stale";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.PrepareProcedureCall(request));
            request.ExpectedSha256 = Sha(f.Module.Code); f.Project.Mode = 1;
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.PrepareProcedureCall(request));
            f.Project.Mode = 2; f.Module.Code = "Private Sub TryMe()\r\nEnd Sub"; request.ExpectedSha256 = Sha(f.Module.Code);
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.PrepareProcedureCall(request));
            foreach (object bad in new object[] { "bad\nline", double.NaN, double.PositiveInfinity, new object(), new int[] { 1 } })
                Assert.ThrowsException<ArgumentException>(() => VbeDebug.ProcedureLiteral(bad));
            var culture = CultureInfo.CurrentCulture;
            try { CultureInfo.CurrentCulture = new CultureInfo("fr-FR"); Assert.AreEqual("1.5", VbeDebug.ProcedureLiteral(1.5m)); }
            finally { CultureInfo.CurrentCulture = culture; }
        }
        [TestMethod]
        public void ParameterizedRunCapturesRequestAndTracksDeliveryWithoutClaimingRuntimeSuccess()
        {
            var previous = SynchronizationContext.Current; var context = new RecordingContext();
            SynchronizationContext.SetSynchronizationContext(context);
            try
            {
                var f = Create(2); f.Project.VBComponents[0].Type = 1;
                f.Module.Code = "Public Function TryMe(ByVal count As Long) As Long\r\nTryMe = count\r\nEnd Function";
                var request = Location(f); request.Procedure = "TryMe"; request.ExpectedSha256 = Sha(f.Module.Code); request.Arguments = new object[] { 7 };
                int delivered = 0; string actual = null;
                f.Service.ShowProcedureImmediate = () => { };
                f.Service.ExecuteProcedureCall = command => { delivered++; actual = command; return Task.FromResult<object>(new { TextAfter = "7" }); };
                dynamic queued = f.Service.RunProcedure(request);
                Assert.IsTrue((bool)queued.Pending);
                Assert.ThrowsException<InvalidOperationException>(() => f.Service.RunProcedure(request));
                request.Arguments[0] = 99; request.Procedure = "Different";
                context.RunAll();
                dynamic status = f.Service.ProcedureRunStatus(new Request { Project = f.Project.Name, Query = (string)queued.Query });
                Assert.AreEqual("Delivered", (string)status.State); Assert.IsFalse((bool)status.RuntimeSuccessVerified);
                Assert.AreEqual("? VBAProject.Module1.TryMe(7)", actual); Assert.AreEqual(1, delivered);
                Assert.ThrowsException<InvalidOperationException>(() => f.Service.ProcedureRunStatus(new Request { Project = "Other", Query = (string)queued.Query }));
            }
            finally { SynchronizationContext.SetSynchronizationContext(previous); }
        }
        [TestMethod]
        public void ScheduledProcedureRefusesInterveningCodeChangeBeforeNativeDelivery()
        {
            var previous = SynchronizationContext.Current; var context = new RecordingContext(); SynchronizationContext.SetSynchronizationContext(context);
            try
            {
                var f = Create(2); f.Project.VBComponents[0].Type = 1; f.Module.Code = "Public Sub TryMe()\r\nEnd Sub";
                var request = Location(f); request.Procedure = "TryMe"; request.ExpectedSha256 = Sha(f.Module.Code); request.Arguments = new object[0];
                int calls = 0; f.Service.ExecuteProcedureCall = command => { calls++; return Task.FromResult<object>(null); };
                dynamic queued = f.Service.RunProcedure(request); f.Module.Code += "\r\n' changed"; context.RunAll();
                dynamic status = f.Service.ProcedureRunStatus(new Request { Project = f.Project.Name, Query = (string)queued.Query });
                Assert.AreEqual("Failed", (string)status.State); Assert.AreEqual(0, calls);
                StringAssert.Contains((string)status.Error, "changed");
            }
            finally { SynchronizationContext.SetSynchronizationContext(previous); }
        }
    }
}
