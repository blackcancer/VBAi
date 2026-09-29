namespace VBAi.Tests.Unit
{
    using System;
    using System.Globalization;
    using System.Threading;
    using System.Threading.Tasks;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class VbeDebugTests
    {
        private static Request ProcedureRequest(Fixture f)
        {
            f.Project.VBComponents[0].Type = 1;
            f.Module.Code = "Public Sub TryMe()\r\nEnd Sub";
            var r = Location(f); r.ExpectedSha256 = Sha(f.Module.Code); r.Procedure = "TryMe"; r.Arguments = new object[0]; return r;
        }
        /// <summary>Vérifie l'association des noms à la signature, les omissions et toutes les gardes avant livraison.</summary>
        [TestMethod]
        public void NamedArgumentsBindOnlyToUniqueParametersInTheInspectedSignature()
        {
            var fixture = Create(2); var request = ProcedureRequest(fixture);
            const string source = "\r\nPublic Sub TryMe(ByVal first As Long, _\r\nOptional ByVal second As String = \"default\")\r\nDim local As Long\r\nEnd Sub\r\nSub Other(ByVal first As Long)\r\nEnd Sub";
            fixture.Module.Code = source; fixture.Module.ProcBodyLine.Body = 2; request.ExpectedSha256 = Sha(source);
            request.Arguments = new object[] { "literal", 7 }; request.ArgumentNames = new[] { "SECOND", "first" };
            Assert.AreEqual("Call VBAProject.Module1.TryMe(second:=\"literal\", first:=7)", fixture.Service.PrepareProcedureCall(request));
            request.Arguments = new object[] { 7 }; request.ArgumentNames = new[] { "first" };
            Assert.AreEqual("Call VBAProject.Module1.TryMe(first:=7)", fixture.Service.PrepareProcedureCall(request));
            foreach (string[] names in new[] { new string[0], new[] { "first", "second" }, new[] { "" }, new[] { (string)null },
                new[] { "unknown" }, new[] { "local" }, new[] { "first):End" } })
            { request.ArgumentNames = names; Assert.ThrowsException<ArgumentException>(() => fixture.Service.PrepareProcedureCall(request)); }
            request.Arguments = new object[] { 1, 2 }; request.ArgumentNames = new[] { "first", "FIRST" };
            Assert.ThrowsException<ArgumentException>(() => fixture.Service.PrepareProcedureCall(request));
            request.Arguments = new object[0]; request.ArgumentNames = new string[0];
            Assert.AreEqual("Call VBAProject.Module1.TryMe()", fixture.Service.PrepareProcedureCall(request));
            fixture.Module.Code = "Public Sub TryMe(ParamArray values() As Variant)\r\nEnd Sub"; fixture.Module.ProcBodyLine.Body = 1;
            request.ExpectedSha256 = Sha(fixture.Module.Code); request.Arguments = new object[] { 1 }; request.ArgumentNames = new[] { "values" };
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.PrepareProcedureCall(request));
            fixture.Module.Code = "#If VBA7 Then\r\nPublic Sub TryMe(ByVal first As Long)\r\nEnd Sub\r\n#End If"; fixture.Module.ProcBodyLine.Body = 2;
            request.ExpectedSha256 = Sha(fixture.Module.Code); request.ArgumentNames = new[] { "first" };
            Assert.ThrowsException<InvalidOperationException>(() => fixture.Service.PrepareProcedureCall(request));
        }

        /// <summary>Capture les tableaux de noms et de valeurs avant le retour de l'appel asynchrone.</summary>
        [TestMethod]
        public void QueuedNamedArgumentsCannotBeChangedByMutatingTheOriginalRequest()
        {
            var previous = SynchronizationContext.Current; var context = new NativeNavigationContext();
            SynchronizationContext.SetSynchronizationContext(context);
            try
            {
                var fixture = Create(2); var request = ProcedureRequest(fixture);
                fixture.Module.Code = "Public Sub TryMe(Optional ByVal first As Long = 4)\r\nEnd Sub"; request.ExpectedSha256 = Sha(fixture.Module.Code);
                request.Arguments = new object[] { 7 }; request.ArgumentNames = new[] { "first" };
                string delivered = null; fixture.Service.ShowProcedureImmediate = () => { };
                fixture.Service.ExecuteProcedureCall = command => { delivered = command; return Task.FromResult<object>(null); };
                dynamic queued = fixture.Service.RunProcedure(request);
                request.ArgumentNames[0] = "injected"; request.Arguments[0] = 99; context.RunAll();
                Assert.AreEqual("Call VBAProject.Module1.TryMe(first:=7)", delivered);
                Assert.AreEqual("Delivered", (string)((dynamic)fixture.Service.ProcedureRunStatus(new Request { Project = fixture.Project.Name, Query = queued.Query })).State);
            }
            finally { SynchronizationContext.SetSynchronizationContext(previous); }
        }

        [TestMethod]
        public void ProcedurePreflightCoversAllIndependentIdentityAndRevisionBoundaries()
        {
            for (int scenario=0; scenario<13; scenario++)
            {
                var f=Create(2); var r=ProcedureRequest(f);
                if(scenario==0)r.ExpectedMode=1;
                if(scenario==1)r.Arguments=null;
                if(scenario==2)r.Arguments=new object[31];
                if(scenario==3)r.Procedure=null;
                if(scenario==4)f.Project.Name=r.Project="bad.name";
                if(scenario==5){f.Vbe.VBProjects.Add(new FakeProject{Name=f.Project.Name,Mode=2});r.Project=f.Project.FileName;}
                if(scenario==6)r.Module="Missing";
                if(scenario==7)f.Project.VBComponents[0].Type=3;
                if(scenario==8)f.Project.VBComponents[0].Name=r.Module="bad.module";
                if(scenario==9)r.ExpectedSha256=null;
                if(scenario==10){f.Module.Code="";r.ExpectedSha256=Sha("");}
                if(scenario==11)f.Module.ProcBodyLine.Body=0;
                if(scenario==12)f.Module.ProcBodyLine.Body=3;
                try {f.Service.PrepareProcedureCall(r);Assert.Fail("Invalid boundary accepted: "+scenario);}
                catch(ArgumentException) {Assert.IsTrue(scenario<=4||scenario==8);}
                catch(InvalidOperationException) {Assert.IsTrue(scenario>=5);}
            }
            var valid=Create(2);var request=ProcedureRequest(valid);
            valid.Vbe.VBProjects.Add(new FakeProject{Name="Other",Mode=2});
            Assert.AreEqual("Call VBAProject.Module1.TryMe()",valid.Service.PrepareProcedureCall(request));
            request.Arguments=new object[]{new string('x',2048)};
            Assert.ThrowsException<ArgumentException>(()=>valid.Service.PrepareProcedureCall(request));
            foreach(object scalar in new object[]{(byte)1,(sbyte)2,(short)3,(ushort)4,5,(uint)6,(long)7,(ulong)8,9m,10f,11d,false})
                Assert.AreEqual(scalar is bool?"False":Convert.ToString(scalar,CultureInfo.InvariantCulture),VbeDebug.ProcedureLiteral(scalar));
            Assert.ThrowsException<ArgumentException>(()=>VbeDebug.ProcedureLiteral(float.NegativeInfinity));
        }
        [TestMethod]
        public async Task ProcedureQueueTracksDeliveringFailuresPostRejectionAndBoundedRetention()
        {
            var previous=SynchronizationContext.Current;var context=new NativeNavigationContext();SynchronizationContext.SetSynchronizationContext(context);
            try
            {
                var f=Create(2);var r=ProcedureRequest(f);
                Assert.ThrowsException<ArgumentNullException>(()=>f.Service.RunProcedure(null));
                r.Arguments=null;Assert.ThrowsException<ArgumentException>(()=>f.Service.RunProcedure(r));r.Arguments=new object[0];
                SynchronizationContext.SetSynchronizationContext(null);Assert.ThrowsException<InvalidOperationException>(()=>f.Service.RunProcedure(r));
                SynchronizationContext.SetSynchronizationContext(context);
                f.Service.ShowProcedureImmediate=()=>{};
                var pending=new TaskCompletionSource<object>();f.Service.ExecuteProcedureCall=command=>pending.Task;
                dynamic queued=f.Service.RunProcedure(r);context.RunNext();
                dynamic delivering=f.Service.ProcedureRunStatus(new Request{Project=f.Project.Name,Query=queued.Query});
                Assert.AreEqual("Delivering",(string)delivering.State);Assert.IsTrue((bool)delivering.Pending);
                Assert.ThrowsException<InvalidOperationException>(()=>f.Service.RunProcedure(r));
                pending.SetResult("echo");context.RunAll();
                dynamic delivered=f.Service.ProcedureRunStatus(new Request{Project=f.Project.Name,Query=queued.Query});Assert.AreEqual("Delivered",(string)delivered.State);
                context.RejectPost=context.Posts+1;dynamic failed=f.Service.RunProcedure(r);Assert.AreEqual("Failed",(string)failed.State);StringAssert.Contains((string)failed.Error,"post");context.RejectPost=0;
                f.Service.ExecuteProcedureCall=command=>Task.FromException<object>(new InvalidOperationException("delivery rejected"));
                failed=f.Service.RunProcedure(r);context.RunAll();Assert.AreEqual("Failed",(string)((dynamic)f.Service.ProcedureRunStatus(new Request{Project=f.Project.Name,Query=failed.Query})).State);
                f.Service.ExecuteProcedureCall=command=>Task.FromResult<object>("ok");
                dynamic changed=f.Service.RunProcedure(r);f.Project.Name=f.Project.Name.ToLowerInvariant();context.RunAll();
                StringAssert.Contains((string)((dynamic)f.Service.ProcedureRunStatus(new Request{Project=r.Project,Query=changed.Query})).Error,"identity changed");
                f.Project.Name=r.Project;
                for(int i=0;i<21;i++){f.Service.RunProcedure(r);context.RunAll();}
                Assert.ThrowsException<InvalidOperationException>(()=>f.Service.ProcedureRunStatus(new Request{Project=r.Project,Query=queued.Query}));
                f.Service.ShowProcedureImmediate=null;dynamic nativePane=f.Service.RunProcedure(r);context.RunAll();
                Assert.AreEqual("Failed",(string)((dynamic)f.Service.ProcedureRunStatus(new Request{Project=r.Project,Query=nativePane.Query})).State);
                // Default transport validates a missing Immediate command before touching any host.
                SynchronizationContext.SetSynchronizationContext(null);
                await Assert.ThrowsExceptionAsync<ArgumentException>(()=>Create(2).Service.ExecuteProcedureCall(null));
            }
            finally{SynchronizationContext.SetSynchronizationContext(previous);}
        }
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
