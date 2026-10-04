using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi;

namespace VBAi.Tests.Unit
{
    [TestClass]
    public sealed class VbeProjectGeneralProjectTests
    {
        [TestMethod]
        public void ExactNonnegativeNumericIntegersAndInvariantIntegerStringsAreAccepted()
        {
            foreach(var value in new object[]{0,321,int.MaxValue,(byte)3,(short)321,(uint)321,(long)321,(ulong)321,321m,321d,321f,"321","2147483647"})
                Assert.AreEqual(Convert.ToInt32(value,System.Globalization.CultureInfo.InvariantCulture),VbeProjectComponents.RequireGeneralInt32(value));
        }
        [TestMethod]
        public void BoolNegativeFractionNonfiniteOverflowAndNonnumericInputsRefuseBeforeAnyHostRead()
        {
            foreach(var value in new object[]{null,true,false,-1,"-1",321.5,321.5f,321.000000000001m,double.NaN,double.PositiveInfinity,float.NegativeInfinity,(long)int.MaxValue+1,ulong.MaxValue,"321.0","0x141","",new object(),'3',DayOfWeek.Monday})
                Assert.ThrowsException<ArgumentException>(()=>VbeProjectComponents.RequireGeneralInt32(value),Convert.ToString(value));
        }
        [TestMethod]
        public void HelpFileRequiresExactExistingRootedChmWithoutUnicodeConversionOrNormalization()
        {
            string dir=Path.Combine(Path.GetTempPath(),"vbai-general-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(dir);
            try
            {
                string path=Path.Combine(dir,"qualification-é漢字.chm");File.WriteAllText(path,"synthetic nonexecuted metadata target");
                Assert.AreEqual(path,VbeProjectComponents.RequireGeneralHelpFile(path));
                foreach(var value in new object[]{true,321,"relative.chm",Path.Combine(dir,"absent.chm"),Path.Combine(dir,"other.hlp"),Path.Combine(dir,".","qualification-é漢字.chm"),path+"\0"})
                    Assert.ThrowsException<ArgumentException>(()=>VbeProjectComponents.RequireGeneralHelpFile(value));
                Assert.AreEqual("synthetic nonexecuted metadata target",File.ReadAllText(path));
            }
            finally{Directory.Delete(dir,true);}
        }
        public sealed class Project : VbeProjectComponentsTests.FakeProject
        { public int Protection{get;set;} public int HelpContextID{get;set;} }
        public sealed class Window { public long HWnd{get;set;}=42; }
        public sealed class Host
        {
            public List<VbeProjectComponentsTests.FakeProject> VBProjects{get;}=new List<VbeProjectComponentsTests.FakeProject>();
            public Project ActiveVBProject{get;set;} public Window MainWindow{get;}=new Window();
        }
        [DataTestMethod][DataRow("selector")][DataRow("active")][DataRow("revision")][DataRow("mode")][DataRow("protection")][DataRow("root")][DataRow("policy")]
        public async Task FinalProjectAndRuntimeChangesRefuseUiBeforeOriginalCommandEntry(string fault)
        {
            var project=new Project{Name="Disposable"};var host=new Host{ActiveVBProject=project};host.VBProjects.Add(project);
            var service=new VbeProjectComponents(host,new VbeForms(host));var native=new VbeProjectGeneralOperationTests.Native();var scheduler=new VbeProjectGeneralOperationTests.Scheduler();
            service.GeneralProjectIdentity=ReferenceEquals;service.GeneralNativeFactory=(root,context)=>native;service.GeneralOperationFactory=n=>new VbeProjectGeneralOperation(n,scheduler);
            dynamic state=service.ProjectProperties(project.Name);int opens=0;bool revoked=false;
            var request=new Request{Project=project.Name,ExpectedMode=2,ExpectedProjectVersion=state.Version,ControlCaption="Properties",Property="HelpContextID",Value=321,ExpectedOptionsVersion=native.State.OptionsVersion};
            request.RevalidateProjectPropertyAuthorization=scope=>{if(revoked)throw new InvalidOperationException("approval revoked");};
            Task<object> task; var previous=SynchronizationContext.Current;
            try
            {
                // Pure managed seams need no ambient WinForms dispatcher leaked by an earlier STA test.
                SynchronizationContext.SetSynchronizationContext(null);
                task=service.ProjectGeneralAsync(request,true,(captured,live)=>before=>{live();before();opens++;},_=>{},()=>{});
            }
            finally { SynchronizationContext.SetSynchronizationContext(previous); }
            if(fault=="selector")host.VBProjects[0]=new Project{Name="Disposable"};
            if(fault=="active")host.ActiveVBProject=new Project{Name="Disposable"};
            if(fault=="revision")project.Description="changed";if(fault=="mode")project.Mode=1;if(fault=="protection")project.Protection=1;
            if(fault=="root")host.MainWindow.HWnd=43;if(fault=="policy")revoked=true;
            scheduler.Posted();var result=(VbeProjectGeneralOperation.Result)await task.ConfigureAwait(false);
            Assert.IsFalse(result.CommandEntered||result.MutationInvoked);Assert.AreEqual(0,opens);Assert.AreEqual(0,native.Writes);Assert.IsNotNull(result.Error);
        }
        [DataTestMethod][DataRow(false)][DataRow(true)]
        public async Task ChangedProjectAfterSettledModalCannotPublishMetadataOrPromoteWrittenOutcome(bool write)
        {
            var project=new Project{Name="Disposable"};var host=new Host{ActiveVBProject=project};host.VBProjects.Add(project);
            var service=new VbeProjectComponents(host,new VbeForms(host));var native=new VbeProjectGeneralOperationTests.Native();var scheduler=new VbeProjectGeneralOperationTests.Scheduler();
            service.GeneralProjectIdentity=ReferenceEquals;service.GeneralNativeFactory=(root,context)=>native;service.GeneralOperationFactory=n=>new VbeProjectGeneralOperation(n,scheduler);
            dynamic before=service.ProjectProperties(project.Name);int terminalReceipts=0;
            var request=new Request{Project=project.Name,ExpectedMode=2,ExpectedProjectVersion=before.Version,ControlCaption="Properties",Property="HelpContextID",Value=321,ExpectedOptionsVersion=native.State.OptionsVersion,RevalidateProjectPropertyAuthorization=_=>{}};
            native.OnClose=()=>host.ActiveVBProject=new Project{Name="Disposable"};
            Task<object> task; var previous=SynchronizationContext.Current;
            try
            {
                SynchronizationContext.SetSynchronizationContext(null);
                task=service.ProjectGeneralAsync(request,write,(captured,live)=>entry=>{live();entry();native.Visible=true;scheduler.Pump();scheduler.Pump();},receipt=>{if(receipt.Terminal)terminalReceipts++;},()=>{});
            }
            finally { SynchronizationContext.SetSynchronizationContext(previous); }
            scheduler.Posted();scheduler.Pump();var result=(VbeProjectGeneralOperation.Result)await task.ConfigureAwait(false);
            Assert.IsFalse(result.Available);Assert.IsNotNull(result.Error);Assert.AreEqual(write,result.Uncertain);
            Assert.IsNull(result.Name);Assert.IsNull(result.Description);Assert.IsNull(result.HelpFile);Assert.IsNull(result.HelpContextText);Assert.IsNull(result.ConditionalCompilation);Assert.IsNull(result.OptionsVersion);
            Assert.AreEqual(write?1:0,native.Writes);Assert.AreEqual(1,native.Closes);Assert.AreEqual(2,terminalReceipts);
        }
        [DataTestMethod][DataRow("none")][DataRow("revision")][DataRow("policy")][DataRow("publicationReceipt")]
        public async Task KnownEncodingRefusalRevalidatesUnchangedProjectAndRedactsMetadataAfterCancel(string fault)
        {
            string directory = Path.Combine(Path.GetTempPath(), "vbai-general-refusal-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try
            {
                string path = Path.Combine(directory, "日本.chm"); File.WriteAllText(path, "inert marker");
                var project = new Project { Name = "Disposable" }; var host = new Host { ActiveVBProject = project }; host.VBProjects.Add(project);
                var service = new VbeProjectComponents(host, new VbeForms(host));
                var native = new VbeProjectGeneralOperationTests.Native(); var scheduler = new VbeProjectGeneralOperationTests.Scheduler();
                service.GeneralProjectIdentity = ReferenceEquals; service.GeneralNativeFactory = (root, context) => native;
                service.GeneralOperationFactory = n => new VbeProjectGeneralOperation(n, scheduler);
                dynamic initial = service.ProjectProperties(project.Name); bool revoked = false; int terminalReceipts = 0;
                var request = new Request { Project = project.Name, ExpectedMode = 2, ExpectedProjectVersion = initial.Version,
                    ControlCaption = "Properties", Property = "HelpFile", Value = path, ExpectedOptionsVersion = native.State.OptionsVersion,
                    RevalidateProjectPropertyAuthorization = _ => { if (revoked) throw new InvalidOperationException("privacy revoked after Cancel"); } };
                native.OnClose = () => {
                    if (fault == "revision") project.Description = "changed after Cancel";
                    if (fault == "policy") revoked = true;
                };
                Task<object> task; var previous = SynchronizationContext.Current;
                try
                {
                    SynchronizationContext.SetSynchronizationContext(null);
                    task = service.ProjectGeneralAsync(request, true, (captured, live) => entry => {
                        live(); entry(); native.Visible = true; scheduler.Pump(); scheduler.Pump();
                    }, receipt => {
                        if (receipt.Terminal && ++terminalReceipts == 2 && fault == "publicationReceipt")
                            throw new InvalidOperationException("final receipt failed");
                    }, () => { });
                }
                finally { SynchronizationContext.SetSynchronizationContext(previous); }
                scheduler.Posted(); scheduler.Pump(); var result = (VbeProjectGeneralOperation.Result)await task.ConfigureAwait(false);
                Assert.IsTrue(result.RefusedBeforeWrite && result.DialogClosed && result.OriginalExecuteReturned && result.Terminal);
                Assert.AreEqual(fault != "none", result.Uncertain); Assert.IsFalse(result.Available || result.MutationInvoked || result.CommittedRequested);
                Assert.AreEqual(0, result.FieldAttempts); Assert.AreEqual(0, result.OkAttempts); Assert.AreEqual(1, result.CancelAttempts);
                Assert.AreEqual(0, native.Writes); Assert.AreEqual(1, native.Closes); Assert.AreEqual(2, terminalReceipts);
                Assert.IsNotNull(result.Error); Assert.IsNull(result.Name); Assert.IsNull(result.Description); Assert.IsNull(result.HelpFile);
                Assert.IsNull(result.HelpContextText); Assert.IsNull(result.ConditionalCompilation); Assert.IsNull(result.OptionsVersion);
                Assert.IsFalse(result.Error.Contains(path), "The refused file path must not appear in the public error.");
            }
            finally { Directory.Delete(directory, true); }
        }
    }
}
