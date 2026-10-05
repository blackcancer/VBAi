using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VBAi;
namespace VBAi.Tests.Unit
{
    [TestClass]
    public sealed class VbeSolidWorksMacroGeneralSelectionTests
    {
        public sealed class Project : VbeProjectComponentsTests.FakeProject
        { public int Type {get;set;}=101;public int Protection {get;set;} }
        public sealed class Host
        {
            public List<Project> VBProjects {get;}=new List<Project>();
            Project active;
            public Action OnSelect;
            public bool FailGetter, IgnoreSetter;
            public Project ActiveVBProject {get{if(FailGetter)throw new InvalidOperationException("Unavailable active project");return active;}set{if(!IgnoreSetter)active=value;OnSelect?.Invoke();}}
        }
        static VbeProjectComponents Service(Host host){var s=new VbeProjectComponents(host,new VbeForms(host));s.GeneralProjectIdentity=ReferenceEquals;return s;}
        static Request ReadRequest(VbeProjectComponents service,Project source)=>new Request{Command="read_project_general",Project=source.Name,ExpectedMode=2,ExpectedProjectVersion=(string)((dynamic)service.ProjectProperties(source.Name)).Version};
        [TestMethod]
        public void GeneralSourceSelectionRestoresDistinctOriginalNativeDestination()
        {
            var source=new Project{Name="Source"};var target=new Project{Name="Destination",Type=100};var host=new Host{ActiveVBProject=target};host.VBProjects.Add(source);host.VBProjects.Add(target);var service=Service(host);
            var captured=service.CapturePublicationGeneralSelection();service.SelectPublicationGeneralProject(ReadRequest(service,source));Assert.AreSame(source,host.ActiveVBProject);
            service.RestorePublicationGeneralSelection(captured);Assert.AreSame(target,host.ActiveVBProject);
        }
        [DataTestMethod][DataRow("stale")][DataRow("mode")][DataRow("protected")][DataRow("type")][DataRow("canonical")][DataRow("name")][DataRow("ignore")][DataRow("getter")]
        public void SourceSelectionRefusesRevisionIdentityAndGetterChanges(string fault)
        {
            var source=new Project{Name="Source"};var target=new Project{Name="Destination",Type=100};var host=new Host{ActiveVBProject=target};host.VBProjects.Add(source);host.VBProjects.Add(target);var service=Service(host);var request=ReadRequest(service,source);
            if(fault=="stale")request.ExpectedProjectVersion="stale";if(fault=="mode")source.Mode=1;if(fault=="protected")source.Protection=1;if(fault=="type")source.Type=100;
            if(fault=="canonical")host.OnSelect=()=>host.VBProjects[0]=new Project{Name="Source"};
            if(fault=="name")host.OnSelect=()=>source.Name="Renamed";
            if(fault=="ignore")host.IgnoreSetter=true;if(fault=="getter")host.FailGetter=true;
            Assert.ThrowsException<InvalidOperationException>(()=>service.SelectPublicationGeneralProject(request));
        }
        [DataTestMethod][DataRow("revision")][DataRow("canonical")][DataRow("mode")][DataRow("protection")][DataRow("ignore")][DataRow("changed-during-set")]
        public void RestoreSelectionRefusesChangedOriginalInsteadOfAssumingSaveTargetActive(string fault)
        {
            var source=new Project{Name="Source"};var target=new Project{Name="Destination",Type=100};var host=new Host{ActiveVBProject=target};host.VBProjects.Add(source);host.VBProjects.Add(target);var service=Service(host);var captured=service.CapturePublicationGeneralSelection();host.ActiveVBProject=source;
            if(fault=="revision")target.Description="Changed";if(fault=="canonical")host.VBProjects[1]=new Project{Name="Destination",Type=100};if(fault=="mode")target.Mode=1;if(fault=="protection")target.Protection=1;
            if(fault=="ignore")host.IgnoreSetter=true;if(fault=="changed-during-set")host.OnSelect=()=>target.Description="Changed after select";
            Assert.ThrowsException<InvalidOperationException>(()=>service.RestorePublicationGeneralSelection(captured));
        }
        [TestMethod]
        public void CaptureSelectionRequiresActualReadableCanonicalActiveProject()
        {
            var host=new Host();var service=Service(host);Assert.ThrowsException<InvalidOperationException>(()=>service.CapturePublicationGeneralSelection());
            host.FailGetter=true;Assert.ThrowsException<InvalidOperationException>(()=>service.CapturePublicationGeneralSelection());
        }
    }
}
