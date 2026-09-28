namespace CodexVBE.Tests.Unit
{
    using System;
    using System.IO;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class VbeProjectExcelHostTests
    {
        [TestMethod]
        public void RenamePreflightRefusesHostScopeIdentifiersProtectionAndPathChangesBeforeSetter()
        {
            for(int scenario=0;scenario<9;scenario++)
            {
                var f=new RenameWorkflowFixture();var r=f.Request();int setters=0;f.Project.OnRename=()=>setters++;
                if(scenario==0)f.Host.IsExcel=false;if(scenario==1)r.Project=null;if(scenario==2)r.Project="relative.xlsm";
                if(scenario==3)r.Value=1;if(scenario==4)r.Value="bad.name";if(scenario==5)r.Value="If";
                if(scenario==6){f.Project.FailProtection=true;r=f.Request();}
                if(scenario==7)f.Workbook.ReadingReadOnly=()=>f.Project.FileName=Path.Combine(Path.GetTempPath(),"changed.xlsm");
                if(scenario==8)r.ExpectedProjectVersion="stale";
                Exception failure=null;try{f.Service.SetProjectProperty(r);}catch(Exception error){failure=error;}
                Assert.IsNotNull(failure,"Boundary "+scenario);Assert.IsTrue(failure is ArgumentException||failure is InvalidOperationException);Assert.AreEqual(0,setters);
            }
        }
        [TestMethod]
        public void RenameReadbackReportsIgnoredSetterAndIndependentPathOrSourceMutations()
        {
            for(int scenario=0;scenario<6;scenario++)
            {
                var f=new RenameWorkflowFixture();var r=f.Request(scenario==0?f.Project.Name:"Renamed");int setters=0;
                f.Project.OnRename=()=>{setters++;
                    if(scenario==2)f.Project.VBComponents.Add(new VbeProjectComponentsTests.FakeComponent("Added",1));
                    if(scenario==3){foreach(var c in f.Project.VBComponents)c.CodeModule.Source="changed";}
                    if(scenario==4){foreach(var c in f.Project.VBComponents)c.Name="Other";}
                    if(scenario==5)f.Project.FileName=Path.Combine(Path.GetDirectoryName(r.Project),".",Path.GetFileName(r.Project));
                };
                if(scenario==1)f.Project.IgnoreName=true;
                dynamic result=f.Service.SetProjectProperty(r);
                Assert.AreEqual(scenario==0,(bool)result.Verified);Assert.AreEqual(scenario==0?0:1,setters);
                Assert.AreEqual(scenario!=2&&scenario!=3&&scenario!=4,(bool)result.SourcePreserved);
                Assert.AreEqual(scenario!=5,(bool)result.HostPathPreserved);
            }
        }
        [TestMethod]
        public void RenameSourceBudgetRefusesNegativeExcessiveLinesAndAggregateCharacters()
        {
            foreach(int scenario in new[]{0,1,2,3})
            {
                var f=new RenameWorkflowFixture();foreach(var c in f.Project.VBComponents){c.CodeModule.LineCount=scenario==0?-1:scenario==1?100001:scenario==3?0:1;c.CodeModule.Source=scenario==2?new string('x',4*1024*1024+1):"";}
                var r=f.Request();
                if(scenario==3)Assert.IsTrue((bool)((dynamic)f.Service.SetProjectProperty(r)).Verified);
                else Assert.ThrowsException<InvalidOperationException>(()=>f.Service.SetProjectProperty(r));
            }
        }
        [TestMethod]
        public void SavedExcelRenameVerifiesStablePathSourcesAndRefusesProtectionAndReadOnly()
        {
            var project=new RenameProject {FileName=Path.Combine(Path.GetTempPath(),"qualified.xlsm")};
            project.VBComponents.Add(new VbeProjectComponentsTests.FakeComponent("Module1",1));
            var vbe=new VbeProjectComponentsTests.FakeVbe();vbe.VBProjects.Add(project);
            var workbook=new FakeWorkbook(project) {FullName=project.FileName,Path=Path.GetDirectoryName(project.FileName)};
            var excel=new FakeExcel();excel.Workbooks.Add(workbook);var host=new FakeHost {Excel=excel};
            var service=new VbeProjectComponents(vbe,new VbeForms(vbe),host);
            dynamic metadata=service.ProjectProperties(project.FileName);
            var request=new Request {Project=project.FileName,Property="Name",Value="Qualified",ExpectedProjectVersion=(string)metadata.Version};
            dynamic renamed=service.SetProjectProperty(request);
            Assert.AreEqual("Qualified",project.Name);Assert.IsTrue((bool)renamed.Verified);Assert.IsTrue((bool)renamed.SourcePreserved);Assert.IsFalse((bool)renamed.PersistenceVerified);
            metadata=service.ProjectProperties(project.FileName);request.ExpectedProjectVersion=(string)metadata.Version;request.Value="Other";
            project.Protection=1;metadata=service.ProjectProperties(project.FileName);request.ExpectedProjectVersion=(string)metadata.Version;
            Assert.ThrowsException<InvalidOperationException>(()=>service.SetProjectProperty(request));Assert.AreEqual("Qualified",project.Name);
            project.Protection=0;workbook.ReadOnly=true;metadata=service.ProjectProperties(project.FileName);request.ExpectedProjectVersion=(string)metadata.Version;
            Assert.ThrowsException<InvalidOperationException>(()=>service.SetProjectProperty(request));
            workbook.ReadOnly=false;request.Value="Module1";
            Assert.ThrowsException<InvalidOperationException>(()=>service.SetProjectProperty(request));
        }
        public sealed class RenameProject : VbeProjectComponentsTests.FakeProject {public int Protection {get;set;}}
    }
}
