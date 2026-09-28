namespace CodexVBE.Tests.Unit
{
    using System;
    using System.IO;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class VbeProjectExcelHostTests
    {
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
