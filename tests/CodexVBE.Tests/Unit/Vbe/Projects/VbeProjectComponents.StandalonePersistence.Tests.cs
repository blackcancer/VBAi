namespace CodexVBE.Tests.Unit
{
    using System;
    using System.IO;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class VbeProjectComponentsTests
    {
        [TestMethod]
        public void StandaloneMacroSaveUsesNativeApiAndChecksExpectedVersionAndPath()
        {
            string directory = Path.Combine(Path.GetTempPath(), "VBAi-Standalone-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(directory);
            string path = Path.Combine(directory,"fixture.swp"); File.WriteAllText(path,"before");
            try
            {
                var project = new StandaloneProject { FileName=path,Saved=false };
                var service = Service(project);
                dynamic state = service.PersistenceStatus(project.Name); Assert.IsTrue((bool)state.HostAvailable); Assert.AreEqual(path,(string)state.HostPath);
                dynamic metadata = service.ProjectProperties(project.Name);
                var request = new Request { Project=project.Name,ExpectedHostPath=path,ExpectedProjectVersion=(string)metadata.Version };
                dynamic saved = service.SaveHostDocument(request);
                Assert.IsTrue((bool)saved.SaveInvoked); Assert.IsFalse((bool)saved.ReloadVerified); Assert.AreEqual(1,project.SaveCalls);
                Assert.AreEqual("native fixture",File.ReadAllText(path));
                request.ExpectedProjectVersion="stale";
                Assert.ThrowsException<InvalidOperationException>(()=>service.SaveHostDocument(request)); Assert.AreEqual(1,project.SaveCalls);
                metadata=service.ProjectProperties(project.Name); request.ExpectedProjectVersion=(string)metadata.Version;
                request.ExpectedHostPath=Path.Combine(directory,"different.swp");
                Assert.ThrowsException<InvalidOperationException>(()=>service.SaveHostDocument(request));
                request.ExpectedHostPath=path; File.SetAttributes(path,FileAttributes.ReadOnly);
                Assert.ThrowsException<InvalidOperationException>(()=>service.SaveHostDocument(request)); Assert.AreEqual(1,project.SaveCalls);
            }
            finally { File.SetAttributes(path,FileAttributes.Normal); Directory.Delete(directory,true); }
        }
        [TestMethod]
        public void StandaloneMacroSaveAsRefusesOverwriteHostProjectsAndWrongFormat()
        {
            string directory=Path.Combine(Path.GetTempPath(),"VBAi-Standalone-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(directory);
            string path=Path.Combine(directory,"new.swp");
            try
            {
                var project=new StandaloneProject {FileName="",Saved=false};var service=Service(project);
                dynamic metadata=service.ProjectProperties(project.Name);
                var request=new Request {Project=project.Name,Path=path,ExpectedProjectVersion=(string)metadata.Version};
                dynamic result=service.SaveHostDocumentAs(request);Assert.IsTrue((bool)result.SaveAsInvoked);Assert.IsTrue(File.Exists(path));
                metadata=service.ProjectProperties(project.Name);request.ExpectedProjectVersion=(string)metadata.Version;
                Assert.ThrowsException<IOException>(()=>service.SaveHostDocumentAs(request));
                request.Path=Path.Combine(directory,"bad.xlsm");Assert.ThrowsException<ArgumentException>(()=>service.SaveHostDocumentAs(request));
                project.Type=100;request.Path=Path.Combine(directory,"host.swp");
                Assert.ThrowsException<InvalidOperationException>(()=>service.SaveHostDocumentAs(request));Assert.AreEqual(1,project.SaveCalls);
            }
            finally { Directory.Delete(directory,true); }
        }
        public sealed class StandaloneProject : FakeProject
        {
            public int Type {get;set;}=101;
            public int SaveCalls {get;private set;}
            public void SaveAs(string destination) {SaveCalls++;File.WriteAllText(destination,"native fixture");FileName=destination;Saved=true;}
        }
    }
}
