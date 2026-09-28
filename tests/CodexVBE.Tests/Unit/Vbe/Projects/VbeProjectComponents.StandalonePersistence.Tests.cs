namespace CodexVBE.Tests.Unit
{
    using System;
    using System.IO;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class VbeProjectComponentsTests
    {
        [TestMethod]
        public void StandaloneDetectionAndPersistenceDistinguishUnsavedRelativeMissingAndReadOnlyFiles()
        {
            string root=Path.Combine(Path.GetTempPath(),"StandaloneMatrix-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(root);string path=Path.Combine(root,"macro.swp");
            try
            {
                var project=new FaultedStandaloneProject();var service=Service(project);
                foreach(string value in new[]{null,"","relative.swp",path})
                {project.FileName=value;dynamic state=service.PersistenceStatus(project.Name);Assert.IsFalse((bool)state.FileExists);Assert.AreEqual(value==path,(bool)state.HostHasPath);Assert.IsNull(state.HostSaved);Assert.IsNull(state.HostReadOnly);}
                File.WriteAllText(path,"x");project.FileName=path;File.SetAttributes(path,FileAttributes.ReadOnly);dynamic present=service.PersistenceStatus(project.Name);Assert.IsTrue((bool)present.HostReadOnly);File.SetAttributes(path,FileAttributes.Normal);
                project.FailFileName=true;Assert.AreEqual(false,typeof(VbeProjectComponents).GetMethod("SupportsStandaloneMacro",System.Reflection.BindingFlags.Static|System.Reflection.BindingFlags.NonPublic).Invoke(null,new object[]{project}));
                project.FailFileName=false;project.FileName="macro.xlsm";dynamic metadata=service.ProjectProperties(project.Name);Assert.ThrowsException<InvalidOperationException>(()=>service.SaveHostDocument(new Request{Project=project.Name,ExpectedHostPath=path,ExpectedProjectVersion=metadata.Version}));
            }
            finally{if(File.Exists(path))File.SetAttributes(path,FileAttributes.Normal);Directory.Delete(root,true);}
        }
        [TestMethod]
        public void StandaloneSaveValidationAndNativeReadbackCoverEveryIndependentBoundary()
        {
            string root=Path.Combine(Path.GetTempPath(),"StandaloneMatrix-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(root);
            try
            {
                for(int scenario=0;scenario<16;scenario++)
                {
                    string path=Path.Combine(root,"macro"+scenario+".swp");File.WriteAllText(path,"before");
                    var p=new FaultedStandaloneProject{FileName=path,Saved=false};var service=Service(p);dynamic metadata=service.ProjectProperties(p.Name);
                    var r=new Request{Project=p.Name,ExpectedProjectVersion=metadata.Version,ExpectedHostPath=path,Path=Path.Combine(root,"new"+scenario+".swp")};
                    bool saveAs=scenario>=5;
                    if(scenario==0)p.FileName="";if(scenario==1)p.FileName="relative.swp";if(scenario==2)r.ExpectedHostPath=null;if(scenario==3)r.ExpectedHostPath="relative.swp";if(scenario==4)File.Delete(path);
                    if(scenario==5)r.Path=null;if(scenario==6)r.Path="relative.swp";if(scenario==7)r.Path=Path.Combine(root,"missing","new.swp");
                    if(scenario>=8)p.Saving=destination=>{
                        p.FileName=destination;p.Saved=true;File.WriteAllText(destination,"native");
                        if(scenario==8)p.FileName=null;if(scenario==9)p.FileName="relative.swp";if(scenario==10)p.FileName=path;
                        if(scenario==11)p.Saved=false;if(scenario==12)File.Delete(destination);if(scenario==13)File.WriteAllText(destination,"");
                        if(scenario==14)throw new IOException("native save rejected");
                    };
                    // Refresh expected version after changing native preflight identity.
                    metadata=service.ProjectProperties(p.Name);r.ExpectedProjectVersion=metadata.Version;
                    if(scenario==15){dynamic result=service.SaveHostDocumentAs(r);Assert.IsTrue((bool)result.SaveInvoked);Assert.IsTrue(File.Exists(r.Path));}
                    else
                    {
                        Exception failure=null;try{if(saveAs)service.SaveHostDocumentAs(r);else service.SaveHostDocument(r);}catch(Exception error){failure=error;}
                        Assert.IsNotNull(failure,"Boundary "+scenario);Assert.IsTrue(failure is InvalidOperationException||failure is ArgumentException||failure is IOException);
                        Assert.AreEqual(scenario>=8?1:0,p.Saves,"Native save must follow preflight only.");
                    }
                }
            }
            finally{Directory.Delete(root,true);}
        }
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
