using System;
using System.Collections.Generic;
using System.Linq;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeSessionContractTests
    {
        [TestMethod]
        public void DispatcherRejectsMissingAndUnknownCommandsWithoutTouchingTheHost()
        {
            var session = new VbeSession(new FakeVbe());
            Assert.AreEqual("A command is required.", session.Execute(null).Error);
            Assert.AreEqual("A command is required.", session.Execute(new Request { Command = " " }).Error);
            Assert.AreEqual("Unknown command: no_such_command", session.Execute(new Request { Command = "no_such_command" }).Error);
            Assert.IsTrue(session.Execute(new Request { Command = "status" }).Ok);
        }

        [TestMethod]
        public void ProjectAndModuleDiscoveryRetainNamesTypesAndLineCounts()
        {
            var project = new FakeProject { Name = "ClasseurÉté", FileName = @"C:\Temp\été.xlsm", Mode = 2 };
            project.VBComponents.Add(new FakeComponent { Name = "Calcul", Type = 1, CodeModule = new FakeModule { CountOfLines = 7 } });
            project.VBComponents.Add(new FakeComponent { Name = "Métier", Type = 2, CodeModule = new FakeModule { CountOfLines = 3 } });
            var host = new FakeVbe();
            host.VBProjects.Add(project);
            var session = new VbeSession(host);
            var projects = (IEnumerable<object>)session.Execute(new Request { Command = "list_projects" }).Data;
            dynamic discoveredProject = projects.Single();
            Assert.AreEqual("ClasseurÉté", (string)discoveredProject.Name);
            Assert.AreEqual(project.FileName, (string)discoveredProject.FileName);
            Assert.AreEqual(2, (int)discoveredProject.Mode);
            var response = session.Execute(new Request { Command = "list_modules", Project = "classeurété" });
            Assert.IsTrue(response.Ok);
            var modules = ((IEnumerable<object>)response.Data).Cast<dynamic>().ToArray();
            Assert.AreEqual(2, modules.Length);
            Assert.AreEqual("Calcul", (string)modules[0].Name);
            Assert.AreEqual(1, (int)modules[0].Type);
            Assert.AreEqual(7, (int)modules[0].Lines);
            Assert.AreEqual("Métier", (string)modules[1].Name);
            Assert.AreEqual(2, (int)modules[1].Type);
            Assert.AreEqual(3, (int)modules[1].Lines);
        }

        public sealed class FakeVbe { public List<FakeProject> VBProjects { get; } = new List<FakeProject>(); }
        public sealed class FakeProject
        {
            public string Name { get; set; }
            public string FileName { get; set; }
            public int Mode { get; set; }
            public List<FakeComponent> VBComponents { get; } = new List<FakeComponent>();
        }
        public sealed class FakeComponent
        {
            public string Name { get; set; }
            public int Type { get; set; }
            public FakeModule CodeModule { get; set; }
        }
        public sealed class FakeModule { public int CountOfLines { get; set; } }
    }
}
