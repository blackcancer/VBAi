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

        [TestMethod]
        public void ReadAndEditModuleRequireTheVersionAndDesignMode()
        {
            var project = new FakeProject { Name = "Projet", Mode = 2 };
            var module = new FakeModule("Option Explicit\r\nSub Essai()\r\nEnd Sub");
            project.VBComponents.Add(new FakeComponent { Name = "Module1", Type = 1, CodeModule = module });
            var host = new FakeVbe();
            host.VBProjects.Add(project);
            var session = new VbeSession(host);
            dynamic read = session.Execute(new Request { Command = "read_module", Project = "Projet", Module = "Module1" }).Data;
            Assert.AreEqual(module.Code, (string)read.Code);
            Assert.IsFalse(session.Execute(new Request { Command = "replace_lines", Project = "Projet", Module = "Module1" }).Ok);
            Assert.IsFalse(session.Execute(new Request { Command = "replace_lines", Project = "Projet", Module = "Module1",
                ExpectedSha256 = "outdated", StartLine = 1, Count = 1, Text = "Option Private Module" }).Ok);
            project.Mode = 1;
            Assert.IsFalse(session.Execute(new Request { Command = "replace_lines", Project = "Projet", Module = "Module1",
                ExpectedSha256 = (string)read.Sha256, StartLine = 1, Count = 1, Text = "Option Private Module" }).Ok);
            project.Mode = 2;
            var edit = session.Execute(new Request { Command = "replace_lines", Project = "Projet", Module = "Module1",
                ExpectedSha256 = (string)read.Sha256, StartLine = 1, Count = 1, Text = "Option Private Module" });
            Assert.IsTrue(edit.Ok);
            StringAssert.StartsWith(module.Code, "Option Private Module");
            Assert.IsFalse(session.Execute(new Request { Command = "replace_lines", Project = "Projet", Module = "Module1",
                ExpectedSha256 = (string)read.Sha256, StartLine = 1, Count = 1, Text = "Option Explicit" }).Ok);
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
        public sealed class FakeModule
        {
            private readonly List<string> lines;
            public FakeModule() : this(0) { }
            public FakeModule(int count) { lines = Enumerable.Repeat(string.Empty, count).ToList(); Lines = new FakeLines(this); }
            public FakeModule(string code) { lines = code.Split(new[] { "\r\n" }, StringSplitOptions.None).ToList(); Lines = new FakeLines(this); }
            public int CountOfLines { get { return lines.Count; } set { lines.Clear(); lines.AddRange(Enumerable.Repeat(string.Empty, value)); } }
            public string Code { get { return string.Join("\r\n", lines); } }
            public FakeLines Lines { get; }
            public bool CorruptNonAsciiOnInsert { get; set; }
            public void DeleteLines(int start, int count) { lines.RemoveRange(start - 1, count); }
            public void InsertLines(int start, string text)
            {
                if (CorruptNonAsciiOnInsert) text = text.Replace('é', '?');
                lines.InsertRange(start - 1, text.Split(new[] { "\r\n" }, StringSplitOptions.None));
            }
            public sealed class FakeLines
            {
                private readonly FakeModule module;
                public FakeLines(FakeModule module) { this.module = module; }
                public string this[int start, int count] { get { return string.Join("\r\n", module.lines.Skip(start - 1).Take(count)); } }
            }
        }
    }
}
