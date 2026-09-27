using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeSessionTests
    {
        private static Fixture Create(string code = "Alpha\r\nBeta", int mode = 2)
        {
            var project = new FakeProject { Name = "VBAProject", FileName = @"C:\Temp\Host.xlsm", Mode = mode };
            project.VBComponents.Items.Add(new FakeComponent { Name = "Module1", Type = 1,
                CodeModule = new FakeModule(code) });
            var vbe = new FakeVbe();
            vbe.VBProjects.Add(project);
            return new Fixture { Vbe = vbe, Project = project, Session = new VbeSession(vbe) };
        }

        private static string Sha(string code)
        {
            using (var sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(code)))
                    .Replace("-", "").ToLowerInvariant();
        }

        [TestMethod]
        public void DispatchReportsMissingUnknownAndStatusCommands()
        {
            var f = Create();
            foreach (var request in new[] { null, new Request(), new Request { Command = " " } })
            {
                var response = f.Session.Execute(request);
                Assert.IsFalse(response.Ok);
                StringAssert.Contains(response.Error, "command is required");
            }
            var unknown = f.Session.Execute(new Request { Command = "does_not_exist" });
            Assert.IsFalse(unknown.Ok);
            Assert.AreEqual("Unknown command: does_not_exist", unknown.Error);
            var status = f.Session.Execute(new Request { Command = "status" });
            Assert.IsTrue(status.Ok);
            Assert.IsTrue((bool)((dynamic)status.Data).Connected);
        }

        [TestMethod]
        public void ListProjectsRetainsUnsavedProjectWhenFileNameIsUnavailable()
        {
            var f = Create();
            f.Vbe.VBProjects.Add(new FakeProject { Name = "Unsaved", ThrowFileName = true, Mode = 1 });
            var response = f.Session.Execute(new Request { Command = "list_projects" });
            Assert.IsTrue(response.Ok);
            var projects = ((IEnumerable)response.Data).Cast<object>().ToArray();
            Assert.AreEqual(2, projects.Length);
            Assert.AreEqual(@"C:\Temp\Host.xlsm", (string)((dynamic)projects[0]).FileName);
            Assert.AreEqual("Unsaved", (string)((dynamic)projects[1]).Name);
            Assert.IsNull((string)((dynamic)projects[1]).FileName);
            Assert.AreEqual(1, (int)((dynamic)projects[1]).Mode);
        }

        [TestMethod]
        public void ListModulesAndReadModuleDispatchToSelectedProject()
        {
            var f = Create();
            f.Project.VBComponents.Items.Add(new FakeComponent { Name = "Class1", Type = 2,
                CodeModule = new FakeModule("") });
            var listed = f.Session.Execute(new Request { Command = "list_modules", Project = "vbaproject" });
            Assert.IsTrue(listed.Ok);
            var modules = ((IEnumerable)listed.Data).Cast<object>().ToArray();
            Assert.AreEqual(2, modules.Length);
            Assert.AreEqual(2, (int)((dynamic)modules[0]).Lines);
            Assert.AreEqual(0, (int)((dynamic)modules[1]).Lines);

            var read = f.Session.Execute(new Request { Command = "read_module", Project = f.Project.Name,
                Module = "module1" });
            Assert.IsTrue(read.Ok);
            Assert.AreEqual("Alpha\r\nBeta", (string)((dynamic)read.Data).Code);
            Assert.AreEqual(Sha("Alpha\r\nBeta"), (string)((dynamic)read.Data).Sha256);
            var empty = f.Session.Execute(new Request { Command = "read_module", Project = f.Project.Name,
                Module = "Class1" });
            Assert.AreEqual("", (string)((dynamic)empty.Data).Code);
            Assert.AreEqual(Sha(""), (string)((dynamic)empty.Data).Sha256);
        }

        [TestMethod]
        public void ReadModuleRejectsMissingOrUnknownModule()
        {
            var f = Create();
            Assert.ThrowsException<ArgumentException>(() => f.Session.Execute(new Request {
                Command = "read_module", Project = f.Project.Name }));
            Assert.ThrowsException<InvalidOperationException>(() => f.Session.Execute(new Request {
                Command = "read_module", Project = f.Project.Name, Module = "Missing" }));
        }

        [TestMethod]
        public void ReplaceLinesRejectsMissingHashRangeAndChangedCodeWithoutMutation()
        {
            var f = Create();
            var request = new Request { Command = "replace_lines", Project = f.Project.Name,
                Module = "Module1", StartLine = 1, Count = 1, Text = "Changed" };
            Assert.IsFalse(f.Session.Execute(request).Ok);
            request.ExpectedSha256 = Sha("stale");
            StringAssert.Contains(f.Session.Execute(request).Error, "module changed");
            request.ExpectedSha256 = Sha("Alpha\r\nBeta");
            request.StartLine = 4;
            StringAssert.Contains(f.Session.Execute(request).Error, "outside the module");
            request.StartLine = 1;
            request.Count = -1;
            StringAssert.Contains(f.Session.Execute(request).Error, "Invalid line range");
            Assert.AreEqual("Alpha\r\nBeta", f.Project.VBComponents.Items[0].CodeModule.Code);
        }

        [TestMethod]
        public void ReplaceLinesRequiresDesignMode()
        {
            var f = Create(mode: 1);
            var response = f.Session.Execute(new Request { Command = "replace_lines", Project = f.Project.Name,
                Module = "Module1", StartLine = 1, Count = 1, Text = "Changed",
                ExpectedSha256 = Sha("Alpha\r\nBeta") });
            Assert.IsFalse(response.Ok);
            StringAssert.Contains(response.Error, "design mode");
            Assert.AreEqual("Alpha\r\nBeta", f.Project.VBComponents.Items[0].CodeModule.Code);
        }

        [TestMethod]
        public void ReplaceLinesDeletesAndInsertsThenReportsCurrentHash()
        {
            var f = Create();
            var response = f.Session.Execute(new Request { Command = "replace_lines", Project = f.Project.Name,
                Module = "Module1", StartLine = 2, Count = 1, Text = "Gamma",
                ExpectedSha256 = Sha("Alpha\r\nBeta") });
            Assert.IsTrue(response.Ok);
            Assert.AreEqual("Alpha\r\nGamma", f.Project.VBComponents.Items[0].CodeModule.Code);
            Assert.AreEqual(Sha("Alpha\r\nGamma"), (string)((dynamic)response.Data).Sha256);
            Assert.AreEqual(2, (int)((dynamic)response.Data).Lines);
        }

        [TestMethod]
        public void CreateModuleAndClassValidateModeAndDuplicateIdentity()
        {
            var f = Create();
            var request = new Request { Command = "create_module", Project = f.Project.Name,
                Module = "Module1", ExpectedMode = 2 };
            Assert.ThrowsException<InvalidOperationException>(() => f.Session.Execute(request));
            request.Module = "2Bad";
            Assert.ThrowsException<ArgumentException>(() => f.Session.Execute(request));
            request.Module = "NewModule";
            request.ExpectedMode = 1;
            Assert.ThrowsException<ArgumentException>(() => f.Session.Execute(request));
            request.ExpectedMode = 2;
            var created = f.Session.Execute(request);
            Assert.IsTrue(created.Ok);
            Assert.AreEqual(1, (int)((dynamic)created.Data).Type);
            Assert.AreEqual(Sha(""), (string)((dynamic)created.Data).Sha256);

            request.Command = "create_class";
            request.Module = "NewClass";
            var createdClass = f.Session.Execute(request);
            Assert.AreEqual(2, (int)((dynamic)createdClass.Data).Type);
            Assert.AreEqual(3, f.Project.VBComponents.Items.Count);
        }

        [TestMethod]
        public void CreateComponentRollsBackOnlyNewItemWhenRenameFails()
        {
            var f = Create();
            f.Project.VBComponents.RejectedName = "Denied";
            var request = new Request { Command = "create_module", Project = f.Project.Name,
                Module = "Denied", ExpectedMode = 2 };
            Assert.ThrowsException<InvalidOperationException>(() => f.Session.Execute(request));
            Assert.AreEqual(1, f.Project.VBComponents.Items.Count);
            Assert.AreEqual("Module1", f.Project.VBComponents.Items[0].Name);
            Assert.AreEqual(1, f.Project.VBComponents.RemoveCount);
        }

        private sealed class Fixture
        {
            public FakeVbe Vbe;
            public FakeProject Project;
            public VbeSession Session;
        }

        public sealed class FakeVbe
        {
            public List<FakeProject> VBProjects { get; } = new List<FakeProject>();
        }

        public sealed class FakeProject
        {
            private string fileName;
            public string Name { get; set; }
            public int Mode { get; set; }
            public bool ThrowFileName { get; set; }
            public string FileName
            {
                get { if (ThrowFileName) throw new InvalidOperationException("Unsaved"); return fileName; }
                set { fileName = value; }
            }
            public FakeComponents VBComponents { get; } = new FakeComponents();
        }

        public sealed class FakeComponents : IEnumerable<FakeComponent>
        {
            public List<FakeComponent> Items { get; } = new List<FakeComponent>();
            public string RejectedName { get; set; }
            public int RemoveCount { get; private set; }
            public FakeComponent Add(int type)
            {
                var component = new FakeComponent { Name = "Temporary", Type = type,
                    CodeModule = new FakeModule("") };
                component.RejectedName = RejectedName;
                Items.Add(component);
                return component;
            }
            public void Remove(FakeComponent component) { RemoveCount++; Items.Remove(component); }
            public IEnumerator<FakeComponent> GetEnumerator() { return Items.GetEnumerator(); }
            IEnumerator IEnumerable.GetEnumerator() { return GetEnumerator(); }
        }

        public sealed class FakeComponent
        {
            private string name;
            public string RejectedName { get; set; }
            public string Name
            {
                get { return name; }
                set
                {
                    if (value == RejectedName) throw new InvalidOperationException("Rejected by VBE");
                    name = value;
                }
            }
            public int Type { get; set; }
            public FakeModule CodeModule { get; set; }
        }

        public sealed class FakeModule
        {
            private readonly List<string> lines;
            public FakeModule(string code)
            {
                lines = code.Length == 0 ? new List<string>() :
                    code.Split(new[] { "\r\n" }, StringSplitOptions.None).ToList();
                Lines = new FakeLines(this);
            }
            public int CountOfLines => lines.Count;
            public string Code => string.Join("\r\n", lines);
            public FakeLines Lines { get; }
            public void DeleteLines(int start, int count) { lines.RemoveRange(start - 1, count); }
            public void InsertLines(int start, string text)
            {
                lines.InsertRange(start - 1, text.Split(new[] { "\r\n" }, StringSplitOptions.None));
            }
            public sealed class FakeLines
            {
                private readonly FakeModule module;
                public FakeLines(FakeModule module) { this.module = module; }
                public string this[int start, int count] =>
                    string.Join("\r\n", module.lines.Skip(start - 1).Take(count));
            }
        }
    }
}
