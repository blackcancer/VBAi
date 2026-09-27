namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Security.Cryptography;
    using System.Text;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class VbeCodeNavigationCoverageTests
    {
        private static VbeCodeNavigation CreateNavigation(string first, string second)
        {
            var host = new FakeVbe();
            var project = new FakeProject
            {
                Name = "Projet",
                Mode = 0
            };
            project.VBComponents.Add(new FakeComponent { Name = "Module1", Type = 1, CodeModule = new FakeModule(first) });
            project.VBComponents.Add(new FakeComponent { Name = "Class1", Type = 2, CodeModule = new FakeModule(second) });
            host.VBProjects.Add(project);
            return new VbeCodeNavigation(host, new VbeForms(host));
        }

        private static string Hash(string source)
        {
            using (var sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(source))).Replace("-", "").ToLowerInvariant();
        }

        public sealed class FakeVbe
        {
            public List<FakeProject> VBProjects { get; } = new List<FakeProject>();
        }

        public sealed class FakeProject
        {
            public string Name { get; set; }
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
            private readonly string[] lines;
            public FakeModule(string code)
            {
                lines = string.IsNullOrEmpty(code) ? new string[0] : code.Split(new[] { "\r\n" }, StringSplitOptions.None);
                Lines = new FakeLines(this);
            }

            public int CountOfLines => lines.Length;
            public FakeLines Lines { get; }

            public sealed class FakeLines
            {
                private readonly FakeModule module;
                public FakeLines(FakeModule module)
                {
                    this.module = module;
                }

                public string this[int start, int count] => string.Join("\r\n", module.lines.Skip(start - 1).Take(count));
            }
        }
    }
}
