namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Security.Cryptography;
    using System.Text;
    using VBAi;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    public sealed partial class VbeSessionTests
    {
        private static Fixture Create(string code = "Alpha\r\nBeta", int mode = 2)
        {
            var project = new FakeProject
            {
                Name = "VBAProject",
                FileName = @"C:\Temp\Host.xlsm",
                Mode = mode
            };
            project.VBComponents.Items.Add(new FakeComponent { Name = "Module1", Type = 1, CodeModule = new FakeModule(code) });
            var vbe = new FakeVbe();
            vbe.VBProjects.Add(project);
            return new Fixture
            {
                Vbe = vbe,
                Project = project,
                Session = new VbeSession(vbe)
            };
        }

        private static string Sha(string code)
        {
            using (var sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(code))).Replace("-", "").ToLowerInvariant();
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
            public FakeProject ActiveVBProject { get; set; }
        }

        public sealed class FakeProject
        {
            public int Protection { get; set; }
            private string fileName;
            public string Name { get; set; }
            public int Mode { get; set; }
            public bool Saved { get; set; } = true;
            public bool ThrowFileName { get; set; }
            public bool ThrowDirectoryNotFound { get; set; }

            public string FileName
            {
                get
                {
                    if (ThrowFileName)
                        throw new InvalidOperationException("Unsaved");
                    if (ThrowDirectoryNotFound)
                        throw new DirectoryNotFoundException("Host document is unavailable");
                    return fileName;
                }

                set
                {
                    fileName = value;
                }
            }

            public FakeComponents VBComponents { get; } = new FakeComponents();
            public FakeReferences References { get; } = new FakeReferences();
        }

        public sealed class FakeReferences : IEnumerable<FakeReference>
        {
            public List<FakeReference> Items { get; } = new List<FakeReference>();

            public void AddFromGuid(string guid, int major, int minor)
            {
                Items.Add(new FakeReference { GUID = guid, Major = major, Minor = minor, Name = "AddedByGuid" });
            }

            public void AddFromFile(string path)
            {
                Items.Add(new FakeReference { GUID = "{FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF}", Major = 1, Minor = 0, Name = "AddedByFile", FullPath = path });
            }

            public void Remove(FakeReference reference)
            {
                Items.Remove(reference);
            }

            public IEnumerator<FakeReference> GetEnumerator()
            {
                return Items.GetEnumerator();
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }
        }

        public sealed class FakeReference
        {
            private string name;
            private string fullPath;
            public string GUID { get; set; }
            public int Major { get; set; }
            public int Minor { get; set; }
            public bool IsBroken { get; set; }
            public bool BuiltIn { get; set; }
            public bool ThrowMetadata { get; set; }

            public string Name
            {
                get
                {
                    if (ThrowMetadata)
                        throw new InvalidOperationException("Unavailable");
                    return name;
                }

                set
                {
                    name = value;
                }
            }

            public string FullPath
            {
                get
                {
                    if (ThrowMetadata)
                        throw new InvalidOperationException("Unavailable");
                    return fullPath;
                }

                set
                {
                    fullPath = value;
                }
            }
        }

        public sealed class FakeComponents : IEnumerable<FakeComponent>
        {
            public List<FakeComponent> Items { get; } = new List<FakeComponent>();
            public string RejectedName { get; set; }
            public string ForcedReadbackName { get; set; }
            public int? ForcedAddedType { get; set; }
            public bool FailRemoval { get; set; }
            public int RemoveCount { get; private set; }

            public FakeComponent Add(int type)
            {
                var component = new FakeComponent
                {
                    Name = "Temporary",
                    Type = type,
                    CodeModule = new FakeModule("")
                };
                component.RejectedName = RejectedName;
                component.ReadbackName = ForcedReadbackName;
                if (ForcedAddedType.HasValue) component.Type = ForcedAddedType.Value;
                Items.Add(component);
                return component;
            }

            public void Remove(FakeComponent component)
            {
                RemoveCount++;
                if (FailRemoval) throw new InvalidOperationException("Removal rejected by VBE");
                Items.Remove(component);
            }

            public IEnumerator<FakeComponent> GetEnumerator()
            {
                return Items.GetEnumerator();
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }
        }

        public sealed class FakeComponent
        {
            private string name;
            public string RejectedName { get; set; }
            public string ReadbackName { get; set; }

            public string Name
            {
                get
                {
                    return ReadbackName ?? name;
                }

                set
                {
                    if (value == RejectedName)
                        throw new InvalidOperationException("Rejected by VBE");
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
                lines = code.Length == 0 ? new List<string>() : code.Split(new[] { "\r\n" }, StringSplitOptions.None).ToList();
                Lines = new FakeLines(this);
            }

            public int CountOfLines => lines.Count;
            public string Code => string.Join("\r\n", lines);
            public FakeLines Lines { get; }

            public void DeleteLines(int start, int count)
            {
                lines.RemoveRange(start - 1, count);
            }

            public int InsertFailuresRemaining { get; set; }
            public void InsertLines(int start, string text)
            {
                if (InsertFailuresRemaining > 0) { InsertFailuresRemaining--; throw new InvalidOperationException("Simulated insert failure"); }
                lines.InsertRange(start - 1, text.Split(new[] { "\r\n" }, StringSplitOptions.None));
            }

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
