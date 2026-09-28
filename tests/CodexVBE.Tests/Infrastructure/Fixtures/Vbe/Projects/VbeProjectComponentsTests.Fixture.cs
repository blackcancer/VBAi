namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using CodexVBE;

    public sealed partial class VbeProjectComponentsTests
    {
        private static VbeProjectComponents Service(FakeProject project)
        {
            var vbe = new FakeVbe();
            vbe.VBProjects.Add(project);
            return new VbeProjectComponents(vbe, new VbeForms(vbe));
        }

        public sealed class FakeVbe
        {
            public List<FakeProject> VBProjects { get; } = new List<FakeProject>();
        }

        public class FakeProject
        {
            public string Name { get; set; } = "VBAProject";
            public string Description { get; set; } = "Original";
            private string fileName = @"C:\fixture\Book.xlsm";
            public bool FailFileName { get; set; }
            public string FileName { get { if (FailFileName) throw new InvalidOperationException("Project filename unavailable"); return fileName; } set { fileName = value; } }
            public int Mode { get; set; } = 2;
            public bool Saved { get; set; } = true;
            public FakeComponentCollection VBComponents { get; } = new FakeComponentCollection();
            public List<object> References { get; } = new List<object>();
        }

        public sealed class FakeComponentCollection : IEnumerable<FakeComponent>
        {
            private readonly List<FakeComponent> items = new List<FakeComponent>();
            public int RemoveAttempts { get; private set; }
            public int ImportAttempts { get; private set; }
            public bool ImportThenThrow { get; set; }
            public int ImportAddedCount { get; set; } = 1;
            public bool FailImportedEnumeration { get; set; }
            public bool FailImportedCodeReadback { get; set; }
            public bool FailImportedProjectReadback { get; set; }

            public void Add(FakeComponent component)
            {
                items.Add(component);
            }

            public void Remove(FakeComponent component)
            {
                RemoveAttempts++;
                items.Remove(component);
            }

            public void Import(string path)
            {
                ImportAttempts++;
                for (int index = 0; index < ImportAddedCount; index++)
                {
                    var imported = new FakeComponent(index == 0 ? "ImportedModule" : "ImportedModule" + index, 1);
                    imported.CodeModule.FailRead = FailImportedCodeReadback;
                    imported.FailName = FailImportedProjectReadback;
                    items.Add(imported);
                }
                if (ImportThenThrow)
                    throw new InvalidOperationException("COM error after add");
            }

            public IEnumerator<FakeComponent> GetEnumerator()
            {
                if (FailImportedEnumeration && ImportAttempts > 0) throw new InvalidOperationException("Enumeration unavailable after import");
                return items.GetEnumerator();
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }
        }

        public class FakeComponent
        {
            public FakeComponent(string name, int type)
            {
                Name = name;
                Type = type;
                if (type == 3)
                {
                    Designer = new FakeDesigner();
                    Properties.Add(new FakeProperty { Name = "Caption", Value = name });
                    Properties.Add(new FakeProperty { Name = "Width", Value = 240d });
                    Properties.Add(new FakeProperty { Name = "Height", Value = 180d });
                }
            }

            private string name;
            public bool IgnoreRename { get; set; }
            public bool FailName { get; set; }
            public string Name { get { if (FailName) throw new InvalidOperationException("Name unavailable"); return name; } set { if (!IgnoreRename) name = value; } }
            public string Description { get; set; } = "Original";
            public int Type { get; set; }
            public FakePropertyCollection Properties { get; } = new FakePropertyCollection();
            public FakeDesigner Designer { get; }
            public FakeCodeModule CodeModule { get; } = new FakeCodeModule();
            public int ExportAttempts { get; private set; }
            public bool WriteExportFile { get; set; }

            public void Export(string path)
            {
                ExportAttempts++;
                if (WriteExportFile)
                    File.WriteAllText(path, "Attribute VB_Name = \"" + Name + "\"");
            }
        }

        public sealed class FakeDesigner
        {
            public List<object> Controls { get; } = new List<object>();
        }

        public sealed class FakePropertyCollection : IEnumerable<FakeProperty>
        {
            private readonly List<FakeProperty> items = new List<FakeProperty>
            {
                new FakeProperty
                {
                    Name = "Instancing",
                    Value = 1
                }
            };
            public FakeProperty Item(string name)
            {
                return items.Single(p => p.Name == name);
            }

            public void Add(FakeProperty property)
            {
                items.Add(property);
            }

            public IEnumerator<FakeProperty> GetEnumerator()
            {
                return items.GetEnumerator();
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }
        }

        public sealed class FakeProperty
        {
            public string Name { get; set; }
            private object value;
            public bool IgnoreSet { get; set; }
            public bool FailRead { get; set; }
            public object Value { get { if (FailRead) throw new InvalidOperationException("Property unavailable"); return value; } set { if (!IgnoreSet) this.value = value; } }
        }

        public sealed class FakeCodeModule
        {
            public int LineCount { get; set; } = 1;
            public bool FailRead { get; set; }
            public int CountOfLines
            {
                get
                {
                    if (FailRead) throw new InvalidOperationException("Code unavailable");
                    return LineCount;
                }
            }

            public string Lines(int start, int count)
            {
                return "Option Explicit";
            }
        }
    }
}
