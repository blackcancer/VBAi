namespace VBAi.Tests.Unit
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using VBAi;

    public sealed partial class VbeProjectLifecycleTests
    {
        private static Fixture Create()
        {
            var vbe = new LifecycleVbe();
            vbe.VBProjects.Items.Add(new LifecycleProject { Name = "Existing", FileName = "" });
            return new Fixture { Vbe = vbe, Service = new VbeProjectComponents(vbe, new VbeForms(vbe)) };
        }
        private sealed class Fixture
        {
            public LifecycleVbe Vbe;
            public VbeProjectComponents Service;
            public Request CollectionRequest(string path = null)
            {
                dynamic state = Service.ProjectCollectionState();
                return new Request { Path = path, ExpectedProjectVersion = state.Version };
            }
            public Request CloseRequest(LifecycleProject project)
            {
                dynamic state = Service.ProjectProperties(project.Name);
                return new Request { Project = project.Name, ExpectedHostPath = project.FileName,
                    ExpectedProjectVersion = state.Version };
            }
        }
        public sealed class LifecycleVbe
        {
            public LifecycleCollection VBProjects { get; } = new LifecycleCollection();
        }
        public sealed class LifecycleProject : VbeProjectComponentsTests.FakeProject
        {
            public int Type { get; set; } = 101;
            public int Protection { get; set; }
            public int? FileNameError { get; set; }
            public bool ClrPathNotFound { get; set; }
            public override string FileName
            {
                get
                {
                    if (ClrPathNotFound) throw new System.IO.DirectoryNotFoundException("Path not found");
                    if (FileNameError.HasValue) throw new System.Runtime.InteropServices.COMException("Path not found", FileNameError.Value);
                    return base.FileName;
                }
                set { base.FileName = value; }
            }
        }
        public sealed class LifecycleCollection : IEnumerable<LifecycleProject>
        {
            public List<LifecycleProject> Items { get; } = new List<LifecycleProject>();
            public int Attempts;
            public int LastType;
            public Action After;
            public bool Fail;
            public bool NoChange;
            public bool WrongPath;
            public bool ReturnExisting;
            public bool ThrowEnumeration;
            public LifecycleProject Add(int type) { LastType = type; return Insert(""); }
            public LifecycleProject Open(string path) { return Insert(path); }
            private LifecycleProject Insert(string path)
            {
                Attempts++;
                var project = new LifecycleProject { Name = "Added", FileName = WrongPath ? "" : path };
                if (!NoChange) Items.Add(project);
                After?.Invoke();
                if (Fail) throw new InvalidOperationException("Native failure after mutation");
                return ReturnExisting ? Items[0] : project;
            }
            public void Remove(LifecycleProject project)
            {
                Attempts++;
                if (!NoChange) Items.Remove(project);
                After?.Invoke();
                if (Fail) throw new InvalidOperationException("Native removal failed");
            }
            public IEnumerator<LifecycleProject> GetEnumerator()
            {
                if (ThrowEnumeration) throw new InvalidOperationException("Native enumeration failed");
                return Items.GetEnumerator();
            }
            IEnumerator IEnumerable.GetEnumerator() { return GetEnumerator(); }
        }
    }
}
