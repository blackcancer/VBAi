using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;

namespace VBAi.Tests.Unit
{
    [TestClass]
    public sealed class VbeProjectLegacyHelpMetadataTests
    {
        public sealed class Project : VbeProjectComponentsTests.FakeProject
        {
            internal int FileWrites, ContextWrites;
            internal Exception TypeError;
            internal int ProjectType = 100;
            private string file = "raw COM ??";
            private int context = 321;
            public int Type { get { if (TypeError != null) throw TypeError; return ProjectType; } }
            public string HelpFile { get => file; set { FileWrites++; file = value; } }
            public int HelpContextID { get => context; set { ContextWrites++; context = value; } }
        }

        private sealed class Fixture
        {
            internal readonly Project Project = new Project();
            internal readonly VbeProjectComponents Service;
            internal int TypedPreparations;
            internal Fixture(string host, bool native = true)
            {
                var vbe = new VbeProjectComponentsTests.FakeVbe(); vbe.VBProjects.Add(Project);
                Service = new VbeProjectComponents(vbe, new VbeForms(vbe));
                Service.LegacyHelpMetadataProcessName = () => host;
                Service.LegacyHelpMetadataNativeProject = target => native && ReferenceEquals(target, Project);
                Service.AccessHelpContextFactory = (target, window) => { TypedPreparations++; throw new InvalidOperationException("Typed setter must not be prepared."); };
            }
            internal Request Request(string property, object value)
            {
                dynamic state = Service.ProjectProperties(Project.Name);
                return new Request { Project = Project.Name, Property = property, Value = value, ExpectedProjectVersion = state.Version };
            }
            internal List<VbePropertyInfo> Properties() => (List<VbePropertyInfo>)((dynamic)Service.ProjectProperties(Project.Name)).Properties;
        }

        [DataTestMethod]
        [DataRow("MSACCESS", "HelpFile")]
        [DataRow("MSPUB", "HelpFile")]
        [DataRow("msaccess", "helpcontextid")]
        [DataRow("mspub", "HELPFILE")]
        [DataRow("MSACCESS", "HelpContextID")]
        [DataRow("MSPUB", "HelpContextID")]
        public void NativeHostHelpWritesAreRefusedBeforeEitherSetter(string host, string property)
        {
            var f = new Fixture(host); var request = f.Request(property, property.Equals("HelpFile", StringComparison.OrdinalIgnoreCase) ? (object)"C:\\new.chm" : 322);
            var error = Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetProjectProperty(request));
            StringAssert.Contains(error.Message, "no setter was invoked"); StringAssert.Contains(error.Message, "read_project_general");
            StringAssert.Contains(error.Message, "set_project_general");
            Assert.AreEqual(0, f.Project.FileWrites + f.Project.ContextWrites + f.TypedPreparations);
            Assert.AreEqual("raw COM ??", f.Project.HelpFile); Assert.AreEqual(321, f.Project.HelpContextID);
            Assert.AreEqual(request.ExpectedProjectVersion, (string)((dynamic)f.Service.ProjectProperties(f.Project.Name)).Version);
        }

        [DataTestMethod]
        [DataRow("EXCEL")]
        [DataRow("WINWORD")]
        [DataRow("POWERPNT")]
        [DataRow("SLDWORKS")]
        [DataRow("OUTLOOK")]
        [DataRow("MSPUB.exe")]
        [DataRow("")]
        [DataRow(null)]
        public void OtherProcessesRetainTheirExistingScalarRoute(string host)
        {
            var f = new Fixture(host); f.Service.SetProjectProperty(f.Request("HelpFile", "C:\\new.chm"));
            Assert.AreEqual(1, f.Project.FileWrites); Assert.AreEqual(0, f.TypedPreparations);
            Assert.AreEqual("DescriptorCandidateUnverified", f.Properties().Single(p => p.Name == "HelpFile").SetterStatus);
        }

        [DataTestMethod]
        [DataRow(0)]
        [DataRow(1)]
        [DataRow(101)]
        public void OtherProjectTypesRetainTheirExistingScalarRoute(int type)
        {
            var f = new Fixture("MSPUB"); f.Project.ProjectType = type;
            f.Service.SetProjectProperty(f.Request("HelpFile", "C:\\new.chm"));
            Assert.AreEqual(1, f.Project.FileWrites);
        }

        [TestMethod]
        public void ManagedFixturesAndOtherPropertiesAreNotBlocked()
        {
            var managed = new Fixture("MSACCESS", false);
            managed.Service.SetProjectProperty(managed.Request("HelpContextID", 322));
            Assert.AreEqual(1, managed.Project.ContextWrites);
            var native = new Fixture("MSPUB"); native.Service.SetProjectProperty(native.Request("Description", "new"));
            Assert.AreEqual("new", native.Project.Description); Assert.AreEqual(0, native.Project.FileWrites + native.Project.ContextWrites);
        }

        [TestMethod]
        public void UnknownNativeTypeAndStaleVersionNeverAdmitASetter()
        {
            var f = new Fixture("MSACCESS"); var request = f.Request("HelpFile", "C:\\new.chm");
            f.Project.TypeError = new InvalidOperationException("Unknown native project type.");
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetProjectProperty(request));
            f.Project.TypeError = null; request.ExpectedProjectVersion = "stale";
            Assert.ThrowsException<InvalidOperationException>(() => f.Service.SetProjectProperty(request));
            Assert.AreEqual(0, f.Project.FileWrites + f.Project.ContextWrites + f.TypedPreparations);
        }

        [TestMethod]
        public void CapabilityIsDeterministicAndPreservesDescriptorAndRawReadback()
        {
            var f = new Fixture("MSPUB");
            f.Service.LegacyHelpMetadataNativeProject = target => false;
            var baseline = f.Properties();
            f.Service.LegacyHelpMetadataNativeProject = target => ReferenceEquals(target, f.Project);
            var properties = f.Properties();
            string version = (string)((dynamic)f.Service.ProjectProperties(f.Project.Name)).Version;
            Assert.AreEqual(version, (string)((dynamic)f.Service.ProjectProperties(f.Project.Name)).Version);
            foreach (string name in new[] { "HelpFile", "HelpContextID" })
            {
                var property = properties.Single(p => p.Name == name);
                var raw = baseline.Single(p => p.Name == name);
                Assert.AreEqual(false, property.ReadOnly);
                Assert.AreEqual(raw.ReadOnly, property.ReadOnly); Assert.AreEqual(raw.Value, property.Value);
                Assert.AreEqual(raw.Type, property.Type); Assert.AreEqual(raw.Digest, property.Digest); Assert.AreEqual(raw.Error, property.Error);
                Assert.AreEqual(VbeProjectComponents.LegacyHelpMetadataSetterStatus, property.SetterStatus);
                StringAssert.Contains(property.Display, "set_project_general"); Assert.IsNull(property.Error);
            }
            Assert.AreEqual("raw COM ??", properties.Single(p => p.Name == "HelpFile").Value);
            Assert.AreEqual(321, properties.Single(p => p.Name == "HelpContextID").Value);
            Assert.AreEqual("DescriptorCandidateUnverified", properties.Single(p => p.Name == "Description").SetterStatus);
            Assert.AreEqual(0, f.Project.FileWrites + f.Project.ContextWrites);
        }
    }
}
