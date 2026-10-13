using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using VBAi.Tests.Integration;

namespace VBAi.Tests.Unit
{
    /// <summary>Exercises the complete empty-Normal admission and metadata-failure matrix without Office or personal file reads.</summary>
    [TestClass, TestCategory("Unit")]
    public sealed class OfficeVbeFixtureWordNormalTemplateMetadataTests
    {
        [TestMethod]
        public void ExactEmptyNormalIsReviewedUsingMetadataOnlyAndEveryAcquisitionIsBalanced()
        {
            var fake = new Template(); var released = new List<object>();
            object result = OfficeVbeFixture.ReviewEmptyStartupNormalTemplate(fake, fake, released.Add);
            Assert.IsNotNull(result);
            Assert.AreEqual(5, released.Count);
            Assert.AreSame(fake.Project.Components.Item.Module, released[0]);
            Assert.AreSame(fake.Project.Components.Item, released[1]);
            Assert.AreSame(fake.Project.Components, released[3]);
            Assert.AreSame(fake.Project, released[4]);
            Assert.IsFalse(released.Contains(fake), "The template arguments are borrowed.");
            Assert.AreEqual(1, fake.Project.Components.Item.Module.MetadataReads);
        }

        [DataTestMethod, DataRow("null-template"), DataRow("null-normal"), DataRow("foreign")]
        public void WrongNormalIdentityRefusesBeforeReadingItsProject(string state)
        {
            var fake = new Template(); var released = new List<object>();
            Assert.ThrowsException<AssertFailedException>(() => OfficeVbeFixture.ReviewEmptyStartupNormalTemplate(
                state == "null-template" ? null : fake, state == "null-normal" ? null : state == "foreign" ? new Template() : fake, released.Add));
            Assert.AreEqual(0, fake.ProjectReads); Assert.AreEqual(0, released.Count);
        }

        [DataTestMethod]
        [DataRow("project-null")]
        [DataRow("name")]
        [DataRow("running")]
        [DataRow("protected")]
        [DataRow("empty-collection")]
        [DataRow("extra-component")]
        [DataRow("module-type")]
        [DataRow("module-name")]
        [DataRow("module-null")]
        [DataRow("code")]
        public void NonemptyProtectedRunningOrIncompleteNormalCannotAdmitUnsuppression(string fault)
        {
            var fake = new Template { Fault = fault }; var released = new List<object>();
            if (fault == "name") fake.Project.Name = "Other";
            if (fault == "running") fake.Project.Mode = 0;
            if (fault == "protected") fake.Project.Protection = 1;
            if (fault == "empty-collection") fake.Project.Components.Count = 0;
            if (fault == "extra-component") fake.Project.Components.Count = 2;
            if (fault == "module-type") fake.Project.Components.Item.Type = 1;
            if (fault == "module-name") fake.Project.Components.Item.Name = "Other";
            if (fault == "module-null") fake.Project.Components.Item.Module = null;
            if (fault == "code") fake.Project.Components.Item.Module.LinesCount = 1;
            Assert.ThrowsException<AssertFailedException>(() => OfficeVbeFixture.ReviewEmptyStartupNormalTemplate(fake, fake, released.Add));
            Assert.IsFalse(released.Contains(fake));
            Assert.AreEqual(released.Count, new HashSet<object>(released).Count, "Each getter acquisition is released once.");
        }

        [DataTestMethod, DataRow("project"), DataRow("template-path"), DataRow("components"), DataRow("enumerator"), DataRow("module"), DataRow("lines")]
        public void NativeGetterFailuresRemainFailuresAndNeverReadCodeTextOrDisk(string fault)
        {
            var fake = new Template { Fault = fault };
            fake.Project.ComponentsFault = fault == "components";
            fake.Project.Components.EnumeratorFault = fault == "enumerator";
            fake.Project.Components.Item.ModuleFault = fault == "module";
            fake.Project.Components.Item.Module.MetadataFault = fault == "lines";
            Assert.ThrowsException<IOException>(() => OfficeVbeFixture.ReviewEmptyStartupNormalTemplate(fake, fake));
        }

        [TestMethod]
        public void FailedReleaseIsNotRepeatedAndCannotReturnSuccessfulAdmission()
        {
            var fake = new Template(); int releases = 0;
            Assert.ThrowsException<IOException>(() => OfficeVbeFixture.ReviewEmptyStartupNormalTemplate(fake, fake,
                value => { releases++; throw new IOException("release"); }));
            Assert.AreEqual(1, releases);
        }

        public sealed class Template
        {
            public string Fault;
            public int ProjectReads;
            public readonly Project Project = new Project();
            public object VBProject { get { ProjectReads++; if (Fault == "project") throw new IOException("project"); return Fault == "project-null" ? null : Project; } }
            public string FullName { get { if (Fault == "template-path") throw new IOException("path"); return @"C:\metadata-only\Normal.dotm"; } }
        }
        public sealed class Project
        {
            public string Name { get; set; } = "Normal";
            public int Mode { get; set; } = 2;
            public int Protection { get; set; }
            public bool ComponentsFault;
            public readonly Components Components = new Components();
            public object VBComponents => ComponentsFault ? throw new IOException("components") : (object)Components;
            public string FileName => throw new AssertFailedException("The VBA backing file must not be inspected.");
        }
        public sealed class Components : IEnumerable
        {
            public int Count = 1;
            public bool EnumeratorFault;
            public readonly Component Item = new Component();
            public IEnumerator GetEnumerator()
            {
                if (EnumeratorFault) throw new IOException("enumerator");
                var items = new List<object>(); for (int i = 0; i < Count; i++) items.Add(i == 0 ? Item : new Component());
                return items.GetEnumerator();
            }
        }
        public sealed class Component
        {
            public int Type { get; set; } = 100;
            public string Name { get; set; } = "ThisDocument";
            public bool ModuleFault;
            public Module Module = new Module();
            public object CodeModule => ModuleFault ? throw new IOException("module") : Module;
        }
        public sealed class Module
        {
            public int LinesCount, MetadataReads;
            public bool MetadataFault;
            public int CountOfLines { get { MetadataReads++; if (MetadataFault) throw new IOException("metadata"); return LinesCount; } }
            public string Lines(int start, int count) => throw new AssertFailedException("Code text must never be read.");
        }
    }
}
