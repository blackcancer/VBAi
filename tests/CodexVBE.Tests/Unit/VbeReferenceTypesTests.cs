using System;
using System.Collections.Generic;
using System.Collections;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using TYPELIBATTR = System.Runtime.InteropServices.ComTypes.TYPELIBATTR;
using CodexVBE;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CodexVBE.Tests.Unit
{
    [TestClass]
    [TestCategory("Unit")]
    public sealed class VbeReferenceTypesTests
    {
        private const string StdOlePath = @"C:\Windows\System32\stdole2.tlb";

        [DllImport("oleaut32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        private static extern void LoadTypeLibEx(string path, int registrationKind, out ITypeLib library);

        [TestMethod]
        public void PaginationAndReferenceIdentityAreValidatedBeforeOpeningComLibrary()
        {
            var reader = new VbeReferenceTypes(new FakeVbe());
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => reader.ListTypes(new Request { Offset = -1 }));
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => reader.ListTypes(new Request { Limit = 51 }));
            Assert.ThrowsException<ArgumentException>(() => reader.ListTypes(new Request {
                Project = "Projet", Guid = "not-a-guid" }));
            Assert.ThrowsException<ArgumentException>(() => reader.ListMembers(new Request {
                Project = "Projet", Guid = Guid.NewGuid().ToString("B") }));
        }

        [TestMethod]
        public void SelectedTypeLibraryCanListTypesAndMembersWithoutRegistration()
        {
            if (!File.Exists(StdOlePath)) Assert.Inconclusive("Windows stdole2.tlb is unavailable.");
            var reference = ReferenceFor(StdOlePath);
            var project = new FakeProject { Name = "Projet" };
            project.References.Add(reference);
            var host = new FakeVbe();
            host.VBProjects.Add(project);
            var reader = new VbeReferenceTypes(host);
            var request = new Request { Project = "Projet", Guid = reference.GUID,
                Major = reference.Major, Minor = reference.Minor, Limit = 50 };
            object page = reader.ListTypes(request);
            Assert.IsTrue(Convert.ToInt32(Prop(page, "TotalTypes")) > 0);
            Assert.IsTrue(Convert.ToInt32(Prop(page, "Returned")) > 0);
            Assert.AreEqual("SelectedReferenceFile: LoadTypeLibEx(REGKIND_NONE)", Prop(page, "Source"));
            Assert.AreEqual("stdole2", Prop(Prop(page, "Reference"), "Name"));

            object selected = null;
            foreach (object type in (IEnumerable)Prop(page, "Types"))
                if ((string)Prop(type, "Kind") != "TKIND_COCLASS" &&
                    Convert.ToInt32(Prop(type, "FunctionCount")) + Convert.ToInt32(Prop(type, "VariableCount")) > 0)
                { selected = type; break; }
            Assert.IsNotNull(selected, "The selected type library should expose at least one member.");
            request.TypeIndex = Convert.ToInt32(Prop(selected, "TypeIndex"));
            request.TypeIdentity = (string)Prop(selected, "TypeIdentity");
            request.Limit = 2;
            object members = reader.ListMembers(request);
            Assert.IsTrue(Convert.ToInt32(Prop(members, "TotalMembers")) > 0);
            Assert.IsTrue(Convert.ToInt32(Prop(members, "Returned")) > 0);
            Assert.AreEqual(request.TypeIdentity, Prop(Prop(members, "Type"), "Identity"));

            request.TypeIdentity = "stale";
            Assert.ThrowsException<InvalidOperationException>(() => reader.ListMembers(request));
            request.TypeIdentity = (string)Prop(selected, "TypeIdentity");
            request.TypeIndex = Convert.ToInt32(Prop(page, "TotalTypes"));
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => reader.ListMembers(request));
        }

        [TestMethod]
        public void BrokenMissingAndAmbiguousReferencesAreRejected()
        {
            if (!File.Exists(StdOlePath)) Assert.Inconclusive("Windows stdole2.tlb is unavailable.");
            var reference = ReferenceFor(StdOlePath);
            var project = new FakeProject { Name = "Projet" };
            project.References.Add(reference);
            var host = new FakeVbe();
            host.VBProjects.Add(project);
            var reader = new VbeReferenceTypes(host);
            var request = new Request { Project = "Projet", Guid = reference.GUID,
                Major = reference.Major, Minor = reference.Minor };
            reference.IsBroken = true;
            Assert.ThrowsException<InvalidOperationException>(() => reader.ListTypes(request));
            reference.IsBroken = false;
            reference.FullPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".tlb");
            Assert.ThrowsException<FileNotFoundException>(() => reader.ListTypes(request));
            reference.FullPath = StdOlePath;
            project.References.Add(ReferenceFor(StdOlePath));
            Assert.ThrowsException<InvalidOperationException>(() => reader.ListTypes(request));
        }

        [TestMethod]
        public void ReferenceMustBeSelectedWithAnExactValidIdentityAndAbsoluteFile()
        {
            if (!File.Exists(StdOlePath)) Assert.Inconclusive("Windows stdole2.tlb is unavailable.");
            var reference = ReferenceFor(StdOlePath);
            var project = new FakeProject { Name = "Projet" };
            project.References.Add(reference);
            var host = new FakeVbe();
            host.VBProjects.Add(project);
            var reader = new VbeReferenceTypes(host);
            var request = new Request { Project = project.Name, Guid = reference.GUID,
                Major = reference.Major, Minor = reference.Minor };
            request.Minor++;
            Assert.ThrowsException<InvalidOperationException>(() => reader.ListTypes(request));
            request.Minor--;
            reference.FullPath = "relative.tlb";
            Assert.ThrowsException<FileNotFoundException>(() => reader.ListTypes(request));
            reference.FullPath = StdOlePath;
            reference.GUID = Guid.NewGuid().ToString("B");
            Assert.ThrowsException<InvalidOperationException>(() => reader.ListTypes(request));
            request.Guid = reference.GUID;
            var error = Assert.ThrowsException<InvalidOperationException>(() => reader.ListTypes(request));
            StringAssert.Contains(error.Message, "identity differs");
        }

        [TestMethod]
        public void TypeAndMemberPagesExposeStableOffsetsAndEmptyFinalPages()
        {
            if (!File.Exists(StdOlePath)) Assert.Inconclusive("Windows stdole2.tlb is unavailable.");
            var reference = ReferenceFor(StdOlePath);
            var project = new FakeProject { Name = "Projet" };
            project.References.Add(reference);
            var host = new FakeVbe();
            host.VBProjects.Add(project);
            var reader = new VbeReferenceTypes(host);
            var request = new Request { Project = project.Name, Guid = reference.GUID,
                Major = reference.Major, Minor = reference.Minor, Limit = 1 };
            object first = reader.ListTypes(request);
            int typeCount = Convert.ToInt32(Prop(first, "TotalTypes"));
            Assert.IsTrue(typeCount > 1);
            Assert.AreEqual(1, Prop(first, "Returned"));
            Assert.AreEqual(true, Prop(first, "HasMore"));
            request.Offset = typeCount;
            object empty = reader.ListTypes(request);
            Assert.AreEqual(0, Prop(empty, "Returned"));
            Assert.AreEqual(false, Prop(empty, "HasMore"));

            request.Offset = 0;
            request.Limit = 50;
            object types = reader.ListTypes(request);
            object selected = null;
            foreach (object type in (IEnumerable)Prop(types, "Types"))
                if ((string)Prop(type, "Kind") != "TKIND_COCLASS" &&
                    Convert.ToInt32(Prop(type, "FunctionCount")) + Convert.ToInt32(Prop(type, "VariableCount")) > 0)
                { selected = type; break; }
            Assert.IsNotNull(selected);
            request.TypeIndex = Convert.ToInt32(Prop(selected, "TypeIndex"));
            request.TypeIdentity = (string)Prop(selected, "TypeIdentity");
            request.Limit = 1;
            object firstMember = reader.ListMembers(request);
            int memberCount = Convert.ToInt32(Prop(firstMember, "TotalMembers"));
            Assert.IsTrue(memberCount > 0);
            Assert.AreEqual(1, Prop(firstMember, "Returned"));
            request.Offset = memberCount;
            object emptyMembers = reader.ListMembers(request);
            Assert.AreEqual(0, Prop(emptyMembers, "Returned"));
            Assert.AreEqual(false, Prop(emptyMembers, "HasMore"));
        }

        private static FakeReference ReferenceFor(string path)
        {
            ITypeLib library;
            LoadTypeLibEx(path, 2, out library);
            IntPtr pointer = IntPtr.Zero;
            try
            {
                library.GetLibAttr(out pointer);
                var attr = (TYPELIBATTR)Marshal.PtrToStructure(pointer, typeof(TYPELIBATTR));
                return new FakeReference { Name = "stdole2", GUID = attr.guid.ToString("B"),
                    Major = attr.wMajorVerNum, Minor = attr.wMinorVerNum, FullPath = path };
            }
            finally
            {
                if (pointer != IntPtr.Zero) library.ReleaseTLibAttr(pointer);
                Marshal.ReleaseComObject(library);
            }
        }

        private static object Prop(object value, string name)
        {
            return value.GetType().GetProperty(name).GetValue(value, null);
        }

        public sealed class FakeVbe { public List<FakeProject> VBProjects { get; } = new List<FakeProject>(); }
        public sealed class FakeProject
        {
            public string Name { get; set; }
            public List<FakeReference> References { get; } = new List<FakeReference>();
        }
        public sealed class FakeReference
        {
            public string Name { get; set; }
            public string GUID { get; set; }
            public int Major { get; set; }
            public int Minor { get; set; }
            public bool IsBroken { get; set; }
            public string FullPath { get; set; }
        }
    }
}
