namespace CodexVBE.Tests.Unit
{
    using System;
    using System.IO;
    using System.Runtime.InteropServices;
    using System.Runtime.InteropServices.ComTypes;
    using TYPEKIND = System.Runtime.InteropServices.ComTypes.TYPEKIND;
    using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
    using TYPELIBATTR = System.Runtime.InteropServices.ComTypes.TYPELIBATTR;
    using FUNCDESC = System.Runtime.InteropServices.ComTypes.FUNCDESC;
    using VARDESC = System.Runtime.InteropServices.ComTypes.VARDESC;
    using VARKIND = System.Runtime.InteropServices.ComTypes.VARKIND;
    using IMPLTYPEFLAGS = System.Runtime.InteropServices.ComTypes.IMPLTYPEFLAGS;
    using INVOKEKIND = System.Runtime.InteropServices.ComTypes.INVOKEKIND;
    using DISPPARAMS = System.Runtime.InteropServices.ComTypes.DISPPARAMS;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class VbeReferenceTypesBranchTests
    {
        [TestMethod]
        public void RejectsBothTypeCountBoundsAndRecordsIndividualTypeFailure()
        {
            var library = new Library(guid)
            {
                Count = -1
            };
            var reader = Reader(library);
            Assert.ThrowsException<InvalidOperationException>(() => reader.ListTypes(Request()));
            library.Count = 10001;
            Assert.ThrowsException<InvalidOperationException>(() => reader.ListTypes(Request()));
            library.Count = 2;
            library.Types = new[]
            {
                new TypeInfo
                {
                    Name = "Usable"
                },
                null
            };
            object page = reader.ListTypes(Request());
            Assert.AreEqual(2, Value(page, "Returned"));
            Assert.IsNotNull(Value(((System.Collections.IList)Value(page, "Types"))[1], "Error"));
        }

        [TestMethod]
        public void RejectsMemberCountAndTypeIndexBounds()
        {
            var info = new TypeInfo
            {
                Name = "Interface",
                Functions = 10001
            };
            var library = new Library(guid)
            {
                Count = 1,
                Types = new[]
                {
                    info
                }
            };
            var reader = Reader(library);
            var request = Request();
            request.TypeIdentity = Identity(reader, request);
            library.Count = -1;
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => reader.ListMembers(request));
            library.Count = 10001;
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => reader.ListMembers(request));
            library.Count = 1;
            request.TypeIndex = -1;
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => reader.ListMembers(request));
            request.TypeIndex = 1;
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => reader.ListMembers(request));
            request.TypeIndex = 0;
            Assert.ThrowsException<InvalidOperationException>(() => reader.ListMembers(request));
        }

        [TestMethod]
        public void ResolvesDefaultInterfaceAndSkipsSourceInterfaces()
        {
            var source = new TypeInfo
            {
                Name = "Events"
            };
            var fallback = new TypeInfo
            {
                Name = "Fallback",
                Variables = 1
            };
            var preferred = new TypeInfo
            {
                Name = "Preferred",
                Functions = 1
            };
            var coclass = new TypeInfo
            {
                Name = "Class",
                Kind = TYPEKIND.TKIND_COCLASS,
                Interfaces = new[]
                {
                    source,
                    fallback,
                    preferred
                },
                Flags = new[]
                {
                    IMPLTYPEFLAGS.IMPLTYPEFLAG_FSOURCE,
                    (IMPLTYPEFLAGS)0,
                    IMPLTYPEFLAGS.IMPLTYPEFLAG_FDEFAULT
                }
            };
            var library = new Library(guid)
            {
                Count = 1,
                Types = new[]
                {
                    coclass
                }
            };
            var reader = Reader(library);
            var request = Request();
            request.TypeIdentity = Identity(reader, request);
            object members = reader.ListMembers(request);
            Assert.AreEqual("DefaultNonSourceInterface", Value(members, "Resolution"));
            Assert.AreEqual("Preferred", Value(Value(members, "MemberInterface"), "Name"));
            Assert.AreEqual(1, Value(members, "TotalMembers"));
            coclass.Flags = new[]
            {
                IMPLTYPEFLAGS.IMPLTYPEFLAG_FSOURCE,
                (IMPLTYPEFLAGS)0,
                (IMPLTYPEFLAGS)0
            };
            members = reader.ListMembers(request);
            Assert.AreEqual("Fallback", Value(Value(members, "MemberInterface"), "Name"));
            coclass.Flags = new[]
            {
                IMPLTYPEFLAGS.IMPLTYPEFLAG_FSOURCE,
                IMPLTYPEFLAGS.IMPLTYPEFLAG_FSOURCE,
                IMPLTYPEFLAGS.IMPLTYPEFLAG_FSOURCE
            };
            members = reader.ListMembers(request);
            Assert.AreEqual("NoDefaultNonSourceInterface", Value(members, "Resolution"));
            Assert.AreEqual(0, Value(members, "TotalMembers"));
        }

        [TestMethod]
        public void ReadsFunctionsAndVariablesAndRecordsIndividualMemberFailures()
        {
            var info = new TypeInfo
            {
                Name = "Mixed",
                Functions = 2,
                Variables = 2,
                FunctionNames = new[]
                {
                    "DoWork",
                    null
                },
                VariableNames = new[]
                {
                    "Value",
                    null
                }
            };
            var library = new Library(guid)
            {
                Count = 1,
                Types = new[]
                {
                    info
                }
            };
            var reader = Reader(library);
            var request = Request();
            request.TypeIdentity = Identity(reader, request);
            object page = reader.ListMembers(request);
            var members = (System.Collections.IList)Value(page, "Members");
            Assert.AreEqual("Function", Value(members[0], "Kind"));
            Assert.AreEqual("DoWork", Value(members[0], "Name"));
            Assert.IsNull(Value(members[1], "Name"));
            Assert.AreEqual("Variable", Value(members[2], "Kind"));
            Assert.AreEqual("Value", Value(members[2], "Name"));
            Assert.IsNotNull(Value(members[3], "Error"));
        }

        [TestMethod]
        public void UsesRegistryFallbackAndReportsBothLoaderErrors()
        {
            var library = new Library(guid)
            {
                Count = 0
            };
            var reader = Reader(library, file =>
            {
                throw new COMException("file", unchecked((int)0x8002801D));
            }, (id, major, minor) => library);
            object page = reader.ListTypes(Request());
            Assert.AreEqual("RegisteredTypeLibrary: LoadRegTypeLib", Value(page, "Source"));
            Assert.IsNotNull(Value(page, "FallbackError"));
            reader = Reader(library, file =>
            {
                throw new COMException("file", unchecked((int)0x8002801D));
            }, (id, major, minor) =>
            {
                throw new COMException("registry", unchecked((int)0x8002801C));
            });
            var error = Assert.ThrowsException<InvalidOperationException>(() => reader.ListTypes(Request()));
            StringAssert.Contains(error.Message, "0x8002801D");
            StringAssert.Contains(error.Message, "0x8002801C");
            reader = Reader(library, file => null, (id, major, minor) => library);
            error = Assert.ThrowsException<InvalidOperationException>(() => reader.ListTypes(Request()));
            StringAssert.Contains(error.Message, "returned no library");
        }

        [TestMethod]
        public void RejectsPageBoundsAndAcceptsDefaultPageSize()
        {
            var library = new Library(guid)
            {
                Count = 0
            };
            var reader = Reader(library);
            var request = Request();
            request.Offset = 10001;
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => reader.ListTypes(request));
            request.Offset = 0;
            request.Limit = -1;
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => reader.ListTypes(request));
            request.Limit = 0;
            Assert.AreEqual(0, Value(reader.ListTypes(request), "Returned"));
        }

        [TestMethod]
        public void RequiresBothLibraryLoaders()
        {
            Assert.ThrowsException<ArgumentNullException>(() => new VbeReferenceTypes(new object (), null, (id, major, minor) => null));
            Assert.ThrowsException<ArgumentNullException>(() => new VbeReferenceTypes(new object (), file => null, null));
        }
    }
}

namespace CodexVBE.Tests.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Collections;
    using System.IO;
    using System.Runtime.InteropServices;
    using System.Runtime.InteropServices.ComTypes;
    using TYPELIBATTR = System.Runtime.InteropServices.ComTypes.TYPELIBATTR;
    using CodexVBE;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    [TestClass]
    [TestCategory("Unit")]
    public sealed partial class VbeReferenceTypesTests
    {
        [TestMethod]
        public void PaginationAndReferenceIdentityAreValidatedBeforeOpeningComLibrary()
        {
            var reader = new VbeReferenceTypes(new FakeVbe());
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => reader.ListTypes(new Request { Offset = -1 }));
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => reader.ListTypes(new Request { Limit = 51 }));
            Assert.ThrowsException<ArgumentException>(() => reader.ListTypes(new Request { Project = "Projet", Guid = "not-a-guid" }));
            Assert.ThrowsException<ArgumentException>(() => reader.ListMembers(new Request { Project = "Projet", Guid = Guid.NewGuid().ToString("B") }));
        }

        [TestMethod]
        public void SelectedTypeLibraryCanListTypesAndMembersWithoutRegistration()
        {
            if (!File.Exists(StdOlePath))
                Assert.Inconclusive("Windows stdole2.tlb is unavailable.");
            var reference = ReferenceFor(StdOlePath);
            var project = new FakeProject
            {
                Name = "Projet"
            };
            project.References.Add(reference);
            var host = new FakeVbe();
            host.VBProjects.Add(project);
            var reader = new VbeReferenceTypes(host);
            var request = new Request
            {
                Project = "Projet",
                Guid = reference.GUID,
                Major = reference.Major,
                Minor = reference.Minor,
                Limit = 50
            };
            object page = reader.ListTypes(request);
            Assert.IsTrue(Convert.ToInt32(Prop(page, "TotalTypes")) > 0);
            Assert.IsTrue(Convert.ToInt32(Prop(page, "Returned")) > 0);
            Assert.AreEqual("SelectedReferenceFile: LoadTypeLibEx(REGKIND_NONE)", Prop(page, "Source"));
            Assert.AreEqual("stdole2", Prop(Prop(page, "Reference"), "Name"));
            object selected = null;
            foreach (object type in (IEnumerable)Prop(page, "Types"))
                if ((string)Prop(type, "Kind") != "TKIND_COCLASS" && Convert.ToInt32(Prop(type, "FunctionCount")) + Convert.ToInt32(Prop(type, "VariableCount")) > 0)
                {
                    selected = type;
                    break;
                }

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
            if (!File.Exists(StdOlePath))
                Assert.Inconclusive("Windows stdole2.tlb is unavailable.");
            var reference = ReferenceFor(StdOlePath);
            var project = new FakeProject
            {
                Name = "Projet"
            };
            project.References.Add(reference);
            var host = new FakeVbe();
            host.VBProjects.Add(project);
            var reader = new VbeReferenceTypes(host);
            var request = new Request
            {
                Project = "Projet",
                Guid = reference.GUID,
                Major = reference.Major,
                Minor = reference.Minor
            };
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
            if (!File.Exists(StdOlePath))
                Assert.Inconclusive("Windows stdole2.tlb is unavailable.");
            var reference = ReferenceFor(StdOlePath);
            var project = new FakeProject
            {
                Name = "Projet"
            };
            project.References.Add(reference);
            var host = new FakeVbe();
            host.VBProjects.Add(project);
            var reader = new VbeReferenceTypes(host);
            var request = new Request
            {
                Project = project.Name,
                Guid = reference.GUID,
                Major = reference.Major,
                Minor = reference.Minor
            };
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
            if (!File.Exists(StdOlePath))
                Assert.Inconclusive("Windows stdole2.tlb is unavailable.");
            var reference = ReferenceFor(StdOlePath);
            var project = new FakeProject
            {
                Name = "Projet"
            };
            project.References.Add(reference);
            var host = new FakeVbe();
            host.VBProjects.Add(project);
            var reader = new VbeReferenceTypes(host);
            var request = new Request
            {
                Project = project.Name,
                Guid = reference.GUID,
                Major = reference.Major,
                Minor = reference.Minor,
                Limit = 1
            };
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
                if ((string)Prop(type, "Kind") != "TKIND_COCLASS" && Convert.ToInt32(Prop(type, "FunctionCount")) + Convert.ToInt32(Prop(type, "VariableCount")) > 0)
                {
                    selected = type;
                    break;
                }

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

        [TestMethod]
        public void CoclassMembersResolveThroughItsDefaultNonSourceInterface()
        {
            if (!File.Exists(StdOlePath))
                Assert.Inconclusive("Windows stdole2.tlb is unavailable.");
            var reference = ReferenceFor(StdOlePath);
            var project = new FakeProject
            {
                Name = "Projet"
            };
            project.References.Add(reference);
            var host = new FakeVbe();
            host.VBProjects.Add(project);
            var reader = new VbeReferenceTypes(host);
            var request = new Request
            {
                Project = project.Name,
                Guid = reference.GUID,
                Major = reference.Major,
                Minor = reference.Minor,
                Limit = 50
            };
            object page = reader.ListTypes(request);
            object selected = null;
            foreach (object type in (IEnumerable)Prop(page, "Types"))
                if ((string)Prop(type, "Kind") == "TKIND_COCLASS")
                {
                    selected = type;
                    break;
                }

            if (selected == null)
                Assert.Inconclusive("stdole2.tlb exposes no coclass on this system.");
            request.TypeIndex = Convert.ToInt32(Prop(selected, "TypeIndex"));
            request.TypeIdentity = (string)Prop(selected, "TypeIdentity");
            object members = reader.ListMembers(request);
            Assert.AreEqual("DefaultNonSourceInterface", Prop(members, "Resolution"));
            Assert.IsTrue(Convert.ToInt32(Prop(members, "TotalMembers")) > 0);
            Assert.AreEqual(request.TypeIdentity, Prop(Prop(members, "Type"), "Identity"));
        }
    }
}
