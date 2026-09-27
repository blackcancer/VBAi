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

    public sealed partial class VbeReferenceTypesBranchTests
    {
        private string path;
        private readonly Guid guid = Guid.NewGuid();
        [TestInitialize]
        public void CreateReferenceFile()
        {
            path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".tlb");
            File.WriteAllBytes(path, new byte[] { 0 });
        }

        [TestCleanup]
        public void RemoveReferenceFile()
        {
            if (path != null && File.Exists(path))
                File.Delete(path);
        }

        private VbeReferenceTypes Reader(Library library, Func<string, ITypeLib> loadFile = null, Func<Guid, ushort, ushort, ITypeLib> loadRegistered = null)
        {
            var project = new VbeReferenceTypesTests.FakeProject
            {
                Name = "Project"
            };
            project.References.Add(new VbeReferenceTypesTests.FakeReference { Name = "Synthetic", GUID = guid.ToString("B"), Major = 1, Minor = 0, FullPath = path });
            var host = new VbeReferenceTypesTests.FakeVbe();
            host.VBProjects.Add(project);
            return new VbeReferenceTypes(host, loadFile ?? (file => library), loadRegistered ?? ((id, major, minor) => library));
        }

        private Request Request()
        {
            return new Request
            {
                Project = "Project",
                Guid = guid.ToString("B"),
                Major = 1,
                Minor = 0
            };
        }

        private static string Identity(VbeReferenceTypes reader, Request request)
        {
            object page = reader.ListTypes(request);
            return (string)Value(((System.Collections.IList)Value(page, "Types"))[0], "TypeIdentity");
        }

        private static object Value(object value, string name)
        {
            return value.GetType().GetProperty(name).GetValue(value, null);
        }

        private sealed class Library : ITypeLib
        {
            public Library(Guid guid)
            {
                Guid = guid;
            }

            public Guid Guid { get; }
            public int Count { get; set; }
            public TypeInfo[] Types { get; set; } = new TypeInfo[0];

            public int GetTypeInfoCount()
            {
                return Count;
            }

            public void GetTypeInfo(int index, out ITypeInfo info)
            {
                if (index >= Types.Length || Types[index] == null)
                    throw new COMException("Missing type.");
                info = Types[index];
            }

            public void GetLibAttr(out IntPtr pointer)
            {
                pointer = Native(new TYPELIBATTR { guid = Guid, wMajorVerNum = 1, wMinorVerNum = 0 });
            }

            public void ReleaseTLibAttr(IntPtr pointer)
            {
                Marshal.FreeHGlobal(pointer);
            }

            public void GetTypeInfoType(int index, out TYPEKIND kind)
            {
                kind = TYPEKIND.TKIND_INTERFACE;
            }

            public void GetTypeInfoOfGuid(ref Guid id, out ITypeInfo info)
            {
                info = null;
            }

            public void GetTypeComp(out ITypeComp typeComp)
            {
                typeComp = null;
            }

            public void GetDocumentation(int index, out string name, out string description, out int helpContext, out string helpFile)
            {
                name = description = helpFile = null;
                helpContext = 0;
            }

            public bool IsName(string name, int hash)
            {
                return false;
            }

            public void FindName(string name, int hash, ITypeInfo[] infos, int[] memberIds, ref short found)
            {
                found = 0;
            }
        }

        private sealed class TypeInfo : ITypeInfo
        {
            public string Name { get; set; }
            public TYPEKIND Kind { get; set; } = TYPEKIND.TKIND_INTERFACE;
            public int Functions { get; set; }
            public int Variables { get; set; }
            public string[] FunctionNames { get; set; } = new string[0];
            public string[] VariableNames { get; set; } = new string[0];
            public TypeInfo[] Interfaces { get; set; } = new TypeInfo[0];
            public IMPLTYPEFLAGS[] Flags { get; set; } = new IMPLTYPEFLAGS[0];

            public void GetTypeAttr(out IntPtr pointer)
            {
                pointer = Native(new TYPEATTR { guid = Guid.Empty, typekind = Kind, cFuncs = checked((short)Functions), cVars = checked((short)Variables), cImplTypes = checked((short)Interfaces.Length) });
            }

            public void ReleaseTypeAttr(IntPtr pointer)
            {
                Marshal.FreeHGlobal(pointer);
            }

            public void GetDocumentation(int index, out string name, out string description, out int helpContext, out string helpFile)
            {
                if (index >= 200 && VariableNames[index - 200] == null)
                    throw new COMException("Missing variable documentation.");
                name = index == -1 ? Name : index >= 200 ? VariableNames[index - 200] : null;
                description = helpFile = null;
                helpContext = 0;
            }

            public void GetFuncDesc(int index, out IntPtr pointer)
            {
                pointer = Native(new FUNCDESC { memid = 100 + index, cParams = 0, invkind = INVOKEKIND.INVOKE_FUNC });
            }

            public void ReleaseFuncDesc(IntPtr pointer)
            {
                Marshal.FreeHGlobal(pointer);
            }

            public void GetVarDesc(int index, out IntPtr pointer)
            {
                pointer = Native(new VARDESC { memid = 200 + index, varkind = VARKIND.VAR_PERINSTANCE });
            }

            public void ReleaseVarDesc(IntPtr pointer)
            {
                Marshal.FreeHGlobal(pointer);
            }

            public void GetNames(int memberId, string[] names, int maximum, out int count)
            {
                string name = FunctionNames[memberId - 100];
                count = name == null ? 0 : 1;
                if (count > 0)
                    names[0] = name;
            }

            public void GetImplTypeFlags(int index, out IMPLTYPEFLAGS flags)
            {
                flags = Flags[index];
            }

            public void GetRefTypeOfImplType(int index, out int href)
            {
                href = index;
            }

            public void GetRefTypeInfo(int href, out ITypeInfo info)
            {
                info = Interfaces[href];
            }

            public void GetTypeComp(out ITypeComp typeComp)
            {
                typeComp = null;
            }

            public void GetIDsOfNames(string[] names, int count, int[] memberIds)
            {
            }

            public void Invoke(object instance, int memberId, short flags, ref DISPPARAMS parameters, IntPtr result, IntPtr exception, out int argumentError)
            {
                argumentError = 0;
            }

            public void GetDllEntry(int memberId, INVOKEKIND kind, IntPtr dllName, IntPtr name, IntPtr ordinal)
            {
            }

            public void AddressOfMember(int memberId, INVOKEKIND kind, out IntPtr pointer)
            {
                pointer = IntPtr.Zero;
            }

            public void CreateInstance(object outer, ref Guid id, out object result)
            {
                result = null;
            }

            public void GetMops(int memberId, out string mops)
            {
                mops = null;
            }

            public void GetContainingTypeLib(out ITypeLib library, out int index)
            {
                library = null;
                index = 0;
            }
        }

        private static IntPtr Native<T>(T value)
            where T : struct
        {
            IntPtr pointer = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(T)));
            Marshal.StructureToPtr(value, pointer, false);
            return pointer;
        }
    }
}
