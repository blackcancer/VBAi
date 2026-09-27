using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Security.Cryptography;
using System.Text;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using TYPELIBATTR = System.Runtime.InteropServices.ComTypes.TYPELIBATTR;
using FUNCDESC = System.Runtime.InteropServices.ComTypes.FUNCDESC;
using VARDESC = System.Runtime.InteropServices.ComTypes.VARDESC;
using TYPEKIND = System.Runtime.InteropServices.ComTypes.TYPEKIND;
using IMPLTYPEFLAGS = System.Runtime.InteropServices.ComTypes.IMPLTYPEFLAGS;

namespace CodexVBE
{
    // Reads only type metadata from a selected VBProject reference file. No
    // object instance is created and REGKIND_NONE never registers a type library.
    internal sealed class VbeReferenceTypes
    {
        private const int MaxPageSize = 50;
        private const int MaxTypes = 10000;
        private const int MaxMembers = 10000;
        private readonly dynamic vbe;
        private readonly Func<string, ITypeLib> loadFile;
        private readonly Func<Guid, ushort, ushort, ITypeLib> loadRegistered;

        [DllImport("oleaut32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        private static extern void LoadTypeLibEx(string file, int regKind, out ITypeLib typeLib);

        [DllImport("oleaut32.dll", PreserveSig = false)]
        private static extern void LoadRegTypeLib(ref Guid guid, ushort major, ushort minor,
            int lcid, out ITypeLib typeLib);

        public VbeReferenceTypes(object vbe)
            : this(vbe, LoadSelectedFile, LoadRegisteredLibrary) { }

        internal VbeReferenceTypes(object vbe, Func<string, ITypeLib> loadFile,
            Func<Guid, ushort, ushort, ITypeLib> loadRegistered)
        {
            this.vbe = vbe;
            this.loadFile = loadFile ?? throw new ArgumentNullException(nameof(loadFile));
            this.loadRegistered = loadRegistered ?? throw new ArgumentNullException(nameof(loadRegistered));
        }

        private static ITypeLib LoadSelectedFile(string path)
        {
            ITypeLib library;
            LoadTypeLibEx(path, 2, out library); // REGKIND_NONE
            return library;
        }

        private static ITypeLib LoadRegisteredLibrary(Guid guid, ushort major, ushort minor)
        {
            ITypeLib library;
            LoadRegTypeLib(ref guid, major, minor, 0, out library);
            return library;
        }

        public object ListTypes(Request request)
        {
            int offset = ValidatePage(request);
            ReferenceSource reference = FindReference(request);
            LoadedLibrary loaded = OpenLibrary(reference);
            ITypeLib library = loaded.Library;
            int count = library.GetTypeInfoCount();
            if (count < 0 || count > MaxTypes)
                throw new InvalidOperationException("Type library type count exceeds the inspection limit.");
            var types = new List<object>();
            int end = Math.Min(count, offset + PageSize(request));
            for (int index = offset; index < end; index++)
            {
                try
                {
                    ITypeInfo typeInfo;
                    library.GetTypeInfo(index, out typeInfo);
                    types.Add(ReadType(typeInfo, index, reference));
                }
                catch (Exception ex)
                {
                    types.Add(new { TypeIndex = index, Error = Error(ex) });
                }
            }
            return new { Project = request.Project, Reference = reference,
                Source = loaded.Source, FallbackError = loaded.FallbackError,
                TotalTypes = count, Offset = offset, Returned = types.Count, HasMore = end < count,
                Types = types };
        }

        public object ListMembers(Request request)
        {
            ValidatePage(request);
            if (string.IsNullOrWhiteSpace(request.TypeIdentity))
                throw new ArgumentException("TypeIdentity from list_reference_types is required.");
            ReferenceSource reference = FindReference(request);
            LoadedLibrary loaded = OpenLibrary(reference);
            ITypeLib library = loaded.Library;
            int typeCount = library.GetTypeInfoCount();
            if (typeCount < 0 || typeCount > MaxTypes || request.TypeIndex < 0 ||
                request.TypeIndex >= typeCount)
                throw new ArgumentOutOfRangeException("TypeIndex");
            ITypeInfo typeInfo;
            library.GetTypeInfo(request.TypeIndex, out typeInfo);
            TypeMetadata metadata = ReadTypeMetadata(typeInfo, request.TypeIndex, reference);
            if (!string.Equals(metadata.Identity, request.TypeIdentity, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The type changed since list_reference_types; list it again.");
            ITypeInfo memberType = typeInfo;
            TypeMetadata memberMetadata = metadata;
            string resolution = "DirectType";
            if (metadata.Kind == TYPEKIND.TKIND_COCLASS.ToString())
            {
                memberType = DefaultNonSourceInterface(typeInfo);
                if (memberType != null)
                {
                    memberMetadata = ReadTypeMetadata(memberType, request.TypeIndex, reference);
                    resolution = "DefaultNonSourceInterface";
                }
                else
                    resolution = "NoDefaultNonSourceInterface";
            }
            int count = memberType == null ? 0 : memberMetadata.FunctionCount + memberMetadata.VariableCount;
            if (count > MaxMembers)
                throw new InvalidOperationException("Type member count exceeds the inspection limit.");
            int end = Math.Min(count, request.Offset + PageSize(request));
            var members = new List<object>();
            for (int index = request.Offset; index < end; index++)
            {
                try
                {
                    members.Add(index < memberMetadata.FunctionCount
                        ? ReadFunction(memberType, index)
                        : ReadVariable(memberType, index - memberMetadata.FunctionCount, index));
                }
                catch (Exception ex)
                {
                    members.Add(new { MemberIndex = index, Error = Error(ex) });
                }
            }
            return new { Project = request.Project, Reference = reference,
                Type = metadata, MemberInterface = memberMetadata,
                Resolution = resolution, Source = loaded.Source,
                FallbackError = loaded.FallbackError,
                Scope = "Raw COM type members and accessors; inherited interfaces and VBE Object Browser filtering are not expanded.",
                TotalMembers = count, Offset = request.Offset, Returned = members.Count,
                HasMore = end < count, Members = members };
        }

        private static int ValidatePage(Request request)
        {
            if (request.Offset < 0 || request.Offset > MaxTypes)
                throw new ArgumentOutOfRangeException("Offset", "Offset must be between 0 and 10000.");
            if (request.Limit < 0 || request.Limit > MaxPageSize)
                throw new ArgumentOutOfRangeException("Limit", "Limit must be between 1 and 50; omit for 50.");
            return request.Offset;
        }

        private static int PageSize(Request request) { return request.Limit == 0 ? MaxPageSize : request.Limit; }

        private ReferenceSource FindReference(Request request)
        {
            Guid expectedGuid;
            if (string.IsNullOrWhiteSpace(request.Project) ||
                !Guid.TryParse(request.Guid, out expectedGuid) || request.Major < 0 || request.Minor < 0 ||
                request.Major > ushort.MaxValue || request.Minor > ushort.MaxValue)
                throw new ArgumentException("Project, reference Guid, Major and Minor are required.");
            dynamic project = VbeProjectResolver.Resolve(vbe, request.Project);
            ReferenceSource found = null;
            foreach (dynamic reference in project.References)
            {
                Guid actualGuid;
                if (!Guid.TryParse((string)reference.GUID, out actualGuid) ||
                    actualGuid != expectedGuid || (int)reference.Major != request.Major ||
                    (int)reference.Minor != request.Minor) continue;
                if (found != null) throw new InvalidOperationException("Reference identity is ambiguous.");
                if ((bool)reference.IsBroken)
                    throw new InvalidOperationException("The selected reference is broken.");
                string path = (string)reference.FullPath;
                if (string.IsNullOrWhiteSpace(path) || !Path.IsPathRooted(path) || !File.Exists(path))
                    throw new FileNotFoundException("Reference type library file is unavailable.", path);
                var file = new FileInfo(path);
                found = new ReferenceSource { Name = (string)reference.Name,
                    Guid = actualGuid.ToString("B"), Major = request.Major, Minor = request.Minor,
                    FullPath = path, FileLength = file.Length,
                    FileLastWriteUtc = file.LastWriteTimeUtc.ToString("O") };
            }
            if (found == null) throw new InvalidOperationException("The exact reference is not selected by this project.");
            return found;
        }

        private LoadedLibrary OpenLibrary(ReferenceSource reference)
        {
            ITypeLib library;
            string source = "SelectedReferenceFile: LoadTypeLibEx(REGKIND_NONE)";
            object fallbackError = null;
            try { library = loadFile(reference.FullPath); }
            catch (COMException ex)
            {
                fallbackError = Error(ex);
                Guid guid = Guid.Parse(reference.Guid);
                try
                {
                    library = loadRegistered(guid, (ushort)reference.Major, (ushort)reference.Minor);
                }
                catch (COMException registryError)
                {
                    throw new InvalidOperationException("Type library failed from selected file (HRESULT " +
                        HResultHex(ex) + ") and registry (HRESULT " + HResultHex(registryError) + ").",
                        registryError);
                }
                source = "RegisteredTypeLibrary: LoadRegTypeLib";
            }
            if (library == null) throw new InvalidOperationException("LoadTypeLibEx returned no library.");
            IntPtr pointer = IntPtr.Zero;
            try
            {
                library.GetLibAttr(out pointer);
                var attr = (TYPELIBATTR)Marshal.PtrToStructure(pointer, typeof(TYPELIBATTR));
                if (!string.Equals(attr.guid.ToString("B"), reference.Guid, StringComparison.OrdinalIgnoreCase) ||
                    attr.wMajorVerNum != reference.Major || attr.wMinorVerNum != reference.Minor)
                    throw new InvalidOperationException("The file type library identity differs from the selected VBE reference.");
            }
            finally { if (pointer != IntPtr.Zero) library.ReleaseTLibAttr(pointer); }
            return new LoadedLibrary { Library = library, Source = source,
                FallbackError = fallbackError };
        }

        private static ITypeInfo DefaultNonSourceInterface(ITypeInfo coclass)
        {
            IntPtr pointer = IntPtr.Zero;
            try
            {
                coclass.GetTypeAttr(out pointer);
                var attr = (TYPEATTR)Marshal.PtrToStructure(pointer, typeof(TYPEATTR));
                int selected = -1;
                for (int index = 0; index < attr.cImplTypes; index++)
                {
                    IMPLTYPEFLAGS flags;
                    coclass.GetImplTypeFlags(index, out flags);
                    if ((flags & IMPLTYPEFLAGS.IMPLTYPEFLAG_FSOURCE) != 0) continue;
                    if ((flags & IMPLTYPEFLAGS.IMPLTYPEFLAG_FDEFAULT) != 0)
                    { selected = index; break; }
                    if (selected == -1) selected = index;
                }
                if (selected == -1) return null;
                int href;
                coclass.GetRefTypeOfImplType(selected, out href);
                ITypeInfo result;
                coclass.GetRefTypeInfo(href, out result);
                return result;
            }
            finally { if (pointer != IntPtr.Zero) coclass.ReleaseTypeAttr(pointer); }
        }

        private static object ReadType(ITypeInfo typeInfo, int index, ReferenceSource reference)
        {
            TypeMetadata type = ReadTypeMetadata(typeInfo, index, reference);
            return new { type.TypeIndex, type.Name, type.Guid, type.Kind,
                type.FunctionCount, type.VariableCount, type.ImplementedInterfaceCount,
                TypeIdentity = type.Identity };
        }

        private static TypeMetadata ReadTypeMetadata(ITypeInfo typeInfo, int index, ReferenceSource reference)
        {
            IntPtr pointer = IntPtr.Zero;
            try
            {
                typeInfo.GetTypeAttr(out pointer);
                var attr = (TYPEATTR)Marshal.PtrToStructure(pointer, typeof(TYPEATTR));
                string name, description, helpFile;
                int helpContext;
                typeInfo.GetDocumentation(-1, out name, out description, out helpContext, out helpFile);
                string identity = Hash(reference.Guid + "|" + reference.Major + "|" + reference.Minor +
                    "|" + reference.FileLength + "|" + reference.FileLastWriteUtc + "|" +
                    index + "|" + attr.guid + "|" + attr.typekind + "|" + name);
                return new TypeMetadata { TypeIndex = index, Name = name, Guid = attr.guid.ToString("B"),
                    Kind = attr.typekind.ToString(), FunctionCount = attr.cFuncs,
                    VariableCount = attr.cVars, ImplementedInterfaceCount = attr.cImplTypes,
                    Identity = identity };
            }
            finally { if (pointer != IntPtr.Zero) typeInfo.ReleaseTypeAttr(pointer); }
        }

        private static object ReadFunction(ITypeInfo typeInfo, int index)
        {
            IntPtr pointer = IntPtr.Zero;
            try
            {
                typeInfo.GetFuncDesc(index, out pointer);
                var descriptor = (FUNCDESC)Marshal.PtrToStructure(pointer, typeof(FUNCDESC));
                string[] names = new string[Math.Min(32, Math.Max(1, descriptor.cParams + 1))];
                int count;
                typeInfo.GetNames(descriptor.memid, names, names.Length, out count);
                return new { MemberIndex = index, Kind = "Function", Name = count > 0 ? names[0] : null,
                    DispId = descriptor.memid, InvocationKind = descriptor.invkind.ToString(),
                    FunctionFlags = descriptor.wFuncFlags,
                    ParameterCount = (int)descriptor.cParams,
                    ParameterNames = names.Skip(1).Take(Math.Max(0, count - 1)).ToArray(),
                    ParameterNamesTruncated = descriptor.cParams + 1 > names.Length };
            }
            finally { if (pointer != IntPtr.Zero) typeInfo.ReleaseFuncDesc(pointer); }
        }

        private static object ReadVariable(ITypeInfo typeInfo, int index, int memberIndex)
        {
            IntPtr pointer = IntPtr.Zero;
            try
            {
                typeInfo.GetVarDesc(index, out pointer);
                var descriptor = (VARDESC)Marshal.PtrToStructure(pointer, typeof(VARDESC));
                string name, description, helpFile;
                int helpContext;
                typeInfo.GetDocumentation(descriptor.memid, out name, out description,
                    out helpContext, out helpFile);
                return new { MemberIndex = memberIndex, Kind = "Variable", Name = name,
                    DispId = descriptor.memid, VariableKind = descriptor.varkind.ToString() };
            }
            finally { if (pointer != IntPtr.Zero) typeInfo.ReleaseVarDesc(pointer); }
        }

        private static string Hash(string value)
        {
            using (var sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(value)))
                    .Replace("-", "").ToLowerInvariant();
        }

        private static object Error(Exception error)
        {
            return new { Type = error.GetType().Name,
                HResult = HResultHex(error),
                Message = error.Message };
        }

        private static string HResultHex(Exception error)
        {
            return "0x" + unchecked((uint)error.HResult).ToString("X8");
        }

        private sealed class ReferenceSource
        {
            public string Name { get; set; }
            public string Guid { get; set; }
            public int Major { get; set; }
            public int Minor { get; set; }
            public string FullPath { get; set; }
            public long FileLength { get; set; }
            public string FileLastWriteUtc { get; set; }
        }

        private sealed class LoadedLibrary
        {
            public ITypeLib Library { get; set; }
            public string Source { get; set; }
            public object FallbackError { get; set; }
        }

        private sealed class TypeMetadata
        {
            public int TypeIndex { get; set; }
            public string Name { get; set; }
            public string Guid { get; set; }
            public string Kind { get; set; }
            public int FunctionCount { get; set; }
            public int VariableCount { get; set; }
            public int ImplementedInterfaceCount { get; set; }
            public string Identity { get; set; }
        }
    }
}
